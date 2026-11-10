using System;
using System.Collections.Generic;
using System.IO;
using CKPNLibrary.Helpers;
using CKPNLibrary.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Panel
{
    /// <summary>
    /// Rekomendasi susunan grup & Top-N (⚙ Pengaturan › Susunan grup).
    ///
    /// Membaca template APOLLO periode Master (Master!D14) dan mengembalikan, per unit penyusun grup
    /// (KC0600–KC1100 pada mode akad; sandi kantor / jenis penggunaan / kelompok sektor pada mode lain):
    ///   - daftar CIF (sebagai indeks bilangan bulat agar JSON ringkas) dan baki debet CIF di unit itu,
    ///   - jumlah rekening.
    /// Panel menggabungkan unit menjadi grup (gabungan CIF dihitung unik, sama dengan CKPN Individu yang
    /// menjumlah semua kontrak satu CIF), lalu menyarankan pengelompokan dan Top-N ≤ 30% CIF grup.
    ///
    /// Baki debet per kontrak memakai <see cref="ExcelHelper.BacaSheetKC"/> — SAMA PERSIS dengan dasar
    /// ranking CKPN Individu (KC0600–KC0900 OS; KC1000/KC1100 tunggakan sesuai ruang lingkup). Karena itu
    /// "CIF berbaki" (baki &gt; 0) bisa jauh lebih kecil dari jumlah CIF untuk ijarah / bagi hasil.
    ///
    /// Unit per kontrak pada mode selain akad mengikuti aturan saringan <see cref="SumberLingkup"/>:
    /// diambil dari baris pembiayaan (nama terisi) pertama kontrak; mode unit tetap (penggunaan/sektor)
    /// yang isiannya kosong masuk unit cadangan; sandi kantor kosong tidak masuk grup mana pun.
    ///
    /// Hanya membaca: template dibuka read-only bila belum terbuka, tidak ada yang ditulis ke database.
    /// </summary>
    internal static class RekomendasiSusunan
    {
        private const int BarisData = 5;   // header template baris 2–4 (SheetSpec.HeaderRow = 4)

        private class Unit
        {
            public string Kode;
            public int Rekening;
            public Dictionary<int, double> Cif = new Dictionary<int, double>();   // indeks CIF → baki debet
        }

        public static Dictionary<string, object> Pindai(Excel.Application app, string path, string tanggal)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
                throw new FileNotFoundException("File template APOLLO (Master!D14) tidak ditemukan: " + path);

            ModeHitung.Mode mode = ModeHitung.Aktif;
            var indeksCif = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var unit = new Dictionary<string, Unit>(StringComparer.OrdinalIgnoreCase);
            var urutan = new List<string>();
            int tanpaUnit = 0;

            Excel.Workbook wb = null;
            bool dibukaDiSini = false, layar = true, alert = true;
            try { layar = app.ScreenUpdating; alert = app.DisplayAlerts; app.ScreenUpdating = false; app.DisplayAlerts = false; } catch { }
            try
            {
                string full = Path.GetFullPath(path);
                foreach (Excel.Workbook w in app.Workbooks)
                    if (string.Equals(w.FullName, full, StringComparison.OrdinalIgnoreCase)) { wb = w; break; }
                if (wb == null) { wb = app.Workbooks.Open(path, UpdateLinks: 0, ReadOnly: true); dibukaDiSini = true; }

                foreach (string kc in ModeHitung.SemuaKC)
                {
                    SheetSpec spec;
                    try { spec = SheetSpec.Buat(kc); } catch { continue; }
                    // dasar baki debet = CKPN Individu (termasuk ruang lingkup tunggakan KC1000/KC1100)
                    Dictionary<string, CIFData> dict = ExcelHelper.BacaSheetKC(wb, spec, null);
                    if (dict.Count == 0) continue;

                    Dictionary<string, string> unitKontrak = mode.Utama ? null : BacaUnitKontrak(wb, kc, spec, mode, ref tanpaUnit);

                    foreach (CIFData c in dict.Values)
                    {
                        int ix;
                        if (!indeksCif.TryGetValue(c.CIF, out ix)) { ix = indeksCif.Count; indeksCif[c.CIF] = ix; }
                        foreach (KontrakData k in c.Kontrak.Values)
                        {
                            string kode;
                            if (mode.Utama) kode = kc;
                            else if (!unitKontrak.TryGetValue(k.NoKontrak ?? "", out kode) || string.IsNullOrEmpty(kode)) continue;

                            Unit u;
                            if (!unit.TryGetValue(kode, out u)) { u = new Unit { Kode = kode }; unit[kode] = u; urutan.Add(kode); }
                            u.Rekening++;
                            double b;
                            u.Cif.TryGetValue(ix, out b);
                            u.Cif[ix] = b + k.OS;
                        }
                    }
                }
            }
            finally
            {
                if (dibukaDiSini && wb != null) try { wb.Close(false); } catch { }
                try { app.DisplayAlerts = alert; app.ScreenUpdating = layar; } catch { }
            }

            var daftar = new List<object>();
            foreach (string kode in urutan)
            {
                Unit u = unit[kode];
                var cif = new List<int>(u.Cif.Count);
                var baki = new List<double>(u.Cif.Count);
                int berbaki = 0;
                foreach (var kv in u.Cif)
                {
                    cif.Add(kv.Key);
                    baki.Add(Math.Round(kv.Value, 0));
                    if (kv.Value > 0) berbaki++;
                }
                daftar.Add(new Dictionary<string, object>
                {
                    { "kode", kode }, { "rekening", u.Rekening }, { "jumlahCif", cif.Count }, { "berbaki", berbaki },
                    { "cif", cif }, { "baki", baki }
                });
            }

            return new Dictionary<string, object>
            {
                { "mode", mode.Id }, { "file", Path.GetFileName(path) }, { "tanggal", tanggal ?? "" },
                { "jumlahCif", indeksCif.Count }, { "tanpaUnit", tanpaUnit }, { "unit", daftar }
            };
        }

        /// <summary>
        /// Nomor kontrak → kode unit (sandi kantor / jenis penggunaan / kelompok sektor) untuk satu sheet KC.
        /// Kunci nomor kontrak dibaca dengan <see cref="ExcelHelper.BacaKolomString"/> agar sama dengan BacaSheetKC.
        /// </summary>
        private static Dictionary<string, string> BacaUnitKontrak(Excel.Workbook wb, string kc, SheetSpec spec,
                                                                  ModeHitung.Mode mode, ref int tanpaUnit)
        {
            var peta = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            Excel.Worksheet ws = ParameterMaster.CariSheet(wb, kc);
            if (ws == null) return peta;
            int akhir = Math.Max(ExcelHelper.CariLastRow(ws, spec.ColCIF, spec.HeaderRow),
                                 ExcelHelper.CariLastRow(ws, spec.ColNoRek, spec.HeaderRow));
            if (akhir < BarisData) return peta;

            string[] rek = ExcelHelper.BacaKolomString(ws, spec.ColNoRek, BarisData, akhir);
            string[] nama = ExcelHelper.BacaKolomString(ws, spec.ColNama, BarisData, akhir);
            string[] isi = ExcelHelper.BacaKolomString(ws, mode.KolomFilter, BarisData, akhir);
            bool perKontrak = ModeHitung.UnitPerKontrak(mode);
            string cadangan = ModeHitung.UnitCadangan(mode);

            for (int i = 0; i < rek.Length; i++)
            {
                string r = (rek[i] ?? "").Trim();
                if (r.Length == 0 || peta.ContainsKey(r)) continue;
                // unit diambil dari baris pembiayaan (nama terisi) pertama kontrak, sama dengan SumberLingkup
                if (string.IsNullOrEmpty((nama[i] ?? "").Trim())) continue;
                string kode = ModeHitung.NormUnit(mode, i < isi.Length ? isi[i] : null);
                if (kode.Length == 0)
                {
                    if (perKontrak) kode = cadangan;
                    else { tanpaUnit++; continue; }   // sandi kantor kosong → tidak masuk grup mana pun
                }
                peta[r] = kode;
            }
            return peta;
        }
    }
}
