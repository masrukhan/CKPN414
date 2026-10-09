using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Globalization;
using System.IO;
using System.Web.Script.Serialization;
using CKPNLibrary.Data;
using CKPNLibrary.Helpers;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Panel
{
    /// <summary>
    /// Data profil kontrak dari file template APOLLO OJK (Master!D14) untuk dokumen CKPN — Tahap 6b.
    ///
    /// Sumber per nomor rekening (kolom J sheet KC):
    ///   KC0600 Murabahah · KC0700 Istishna · KC0800 Multijasa · KC0900 Qardh · KC1000 Bagi hasil · KC1100 Ijarah
    ///     identitas (CIF, nama, NIK), klasifikasi, jangka waktu, jenis penggunaan, sektor ekonomi, akad,
    ///     nilai kontrak, imbalan, kualitas, saldo, tunggakan, agunan (satu baris per agunan), CKPN template,
    ///     tanggal mulai macet, tanggal akad awal/akhir
    ///   GB0500 restrukturisasi (cara, frekuensi, kondisi sebelum/sesudah)
    ///   KC2900 hapus buku (tanggal, jumlah, dipulihkan, baki debet)
    ///
    /// Kolom dipetakan TETAP per sheet sesuai format template APOLLO (judul kolom KC1000 untuk agunan tidak
    /// konsisten dengan isinya, sehingga pemetaan berdasarkan judul tidak dapat diandalkan).
    /// Disimpan per periode di tabel profil_template (skema v10) agar dokumen tidak perlu membuka template lagi.
    /// Alamat debitur TIDAK ada di template APOLLO → tetap diisi manual di memo.
    /// </summary>
    internal static class ProfilTemplate
    {
        private static readonly JavaScriptSerializer _json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

        private class Peta
        {
            public string Akad;
            public string[,] Kolom;      // { kunci, kolom }
            public string[,] Agunan;     // { kunci, kolom }
            public string KolomNilaiDipakai;   // kolom nilai agunan yang dipakai perhitungan CKPN Individu (SheetSpec.ColJaminan)
        }

        private static readonly string[,] AgunanStd =
        {
            { "jenis", "AJ" }, { "pengikatan", "AK" }, { "nomor", "AL" }, { "karat", "AM" }, { "berat", "AN" }, { "lat", "AO" }, { "lon", "AP" },
            { "penjamin", "AQ" }, { "tglNilai", "AR" }, { "nilai", "AS" }, { "diperhitungkan", "AT" }, { "bagianDijamin", "AU" }
        };

        private static readonly Dictionary<string, Peta> PetaSheet = new Dictionary<string, Peta>(StringComparer.OrdinalIgnoreCase)
        {
            { "KC0600", Std("Murabahah") }, { "KC0700", Std("Istishna") }, { "KC0800", Std("Multijasa") },
            { "KC0900", new Peta
                {
                    Akad = "Qardh", KolomNilaiDipakai = "AQ",
                    Kolom = Gabung(Identitas(), new[,] {
                        { "nilaiKontrak", "U" }, { "periodeAngsuran", "V" }, { "imbalanAwal", "W" }, { "imbalanKini", "X" }, { "kualitas", "Y" },
                        { "nominal", "AA" }, { "jumlah", "AC" }, { "hari", "AD" }, { "tunggakanPokok", "AE" }, { "tunggakanMargin", "AF" },
                        { "ckpnTemplate", "AT" }, { "jenisCkpn", "AU" }, { "tglMacet", "AV" }, { "program", "AW" }, { "akadAwal", "AY" }, { "akadAkhir", "AZ" } }),
                    Agunan = new[,] {
                        { "jenis", "AH" }, { "pengikatan", "AI" }, { "nomor", "AJ" }, { "karat", "AK" }, { "berat", "AL" }, { "lat", "AM" }, { "lon", "AN" },
                        { "penjamin", "AO" }, { "tglNilai", "AP" }, { "nilai", "AQ" }, { "diperhitungkan", "AR" }, { "bagianDijamin", "AS" } }
                } },
            { "KC1000", new Peta
                {
                    Akad = "Bagi hasil", KolomNilaiDipakai = "AY",
                    Kolom = Gabung(Identitas(), new[,] {
                        { "jenisAkad", "U" }, { "sifatInvestasi", "V" }, { "metodeBagiHasil", "W" }, { "nisbah", "X" }, { "nilaiKontrak", "Y" },
                        { "periodeAngsuran", "Z" }, { "imbalanAwal", "AB" }, { "imbalanKini", "AC" }, { "kualitas", "AE" }, { "nominal", "AG" },
                        { "jumlah", "AI" }, { "hari", "AK" }, { "tunggakanPokok", "AL" }, { "hariBagiHasil", "AM" }, { "tunggakanMargin", "AN" },
                        { "ckpnTemplate", "BB" }, { "jenisCkpn", "BC" }, { "tglMacet", "BD" }, { "program", "BE" }, { "akadAwal", "BG" }, { "akadAkhir", "BH" } }),
                    Agunan = new[,] {
                        { "jenis", "AP" }, { "pengikatan", "AQ" }, { "nomor", "AR" }, { "karat", "AS" }, { "berat", "AT" }, { "lat", "AU" }, { "lon", "AV" },
                        { "penjamin", "AW" }, { "tglNilai", "AX" }, { "nilai", "AY" }, { "diperhitungkan", "AZ" }, { "bagianDijamin", "BA" } }
                } },
            { "KC1100", new Peta
                {
                    Akad = "Ijarah", KolomNilaiDipakai = "AY",
                    Kolom = Gabung(Identitas(), new[,] {
                        { "jenisAkad", "U" }, { "jenisAset", "V" }, { "waktuPerolehan", "W" }, { "hargaPerolehan", "X" }, { "periodeAngsuran", "AA" },
                        { "nilaiKontrak", "AB" }, { "sewaPerPeriode", "AC" }, { "imbalanAwal", "AD" }, { "imbalanKini", "AE" }, { "kualitas", "AF" },
                        { "akumulasiPenyusutan", "AI" }, { "hari", "AK" }, { "tunggakanPokok", "AL" }, { "tunggakanMargin", "AM" },
                        { "ckpnTemplate", "BA" }, { "jenisCkpn", "BB" }, { "tglMacet", "BC" }, { "program", "BD" }, { "akadAwal", "BF" }, { "akadAkhir", "BG" } }),
                    Agunan = new[,] {
                        { "jenis", "AO" }, { "pengikatan", "AP" }, { "nomor", "AQ" }, { "karat", "AR" }, { "berat", "AS" }, { "lat", "AT" }, { "lon", "AU" },
                        { "penjamin", "AV" }, { "tglNilai", "AW" }, { "nilai", "AX" }, { "diperhitungkan", "AY" }, { "bagianDijamin", "AZ" } }
                } }
        };

        private static string[,] Identitas()
        {
            return new[,] {
                { "sandiKantor", "B" }, { "cif", "C" }, { "nama", "D" }, { "nik", "E" }, { "kelompok", "F" }, { "golongan", "G" }, { "hubungan", "H" },
                { "kategoriUsaha", "I" }, { "mulai", "K" }, { "jatuhTempo", "L" }, { "sumberDana", "M" }, { "lokasi", "O" }, { "jenisPiutang", "P" },
                { "sifat", "Q" }, { "status", "R" }, { "jenisPenggunaan", "S" }, { "sektor", "T" } };
        }

        private static Peta Std(string akad)
        {
            return new Peta
            {
                Akad = akad, KolomNilaiDipakai = "AS", Agunan = AgunanStd,
                Kolom = Gabung(Identitas(), new[,] {
                    { "nilaiKontrak", "U" }, { "periodeAngsuran", "V" }, { "imbalanAwal", "W" }, { "imbalanKini", "X" }, { "kualitas", "Y" },
                    { "nominal", "AA" }, { "saldoPokok", "AC" }, { "marginDitangguhkan", "AD" }, { "jumlah", "AE" }, { "hari", "AF" },
                    { "tunggakanPokok", "AG" }, { "tunggakanMargin", "AH" }, { "ckpnTemplate", "AV" }, { "jenisCkpn", "AW" }, { "tglMacet", "AX" },
                    { "program", "AY" }, { "akadAwal", "BA" }, { "akadAkhir", "BB" } })
            };
        }

        private static string[,] Gabung(string[,] a, string[,] b)
        {
            var h = new string[a.GetLength(0) + b.GetLength(0), 2];
            for (int i = 0; i < a.GetLength(0); i++) { h[i, 0] = a[i, 0]; h[i, 1] = a[i, 1]; }
            for (int i = 0; i < b.GetLength(0); i++) { h[a.GetLength(0) + i, 0] = b[i, 0]; h[a.GetLength(0) + i, 1] = b[i, 1]; }
            return h;
        }

        private static readonly HashSet<string> KunciTanggal = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        { "mulai", "jatuhTempo", "tglMacet", "akadAwal", "akadAkhir", "tglNilai", "waktuPerolehan" };

        private const int BarisDataPertama = 5;   // header baris 2–4 (SheetSpec.HeaderRow = 4)

        // =================================================================
        // Baca dari file template
        // =================================================================
        /// <summary>Baca profil untuk rekening tertentu. Template dibuka read-only (atau dipakai bila sudah terbuka).</summary>
        public static Dictionary<string, Dictionary<string, object>> Baca(Excel.Application app, string path, ICollection<string> rekening)
        {
            var dicari = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var r in rekening) { string k = Norm(r); if (k.Length > 0) dicari.Add(k); }
            var hasil = new Dictionary<string, Dictionary<string, object>>(StringComparer.OrdinalIgnoreCase);
            if (dicari.Count == 0) return hasil;
            if (!File.Exists(path)) throw new FileNotFoundException("File template tidak ditemukan: " + path);

            Excel.Workbook wb = null;
            bool dibukaDiSini = false;
            bool layar = true, alert = true;
            try { layar = app.ScreenUpdating; alert = app.DisplayAlerts; app.ScreenUpdating = false; app.DisplayAlerts = false; } catch { }
            try
            {
                string full = Path.GetFullPath(path);
                foreach (Excel.Workbook w in app.Workbooks)
                    if (string.Equals(w.FullName, full, StringComparison.OrdinalIgnoreCase)) { wb = w; break; }
                if (wb == null) { wb = app.Workbooks.Open(path, UpdateLinks: 0, ReadOnly: true); dibukaDiSini = true; }

                foreach (var kv in PetaSheet)
                {
                    Excel.Worksheet ws = ParameterMaster.CariSheet(wb, kv.Key);
                    if (ws == null) continue;
                    object[,] data = BacaArea(ws);
                    if (data == null) continue;
                    BacaSheetKC(kv.Key, kv.Value, data, dicari, hasil);
                }

                Excel.Worksheet gb = ParameterMaster.CariSheet(wb, "GB0500");
                if (gb != null) BacaRestrukturisasi(BacaArea(gb), dicari, hasil);
                Excel.Worksheet hb = ParameterMaster.CariSheet(wb, "KC2900");
                if (hb != null) BacaHapusBuku(BacaArea(hb), dicari, hasil);
            }
            finally
            {
                if (dibukaDiSini && wb != null) try { wb.Close(false); } catch { }
                try { app.DisplayAlerts = alert; app.ScreenUpdating = layar; } catch { }
            }
            return hasil;
        }

        private static object[,] BacaArea(Excel.Worksheet ws)
        {
            Excel.Range used = ws.UsedRange;
            int baris = used.Row + used.Rows.Count - 1, kolom = used.Column + used.Columns.Count - 1;
            if (baris < 2 || kolom < 2) return null;
            object v = ((Excel.Range)ws.Range[ws.Cells[1, 1], ws.Cells[baris, kolom]]).Value2;
            return v as object[,];
        }

        private static object Sel(object[,] d, int baris, string kolom)
        {
            int c = IndeksKolom(kolom);
            if (baris > d.GetUpperBound(0) || c > d.GetUpperBound(1)) return null;
            return d[baris, c];
        }

        private static void BacaSheetKC(string sheet, Peta p, object[,] d, HashSet<string> dicari,
                                        Dictionary<string, Dictionary<string, object>> hasil)
        {
            int akhir = d.GetUpperBound(0);
            for (int r = BarisDataPertama; r <= akhir; r++)
            {
                string rek = Norm(Sel(d, r, "J"));
                if (rek.Length == 0 || !dicari.Contains(rek)) continue;

                Dictionary<string, object> h;
                if (!hasil.TryGetValue(rek, out h))
                {
                    h = new Dictionary<string, object> { { "sheet", sheet }, { "akad", p.Akad }, { "rek", rek }, { "agunan", new List<object>() } };
                    hasil[rek] = h;
                }
                // Baris pertama (bernama) mengisi data kontrak; baris berikutnya rekening yang sama = agunan tambahan
                bool barisUtama = !h.ContainsKey("nama") || string.IsNullOrEmpty(Convert.ToString(h["nama"]));
                if (barisUtama)
                    for (int i = 0; i < p.Kolom.GetLength(0); i++)
                    {
                        object v = Nilai(p.Kolom[i, 0], Sel(d, r, p.Kolom[i, 1]));
                        if (v != null) h[p.Kolom[i, 0]] = v;
                    }

                var ag = new Dictionary<string, object>();
                for (int i = 0; i < p.Agunan.GetLength(0); i++)
                {
                    object v = Nilai(p.Agunan[i, 0], Sel(d, r, p.Agunan[i, 1]));
                    if (v != null) ag[p.Agunan[i, 0]] = v;
                }
                object nilaiPakai = Sel(d, r, p.KolomNilaiDipakai);
                ag["nilaiDipakai"] = nilaiPakai is double ? nilaiPakai : (object)0.0;
                if (ag.ContainsKey("jenis") || ag.ContainsKey("nomor") || Angka(ag, "nilai") > 0)
                    ((List<object>)h["agunan"]).Add(ag);
            }
        }

        // GB0500: B nama, C nomor nasabah, D rekening, G cara, H frekuensi, I–M sebelum, N–R sesudah
        private static void BacaRestrukturisasi(object[,] d, HashSet<string> dicari, Dictionary<string, Dictionary<string, object>> hasil)
        {
            if (d == null) return;
            for (int r = BarisDataPertama; r <= d.GetUpperBound(0); r++)
            {
                string rek = Norm(Sel(d, r, "D"));
                if (rek.Length == 0 || !dicari.Contains(rek)) continue;
                Dictionary<string, object> h;
                if (!hasil.TryGetValue(rek, out h)) { h = new Dictionary<string, object> { { "rek", rek }, { "agunan", new List<object>() } }; hasil[rek] = h; }
                h["restruktur"] = new Dictionary<string, object>
                {
                    { "cara", Teks(Sel(d, r, "G")) }, { "frekuensi", Sel(d, r, "H") },
                    { "sebelum", new Dictionary<string, object> {
                        { "akad", Teks(Sel(d, r, "I")) }, { "sisa", Sel(d, r, "J") }, { "awal", Tanggal(Sel(d, r, "K")) },
                        { "akhir", Tanggal(Sel(d, r, "L")) }, { "kualitas", Teks(Sel(d, r, "M")) } } },
                    { "sesudah", new Dictionary<string, object> {
                        { "akad", Teks(Sel(d, r, "N")) }, { "sisa", Sel(d, r, "O") }, { "awal", Tanggal(Sel(d, r, "P")) },
                        { "akhir", Tanggal(Sel(d, r, "Q")) }, { "kualitas", Teks(Sel(d, r, "R")) } } }
                };
            }
        }

        // KC2900 (baris data mulai 3): C CIF, H rekening, I tanggal hapus buku, J jumlah, K dipulihkan, L baki debet, M kode KC
        private static void BacaHapusBuku(object[,] d, HashSet<string> dicari, Dictionary<string, Dictionary<string, object>> hasil)
        {
            if (d == null) return;
            for (int r = 3; r <= d.GetUpperBound(0); r++)
            {
                string rek = Norm(Sel(d, r, "H"));
                if (rek.Length == 0 || !dicari.Contains(rek)) continue;
                Dictionary<string, object> h;
                if (!hasil.TryGetValue(rek, out h)) { h = new Dictionary<string, object> { { "rek", rek }, { "agunan", new List<object>() } }; hasil[rek] = h; }
                if (!h.ContainsKey("cif")) h["cif"] = Teks(Sel(d, r, "C"));
                h["hapusBuku"] = new Dictionary<string, object>
                {
                    { "tanggal", Tanggal(Sel(d, r, "I")) }, { "jumlah", Sel(d, r, "J") }, { "dipulihkan", Sel(d, r, "K") },
                    { "bakiDebet", Sel(d, r, "L") }, { "kodeKC", Teks(Sel(d, r, "M")) }
                };
            }
        }

        // =================================================================
        // Simpan / baca database
        // =================================================================
        /// <summary>
        /// Ambil profil rekening untuk satu periode dari file template lalu simpan. Bila <paramref name="hanyaYangBelum"/>,
        /// rekening yang sudah tersimpan dari file yang sama (dan file tidak berubah) dilewati.
        /// </summary>
        public static Dictionary<string, object> AmbilDanSimpan(Excel.Application app, string tanggal, string path,
                                                               ICollection<string> rekening, bool hanyaYangBelum)
        {
            string info;
            if (!Database.BolehMenulis(out info)) throw new InvalidOperationException(info);
            string cap = PenjelasanCKPN.CapTemplate(path) ?? "";
            var perlu = new List<string>();
            var sudah = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (hanyaYangBelum && Database.Ada)
            {
                using (var con = Database.Buka(false))
                {
                    object capLama = Database.Scalar(con, "SELECT cap FROM profil_template_info WHERE tanggal=@p0 AND file=@p1", tanggal, Path.GetFullPath(path));
                    if (capLama != null && Convert.ToString(capLama) == cap)
                        using (var cmd = Database.Cmd(con, "SELECT no_rek FROM profil_template WHERE tanggal=@p0", tanggal))
                        using (var rd = cmd.ExecuteReader())
                            while (rd.Read()) sudah.Add(rd.GetString(0));
                }
            }
            foreach (var r in rekening) { string k = Norm(r); if (k.Length > 0 && !sudah.Contains(k) && !perlu.Contains(k)) perlu.Add(k); }
            if (perlu.Count == 0)
                return new Dictionary<string, object> { { "dibaca", 0 }, { "ditemukan", 0 }, { "tidakDitemukan", new List<string>() }, { "dilewati", true } };

            var hasil = Baca(app, path, perlu);
            var tidak = new List<string>();
            string now = Database.Sekarang();
            using (var con = Database.Buka(true))
            using (var tx = con.BeginTransaction())
            {
                foreach (var rek in perlu)
                {
                    Dictionary<string, object> h;
                    if (!hasil.TryGetValue(rek, out h)) { tidak.Add(rek); continue; }
                    Database.Exec(con, "INSERT OR REPLACE INTO profil_template(tanggal,no_rek,data,waktu) VALUES(@p0,@p1,@p2,@p3)",
                        tanggal, rek, _json.Serialize(h), now);
                }
                Database.Exec(con,
                    "INSERT OR REPLACE INTO profil_template_info(tanggal,file,cap,pengguna,waktu) VALUES(@p0,@p1,@p2,@p3,@p4)",
                    tanggal, Path.GetFullPath(path), cap, Environment.UserName, now);
                tx.Commit();
            }
            LogProses.Catat("Profil template", tanggal, tidak.Count == 0 ? LogProses.OK : LogProses.Peringatan, LogProses.R()
                .Tambah("File", Path.GetFileName(path)).Tambah("Rekening dibaca", perlu.Count).Tambah("Ditemukan", hasil.Count)
                .TambahBila(tidak.Count > 0, "Tidak ditemukan", string.Join(", ", tidak.ToArray())));
            return new Dictionary<string, object>
            {
                { "dibaca", perlu.Count }, { "ditemukan", hasil.Count }, { "tidakDitemukan", tidak }, { "file", Path.GetFileName(path) }
            };
        }

        /// <summary>Rekening Individu + LGD CS dari versi aktif satu periode (untuk "ambil ulang" dari tab Dokumen).</summary>
        public static List<string> RekeningPeriode(string tanggal)
        {
            var list = new List<string>();
            if (!Database.Ada) return list;
            using (var con = Database.Buka(false))
            {
                foreach (var sql in new[] {
                    "SELECT DISTINCT h.no_kontrak FROM hasil_individu h JOIN run_grup g ON g.id=h.run_id JOIN periode p ON p.id=g.periode_id WHERE p.tanggal=@p0 AND g.aktif=1 AND g.dihapus=0",
                    "SELECT DISTINCT c.no_rek FROM hasil_lgdcs c JOIN run_grup g ON g.id=c.run_id JOIN periode p ON p.id=g.periode_id WHERE p.tanggal=@p0 AND g.aktif=1 AND g.dihapus=0" })
                    using (var cmd = Database.Cmd(con, sql, tanggal))
                    using (var rd = cmd.ExecuteReader())
                        while (rd.Read()) if (!(rd[0] is DBNull)) list.Add(rd.GetString(0));
            }
            return list;
        }

        /// <summary>Profil tersimpan satu periode → { rek: data } + info file.</summary>
        public static void BacaTersimpan(SQLiteConnection con, string tanggal, Dictionary<string, object> keluar)
        {
            var profil = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            using (var cmd = Database.Cmd(con, "SELECT no_rek, data FROM profil_template WHERE tanggal=@p0", tanggal))
            using (var rd = cmd.ExecuteReader())
                while (rd.Read())
                {
                    try { profil[rd.GetString(0)] = _json.DeserializeObject(rd.GetString(1)); } catch { }
                }
            keluar["profil"] = profil;
            using (var cmd = Database.Cmd(con, "SELECT file, pengguna, waktu FROM profil_template_info WHERE tanggal=@p0", tanggal))
            using (var rd = cmd.ExecuteReader())
                keluar["profilInfo"] = rd.Read()
                    ? new Dictionary<string, object> { { "file", Path.GetFileName(rd.GetString(0)) }, { "pengguna", rd[1] as string }, { "waktu", rd[2] as string } }
                    : null;
        }

        // =================================================================
        // Referensi kode APOLLO (config\kode_apollo.txt) — diisi/diubah sendiri oleh bank
        //   format per baris:  kategori;kode;keterangan      (baris diawali # = komentar)
        // =================================================================
        public static string FileKode { get { return Path.Combine(AppPaths.FolderConfig, "kode_apollo.txt"); } }

        public static Dictionary<string, object> BacaKode()
        {
            var h = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            try
            {
                if (!File.Exists(FileKode)) BuatFileKodeAwal();
                foreach (var baris in File.ReadAllLines(FileKode))
                {
                    string s = baris.Trim();
                    if (s.Length == 0 || s.StartsWith("#")) continue;
                    var p = s.Split(new[] { ';' }, 3);
                    if (p.Length < 3) continue;
                    object kat;
                    if (!h.TryGetValue(p[0].Trim(), out kat)) { kat = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase); h[p[0].Trim()] = kat; }
                    ((Dictionary<string, object>)kat)[p[1].Trim()] = p[2].Trim();
                }
            }
            catch (Exception ex) { CatatanLog.Tulis("Baca kode_apollo.txt gagal: " + ex.Message); }
            return h;
        }

        private static void BuatFileKodeAwal()
        {
            try
            {
                string info;
                if (!Database.BolehMenulis(out info)) return;
                Directory.CreateDirectory(AppPaths.FolderConfig);
                File.WriteAllLines(FileKode, new[]
                {
                    "# Keterangan kode template APOLLO untuk dokumen CKPN (Panel CKPN Tahap 6b).",
                    "# Format: kategori;kode;keterangan — satu kode per baris. Lengkapi sesuai pedoman pelaporan APOLLO/SLIK yang berlaku.",
                    "# Kategori: kualitas, jenisPenggunaan, kategoriUsaha, golongan, sektor, jenisAgunan, pengikatan, periodeAngsuran,",
                    "#           sumberDana, jenisAkad, caraRestrukturisasi, penjamin, lokasi",
                    "# Kode yang belum ada di file ini tetap tampil sebagai kode.",
                    "kualitas;1;Lancar",
                    "kualitas;2;Dalam Perhatian Khusus",
                    "kualitas;3;Kurang Lancar",
                    "kualitas;4;Diragukan",
                    "kualitas;5;Macet",
                    "jenisPenggunaan;1;Modal kerja",
                    "jenisPenggunaan;2;Investasi",
                    "jenisPenggunaan;3;Konsumsi"
                });
            }
            catch (Exception ex) { CatatanLog.Tulis("Buat kode_apollo.txt gagal: " + ex.Message); }
        }

        // =================================================================
        // Util
        // =================================================================
        internal static string Norm(object v)
        {
            if (v == null) return "";
            if (v is double) return ((double)v).ToString("0", CultureInfo.InvariantCulture);
            return Convert.ToString(v).Trim().TrimStart('\'');
        }

        private static object Nilai(string kunci, object v)
        {
            if (v == null) return null;
            if (KunciTanggal.Contains(kunci)) return Tanggal(v);
            if (v is double) return v;
            string s = Convert.ToString(v).Trim();
            return s.Length == 0 ? null : s;
        }

        private static string Teks(object v) { return v == null ? "" : Norm(v); }

        /// <summary>"16-10-2012" / tanggal Excel → "2012-10-16"; teks lain dikembalikan apa adanya.</summary>
        internal static string Tanggal(object v)
        {
            if (v == null) return null;
            if (v is double)
            {
                double d = (double)v;
                if (d > 1 && d < 2958465) return DateTime.FromOADate(d).ToString("yyyy-MM-dd");
                return null;
            }
            string s = Convert.ToString(v).Trim();
            if (s.Length == 0) return null;
            DateTime t;
            if (DateTime.TryParseExact(s, new[] { "dd-MM-yyyy", "d-M-yyyy", "dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd", "yyyyMMdd" },
                                       CultureInfo.InvariantCulture, DateTimeStyles.None, out t))
                return t.ToString("yyyy-MM-dd");
            return s;
        }

        private static double Angka(Dictionary<string, object> d, string k)
        {
            object v;
            return d.TryGetValue(k, out v) && v is double ? (double)v : 0;
        }

        private static int IndeksKolom(string huruf)
        {
            int n = 0;
            foreach (char c in huruf.ToUpperInvariant()) n = n * 26 + (c - 'A' + 1);
            return n;
        }
    }
}
