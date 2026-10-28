using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;
using CKPNLibrary.Helpers;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Panel
{
    /// <summary>
    /// Tahap 8 — salinan file sumber yang sudah disaring per lingkup grup (mode per kantor).
    ///
    /// Modul perhitungan (Individu, Net Flow, Migration, LGD CS, Refresh Summary) TIDAK diubah: pipeline panel
    /// menyerahkan path salinan tersaring sebagai pengganti file asli. Pada salinan, baris sheet KC0600–KC1100
    /// yang sandi kantornya (kolom B) di luar grup dihapus; sheet lain (KC0500, KC2900, GB…) dibiarkan utuh
    /// sehingga ABA, data keuangan, dan daftar restrukturisasi tetap tingkat bank. LGD ER tidak memakai salinan —
    /// saringannya ada di referensi KC (ReferensiKC.Resolver).
    ///
    /// Aturan khusus PD Migration ("kantor posisi awal"): file akhir triwulan memuat baris sandi grup DAN seluruh
    /// rekening yang ada di file awal grup ini, sehingga rekening yang pindah kantor tetap dilacak dari kantor asalnya.
    ///
    /// File asli tidak pernah dibuka untuk ditulis: berkasnya disalin dulu ke %LOCALAPPDATA%\CKPN414\lingkup\,
    /// lalu salinan itu yang disaring dan disimpan. Salinan dipakai ulang selama file asli tidak berubah;
    /// salinan berumur lebih dari 3 hari dihapus otomatis.
    /// </summary>
    internal static class SumberLingkup
    {
        private const int BarisData = 5;          // header template baris 2–4 (SheetSpec.HeaderRow = 4)
        private const string KolomRekening = "J";
        private const string KolomNama = "D";
        private static readonly JavaScriptSerializer _json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

        public static string Folder { get { return Path.Combine(AppPaths.FolderLokal, "lingkup"); } }

        internal class Hasil
        {
            public string Path;
            public HashSet<string> Rekening = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            public int BarisAsal, BarisTersisa;
            /// <summary>Tahap 8c: baris pembiayaan tanpa jenis penggunaan (→ Konsumsi) dan kontrak dengan isian berbeda antarbaris.</summary>
            public int Cadangan, Berbeda;
            public bool DariCache;
        }

        /// <summary>
        /// Salinan tersaring satu file. rekTambahan: rekening (ternormalisasi) yang tetap disertakan walau sandinya
        /// di luar grup (file akhir PD Migration). Mengembalikan path salinan dan rekening yang tersisa.
        /// </summary>
        public static Hasil Salin(Excel.Application app, string path, FilterLingkup f, ICollection<string> rekTambahan = null)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path)) throw new FileNotFoundException("File sumber tidak ditemukan: " + path);
            Directory.CreateDirectory(Folder);

            string full = System.IO.Path.GetFullPath(path);
            var fi = new FileInfo(full);
            string tambahanKunci = "";
            if (rekTambahan != null && rekTambahan.Count > 0)
            {
                var l = new List<string>(rekTambahan);
                l.Sort(StringComparer.OrdinalIgnoreCase);
                tambahanKunci = Hash(string.Join(",", l.ToArray()));
            }
            string kunci = full + "|" + fi.Length + "@" + fi.LastWriteTimeUtc.Ticks + "|" + f.ModeId + ":" + f.Kode + "|" + tambahanKunci;
            string nama = System.IO.Path.GetFileNameWithoutExtension(full) + "__" + f.ModeId + "_" + f.Kode.Replace(",", "-") +
                          (tambahanKunci.Length > 0 ? "_m" + tambahanKunci.Substring(0, 8) : "") + "_" + Hash(full).Substring(0, 6) +
                          System.IO.Path.GetExtension(full);
            string tujuan = System.IO.Path.Combine(Folder, nama);
            string sidecar = tujuan + ".json";

            // ---- pakai ulang bila file asli tidak berubah ----
            try
            {
                if (File.Exists(tujuan) && File.Exists(sidecar))
                {
                    var info = _json.Deserialize<Dictionary<string, object>>(File.ReadAllText(sidecar));
                    if (info != null && Convert.ToString(info["kunci"]) == kunci)
                    {
                        var h0 = new Hasil { Path = tujuan, DariCache = true };
                        var arr = info["rekening"] as System.Collections.IEnumerable;
                        if (arr != null) foreach (var x in arr) h0.Rekening.Add(Convert.ToString(x));
                        h0.BarisAsal = Convert.ToInt32(info["barisAsal"]);
                        h0.BarisTersisa = Convert.ToInt32(info["barisTersisa"]);
                        object c0;
                        if (info.TryGetValue("cadangan", out c0)) h0.Cadangan = Convert.ToInt32(c0);
                        if (info.TryGetValue("berbeda", out c0)) h0.Berbeda = Convert.ToInt32(c0);
                        File.SetLastWriteTimeUtc(sidecar, DateTime.UtcNow);   // tandai masih dipakai (pembersihan berbasis umur)
                        return h0;
                    }
                }
            }
            catch { /* sidecar rusak → buat ulang */ }

            if (File.Exists(tujuan)) File.Delete(tujuan);
            File.Copy(full, tujuan);
            var tambahan = rekTambahan == null ? null : new HashSet<string>(rekTambahan, StringComparer.OrdinalIgnoreCase);
            var h = new Hasil { Path = tujuan };

            Excel.Workbook wb = null;
            bool layar = true, alert = true, ev = true;
            try { layar = app.ScreenUpdating; alert = app.DisplayAlerts; ev = app.EnableEvents; } catch { }
            try
            {
                try { app.ScreenUpdating = false; app.DisplayAlerts = false; app.EnableEvents = false; } catch { }
                wb = app.Workbooks.Open(tujuan, UpdateLinks: 0, ReadOnly: false);
                foreach (var kc in ModeHitung.SemuaKC)
                {
                    Excel.Worksheet ws = ParameterMaster.CariSheet(wb, kc);
                    if (ws == null) continue;
                    SaringSheet(ws, f, tambahan, h);
                }
                wb.Save();
            }
            catch
            {
                if (wb != null) try { wb.Close(false); } catch { }
                wb = null;
                try { File.Delete(tujuan); } catch { }
                throw;
            }
            finally
            {
                if (wb != null) try { wb.Close(false); } catch { }
                try { app.EnableEvents = ev; app.DisplayAlerts = alert; app.ScreenUpdating = layar; } catch { }
            }

            var rek = new List<string>(h.Rekening);
            File.WriteAllText(sidecar, _json.Serialize(new Dictionary<string, object>
            {
                { "kunci", kunci }, { "asal", full }, { "rekening", rek }, { "barisAsal", h.BarisAsal }, { "barisTersisa", h.BarisTersisa },
                { "cadangan", h.Cadangan }, { "berbeda", h.Berbeda },
                { "dibuat", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") }
            }));
            return h;
        }

        /// <summary>
        /// Hapus baris data di luar grup. Baris yang tersisa dipertahankan urutannya: kolom bantu berisi nomor urut
        /// (baris tersisa) atau nomor urut + n (baris dibuang), area diurutkan menurut kolom itu, lalu blok baris
        /// dibuang di bagian bawah dihapus sekaligus. Nilai sel (termasuk teks seperti "0010003759") tidak ditulis ulang.
        /// </summary>
        private static void SaringSheet(Excel.Worksheet ws, FilterLingkup f, HashSet<string> tambahan, Hasil h)
        {
            if (ws.ProtectContents)
                throw new InvalidOperationException("Sheet " + ws.Name + " pada file sumber terproteksi; saringan per grup tidak dapat dibuat.");
            try { if (ws.AutoFilterMode) ws.AutoFilterMode = false; } catch { }

            Excel.Range used = ws.UsedRange;
            int kolomAkhir = used.Column + used.Columns.Count - 1;
            // Baris terakhir berisi data (kolom sandi atau nomor rekening) — UsedRange bisa jauh lebih panjang karena format
            int barisAkhir = Math.Max(BarisTerakhir(ws, f.NomorKolom), BarisTerakhir(ws, FilterLingkup.NomorDariHuruf(KolomRekening)));
            if (barisAkhir < BarisData) return;
            int n = barisAkhir - BarisData + 1;

            object[] sandi = Kolom(ws, f.NomorKolom, n);
            object[] rek = Kolom(ws, FilterLingkup.NomorDariHuruf(KolomRekening), n);

            // Tahap 8c: jenis penggunaan hanya terisi di baris pembiayaan (nama terisi). Setiap kontrak memakai isian
            // baris pembiayaan PERTAMA-nya; baris agunan mengikuti kontraknya; kosong / tidak dikenal → Konsumsi.
            object[] unitBaris = sandi;
            ModeHitung.Mode mf = ModeHitung.Cari(f.ModeId);
            if (ModeHitung.UnitPerKontrak(mf))
            {
                // Tahap 10: berlaku juga untuk kelompok sektor ekonomi (kolom T); kosong → kelompok cadangan
                string cad = ModeHitung.UnitCadangan(mf);
                object[] nama = Kolom(ws, FilterLingkup.NomorDariHuruf(KolomNama), n);
                var perKontrak = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < n; i++)
                {
                    if (Kosong(nama[i])) continue;
                    string kode = ModeHitung.NormUnit(mf, sandi[i]);
                    if (kode.Length == 0) { kode = cad; if (!Kosong(rek[i]) || !Kosong(sandi[i])) h.Cadangan++; }
                    string r = NormRek(rek[i]);
                    if (r.Length == 0) continue;
                    string ada;
                    if (!perKontrak.TryGetValue(r, out ada)) perKontrak[r] = kode;
                    else if (ada != kode) h.Berbeda++;
                }
                unitBaris = new object[n];
                for (int i = 0; i < n; i++)
                {
                    string r = NormRek(rek[i]), kode;
                    if (r.Length == 0 || !perKontrak.TryGetValue(r, out kode))
                    {
                        kode = ModeHitung.NormUnit(mf, sandi[i]);
                        if (kode.Length == 0) kode = cad;
                    }
                    unitBaris[i] = kode;
                }
            }

            var simpan = new bool[n];
            int nSimpan = 0, nIsi = 0;
            for (int i = 0; i < n; i++)
            {
                bool kosong = Kosong(sandi[i]) && Kosong(rek[i]);
                if (kosong) { simpan[i] = true; continue; }   // baris kosong tidak mengubah apa pun → biarkan
                nIsi++;
                string r = NormRek(rek[i]);
                bool masuk = f.Cocok(unitBaris[i]) || (tambahan != null && r.Length > 0 && tambahan.Contains(r));
                simpan[i] = masuk;
                if (masuk) { nSimpan++; if (r.Length > 0) h.Rekening.Add(r); }
            }
            h.BarisAsal += nIsi;
            h.BarisTersisa += nSimpan;

            int dibuang = 0;
            for (int i = 0; i < n; i++) if (!simpan[i]) dibuang++;
            if (dibuang == 0) return;
            int tersisa = n - dibuang;

            // Kolom bantu tepat di kanan area terpakai
            int kBantu = kolomAkhir + 1;
            var kunci = new object[n, 1];
            for (int i = 0; i < n; i++) kunci[i, 0] = simpan[i] ? (double)i : (double)(n + i);
            Excel.Range rBantu = (Excel.Range)ws.Range[ws.Cells[BarisData, kBantu], ws.Cells[barisAkhir, kBantu]];
            rBantu.Value2 = kunci;
            try
            {
                Excel.Range area = (Excel.Range)ws.Range[ws.Cells[BarisData, 1], ws.Cells[barisAkhir, kBantu]];
                area.Sort(Key1: rBantu, Order1: Excel.XlSortOrder.xlAscending,
                          Header: Excel.XlYesNoGuess.xlNo, MatchCase: false,
                          Orientation: Excel.XlSortOrientation.xlSortColumns);   // xlSortColumns = urut atas ke bawah
                ((Excel.Range)ws.Range[ws.Cells[BarisData + tersisa, 1], ws.Cells[barisAkhir, 1]]).EntireRow.Delete();
            }
            catch (Exception ex)
            {
                // Mis. ada sel gabungan di area data → hapus blok baris dari bawah ke atas (lebih lambat, hasil sama)
                CatatanLog.Tulis("Saring " + ws.Name + ": urut gagal (" + ex.Message + "), hapus per blok");
                try { rBantu.ClearContents(); } catch { }
                int i = n - 1;
                while (i >= 0)
                {
                    if (simpan[i]) { i--; continue; }
                    int akhir = i;
                    while (i >= 0 && !simpan[i]) i--;
                    ((Excel.Range)ws.Range[ws.Cells[BarisData + i + 1, 1], ws.Cells[BarisData + akhir, 1]]).EntireRow.Delete();
                }
            }
            try { ((Excel.Range)ws.Columns[kBantu]).Delete(); } catch { }
        }

        private static int BarisTerakhir(Excel.Worksheet ws, int kolom)
        {
            try
            {
                Excel.Range c = (Excel.Range)ws.Cells[ws.Rows.Count, kolom];
                return c.End[Excel.XlDirection.xlUp].Row;
            }
            catch { return 0; }
        }

        private static object[] Kolom(Excel.Worksheet ws, int kolom, int n)
        {
            var hasil = new object[n];
            object v = ((Excel.Range)ws.Range[ws.Cells[BarisData, kolom], ws.Cells[BarisData + n - 1, kolom]]).Value2;
            var arr = v as object[,];
            if (arr == null) { if (n == 1) hasil[0] = v; return hasil; }
            int b0 = arr.GetLowerBound(0), b1 = arr.GetLowerBound(1);
            for (int i = 0; i < n; i++) hasil[i] = arr[b0 + i, b1];
            return hasil;
        }

        private static bool Kosong(object v)
        {
            return v == null || Convert.ToString(v).Trim().Length == 0;
        }

        /// <summary>Nomor rekening ternormalisasi (tanpa apostrof/spasi, angka tanpa nol di depan) — sama dengan ReferensiKC.NormRek.</summary>
        public static string NormRek(object v)
        {
            if (v == null) return "";
            string s = v is double ? ((long)Math.Round((double)v)).ToString() : Convert.ToString(v);
            return Data.ReferensiKC.NormRek(s);
        }

        private static string Hash(string s)
        {
            using (var sha = System.Security.Cryptography.SHA1.Create())
            {
                var b = sha.ComputeHash(Encoding.UTF8.GetBytes(s ?? ""));
                var sb = new StringBuilder();
                foreach (var x in b) sb.Append(x.ToString("x2"));
                return sb.ToString();
            }
        }

        /// <summary>Hapus salinan berumur lebih dari 3 hari (dipanggil sebelum menyiapkan salinan baru).</summary>
        public static void Bersihkan()
        {
            try
            {
                if (!Directory.Exists(Folder)) return;
                foreach (var f in Directory.GetFiles(Folder))
                {
                    try
                    {
                        string acuan = f.EndsWith(".json", StringComparison.OrdinalIgnoreCase) ? f : f + ".json";
                        DateTime t = File.Exists(acuan) ? File.GetLastWriteTimeUtc(acuan) : File.GetLastWriteTimeUtc(f);
                        if ((DateTime.UtcNow - t).TotalDays > 3) File.Delete(f);
                    }
                    catch { }
                }
            }
            catch { }
        }

        // =================================================================
        // Penyiapan untuk pipeline: ganti path file sumber pada parameter run
        // =================================================================
        /// <summary>
        /// Ganti path sumber langkah yang akan dijalankan dengan salinan tersaring. Mengembalikan ringkasan untuk log.
        /// Dipanggil di konteks makro (thread Excel).
        /// </summary>
        public static List<string> Siapkan(Excel.Application app, ParameterCKPN p, FilterLingkup f, ICollection<string> langkah)
        {
            var ringkas = new List<string>();
            var cache = new Dictionary<string, Hasil>(StringComparer.OrdinalIgnoreCase);
            var ids = new HashSet<string>(langkah, StringComparer.OrdinalIgnoreCase);
            Bersihkan();

            // Hitung jumlah file supaya progres jelas
            var daftar = new List<string>();
            if (ids.Contains("individu") && p.Individu != null) daftar.Add(p.Individu.FilePath);
            if (ids.Contains("summary") && p.Summary != null) daftar.Add(p.Summary.FilePath);
            if (ids.Contains("netflow") && p.NetFlow != null) daftar.AddRange(Pecah(p.NetFlow.FilePaths));
            if (ids.Contains("migration") && p.Migration != null)
            {
                if (p.Migration.RunAll) { daftar.AddRange(Pecah(p.Migration.AwalGab)); daftar.AddRange(Pecah(p.Migration.AkhirGab)); }
                else { daftar.Add(p.Migration.PathAwal); daftar.Add(p.Migration.PathAkhir); }
            }
            if (ids.Contains("lgdcs") && p.LgdCs != null) foreach (var e in EntriLgdCs(p.LgdCs.FileConfigStr)) daftar.Add(e[3]);
            int total = 0;
            var unik = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var d in daftar) if (!string.IsNullOrEmpty(d) && unik.Add(d)) total++;
            int ke = 0;

            ModeHitung.Mode f_mode = ModeHitung.Cari(f.ModeId);
            string f_unit = f_mode == null ? "unit" : f_mode.NamaUnit;
            string f_cad = ModeHitung.NamaUnitKode(f_mode, ModeHitung.UnitCadangan(f_mode));
            Func<string, ICollection<string>, Hasil> salin = (path, tambahan) =>
            {
                if (string.IsNullOrEmpty(path)) return null;
                string kunci = path + "|" + (tambahan == null ? "" : tambahan.Count.ToString());
                Hasil hs;
                if (tambahan == null && cache.TryGetValue(kunci, out hs)) return hs;
                try { Pemberitahu.Progres("Saring sumber per grup", Math.Min(++ke, Math.Max(total, 1)), Math.Max(total, 1), System.IO.Path.GetFileName(path)); } catch { }
                hs = Salin(app, path, f, tambahan);
                if (tambahan == null) cache[kunci] = hs;
                ringkas.Add(System.IO.Path.GetFileName(path) + ": " + hs.BarisTersisa + "/" + hs.BarisAsal + " baris" +
                            (hs.Cadangan > 0 ? ", " + hs.Cadangan + " baris tanpa " + f_unit + " → " + f_cad : "") +
                            (hs.Berbeda > 0 ? ", " + hs.Berbeda + " baris " + f_unit + " berbeda dalam satu kontrak (dipakai isian pertama)" : "") +
                            (hs.DariCache ? " (salinan tersimpan)" : ""));
                return hs;
            };

            if (ids.Contains("individu") && p.Individu != null && !string.IsNullOrEmpty(p.Individu.FilePath))
                p.Individu.FilePath = salin(p.Individu.FilePath, null).Path;
            if (ids.Contains("summary") && p.Summary != null && !string.IsNullOrEmpty(p.Summary.FilePath))
                p.Summary.FilePath = salin(p.Summary.FilePath, null).Path;

            if (ids.Contains("netflow") && p.NetFlow != null && !string.IsNullOrEmpty(p.NetFlow.FilePaths))
                p.NetFlow.FilePaths = GantiDaftar(p.NetFlow.FilePaths, x => salin(x, null).Path);

            if (ids.Contains("migration") && p.Migration != null)
            {
                if (p.Migration.RunAll)
                {
                    // Pasangan awal/akhir per triwulan sesuai urutan; file akhir ikut rekening dari file awal
                    var awal = (p.Migration.AwalGab ?? "").Split('|');
                    var akhir = (p.Migration.AkhirGab ?? "").Split('|');
                    for (int i = 0; i < awal.Length; i++)
                    {
                        if (awal[i].Trim().Length == 0) continue;
                        Hasil ha = salin(awal[i], null);
                        awal[i] = ha.Path;
                        if (i < akhir.Length && akhir[i].Trim().Length > 0) akhir[i] = salin(akhir[i], ha.Rekening).Path;
                    }
                    p.Migration.AwalGab = string.Join("|", awal);
                    p.Migration.AkhirGab = string.Join("|", akhir);
                }
                else if (!string.IsNullOrEmpty(p.Migration.PathAwal))
                {
                    Hasil ha = salin(p.Migration.PathAwal, null);
                    p.Migration.PathAwal = ha.Path;
                    if (!string.IsNullOrEmpty(p.Migration.PathAkhir)) p.Migration.PathAkhir = salin(p.Migration.PathAkhir, ha.Rekening).Path;
                }
            }

            if (ids.Contains("lgdcs") && p.LgdCs != null && !string.IsNullOrEmpty(p.LgdCs.FileConfigStr))
            {
                var baru = new List<string>();
                foreach (var e in EntriLgdCs(p.LgdCs.FileConfigStr))
                {
                    string path = e[3].Length > 0 ? salin(e[3], null).Path : "";
                    baru.Add(e[0] + "=" + e[1] + "=" + e[2] + "=" + path);
                }
                p.LgdCs.FileConfigStr = string.Join("|", baru.ToArray());
            }
            return ringkas;
        }

        /// <summary>
        /// Refresh Summary di luar run (Simpan grup, tombol Refresh Summary panel): pada mode per kantor template
        /// Master!D14 diganti salinan tersaring grup di Master!H8 supaya PPKA per KC tetap milik grup itu.
        /// </summary>
        public static void TerapkanKeSummary(Excel.Application app, ParameterCKPN p)
        {
            if (p == null || p.Summary == null || string.IsNullOrEmpty(p.Summary.FilePath)) return;
            FilterLingkup f = FilterLingkup.Dari(app);
            if (f == null) return;
            Siapkan(app, p, f, new[] { "summary" });
        }

        private static IEnumerable<string> Pecah(string s)
        {
            foreach (var x in (s ?? "").Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries))
                if (x.Trim().Length > 0) yield return x.Trim();
        }

        /// <summary>Ganti setiap path di string "a|b|c|" tanpa mengubah susunan pemisahnya.</summary>
        private static string GantiDaftar(string s, Func<string, string> ganti)
        {
            var bagian = (s ?? "").Split('|');
            for (int i = 0; i < bagian.Length; i++)
                if (bagian[i].Trim().Length > 0) bagian[i] = ganti(bagian[i].Trim());
            return string.Join("|", bagian);
        }

        /// <summary>"label=tahun=bulan=path|…" → [label, tahun, bulan, path]</summary>
        private static List<string[]> EntriLgdCs(string s)
        {
            var hasil = new List<string[]>();
            foreach (var x in (s ?? "").Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var p = x.Split(new[] { '=' }, 4);
                if (p.Length < 4) continue;
                hasil.Add(new[] { p[0], p[1], p[2], p[3].Trim() });
            }
            return hasil;
        }
    }
}
