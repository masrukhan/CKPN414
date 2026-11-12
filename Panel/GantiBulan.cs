using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using CKPNLibrary.Helpers;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Panel
{
    /// <summary>
    /// Ganti bulan pelaporan kertas kerja (sheet Master) dari panel.
    ///
    /// Alur:
    ///   1. <see cref="Rencana"/>: untuk bulan baru, hitung bulan/tahun setiap slot path di Master
    ///      (meniru rumus label Master yang bergantung pada C4 dan D17), lalu cari file sumber tiap slot
    ///      LANGSUNG DI SATU FOLDER SUMBER (termasuk subfolder), bukan dari isi sheet Master:
    ///        a. nama file berakhiran "_MMYYYY" sebelum ekstensi (mis. "apa_saja_012026.xlsm") — nama di depannya bebas;
    ///        b. bila tidak ada, file lain di folder itu yang namanya memuat bulan tersebut (mis. "jan26") → "nama beda";
    ///        c. bila tetap tidak ada, path Master saat ini untuk bulan yang sama (bila filenya masih ada).
    ///      Folder sumber = pilihan user di panel; bawaan = folder yang paling banyak dipakai path Master saat ini.
    ///   2. Panel menampilkan status per slot; user dapat memilih file lain (<see cref="PilihFile"/>).
    ///   3. <see cref="Terapkan"/>: tulis Master!C4 dan path terpilih. Label bulan, tanggal ref, dan periode
    ///      penilaian individu adalah rumus Excel yang mengikuti C4 dengan sendirinya.
    ///
    /// Slot path (sel) yang dikelola:
    ///   D14            CKPN Individu (bulan laporan)
    ///   D20:D32        PD Net Flow, 13 bulan (label B, tanggal ref F)
    ///   D47:D50/F47:F50 PD Migration, file posisi awal / akhir Triwulan I–IV (tanggal C40:D43)
    ///   D60:D65        LGD ER, Desember 6 tahun (tahun awal boleh kosong bila data belum ada)
    ///   D72:D83        LGD CS, Juni & Desember 6 tahun (periode awal boleh kosong bila data belum ada)
    /// </summary>
    internal static class GantiBulan
    {
        private static readonly string[] Bln3 = { "jan", "feb", "mar", "apr", "mei", "jun", "jul", "agu", "sep", "okt", "nov", "des" };
        private static readonly string[] BlnPanjang = { "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember" };
        private static readonly string[] BlnPendek = { "Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des" };

        internal class Slot
        {
            public string Bagian, Sel, Label;
            public int Tahun, Bulan;
            public bool BolehKosong;      // LGD ER / LGD CS: periode awal boleh kosong bila data belum ada
            public string PathKini = "", PathUsul = "", Status = "", Catatan = "";
        }

        // =============================================================
        // Slot path untuk tanggal posisi tertentu (meniru rumus Master)
        // =============================================================
        internal static List<Slot> SlotUntuk(DateTime c4, bool setahun)
        {
            var l = new List<Slot>();
            int y = c4.Year, m = c4.Month;

            l.Add(new Slot { Bagian = "individu", Sel = "D14", Tahun = y, Bulan = m, Label = BlnPendek[m - 1] + " " + Yy(y) });

            // Net Flow: F20:F32 = EDATE(basis, -(32-ROW())); basis = Des tahun lalu (setahun, bila C4 bukan Des) atau bulan C4
            var basisNf = setahun ? new DateTime(y - (m == 12 ? 0 : 1), 12, 1) : new DateTime(y, m, 1);
            for (int r = 20; r <= 32; r++)
            {
                var t = basisNf.AddMonths(-(32 - r));
                l.Add(new Slot { Bagian = "netflow", Sel = "D" + r, Tahun = t.Year, Bulan = t.Month, Label = BlnPendek[t.Month - 1] + " " + Yy(t.Year) });
            }

            // Migration C40:D43: akhir triwulan q (Mar/Jun/Sep/Des); awal = tahun sebelumnya
            string[] tw = { "Triwulan I", "Triwulan II", "Triwulan III", "Triwulan IV" };
            for (int q = 0; q < 4; q++)
            {
                int bq = (q + 1) * 3;
                int yAkhir = setahun ? y - (m == 12 ? 0 : 1) : y - (m >= bq ? 0 : 1);
                int yAwal = yAkhir - 1;
                l.Add(new Slot { Bagian = "migration", Sel = "D" + (47 + q), Tahun = yAwal, Bulan = bq, Label = tw[q] + " · awal " + BlnPendek[bq - 1] + " " + Yy(yAwal) });
                l.Add(new Slot { Bagian = "migration", Sel = "F" + (47 + q), Tahun = yAkhir, Bulan = bq, Label = tw[q] + " · akhir " + BlnPendek[bq - 1] + " " + Yy(yAkhir) });
            }

            // LGD ER D60:D65: Desember (YEAR(EDATE(C4,-11)) - (65-ROW()))
            int cy = c4.AddMonths(-11).Year;
            for (int r = 60; r <= 65; r++)
            {
                int ty = cy - (65 - r);
                l.Add(new Slot { Bagian = "lgder", Sel = "D" + r, Tahun = ty, Bulan = 12, Label = "Des " + ty, BolehKosong = r < 65 });
            }

            // LGD CS D72:D83: Juni & Desember, tahun cy-5 … cy
            for (int r = 72; r <= 83; r++)
            {
                int ty = cy - 5 + (r - 72) / 2, tb = (r - 72) % 2 == 0 ? 6 : 12;
                l.Add(new Slot { Bagian = "lgdcs", Sel = "D" + r, Tahun = ty, Bulan = tb, Label = BlnPendek[tb - 1] + " " + ty, BolehKosong = r < 83 });
            }
            return l;
        }

        private static string Yy(int y) { return (y % 100).ToString("00"); }

        private static readonly HashSet<string> SelDiizinkan = new HashSet<string>(
            SlotUntuk(new DateTime(2026, 1, 31), false).Select(s => s.Sel), StringComparer.OrdinalIgnoreCase);

        // =============================================================
        // Aturan nama file: <nama bebas>_MMYYYY.<ekstensi>
        // =============================================================
        private static readonly Regex RxAkhiranBulan = new Regex("_([0-9]{2})([0-9]{4})$");
        private const string ContohPola = "…_MMYYYY.xlsm / .xlsx (mis. …_012026.xlsm)";
        private const int KedalamanFolder = 3;   // folder sumber + subfolder s.d. 3 tingkat

        /// <summary>Bulan-tahun (yyyymm) bila nama file berakhiran _MMYYYY; 0 bila tidak.</summary>
        internal static int KunciAkhiran(string namaFile)
        {
            var mt = RxAkhiranBulan.Match(Path.GetFileNameWithoutExtension(namaFile ?? ""));
            if (!mt.Success) return 0;
            int bl = int.Parse(mt.Groups[1].Value), th = int.Parse(mt.Groups[2].Value);
            return bl >= 1 && bl <= 12 && th >= 2000 && th <= 2100 ? th * 100 + bl : 0;
        }

        /// <summary>Nama depan (sebelum _MMYYYY) dan ekstensi dari path yang mengikuti aturan; null bila tidak.</summary>
        private static string[] NamaDepan(string path)
        {
            if (string.IsNullOrEmpty(path) || KunciAkhiran(path) == 0) return null;
            string n = Path.GetFileNameWithoutExtension(path);
            return new[] { n.Substring(0, n.Length - 7), Path.GetExtension(path) };
        }

        private class Calon
        {
            public string Path; public int Kedalaman;
        }

        /// <summary>Daftar file Excel di folder sumber + subfolder (tanpa file kunci "~$").</summary>
        private static List<Calon> DaftarFile(string folder)
        {
            var hasil = new List<Calon>();
            var antre = new Queue<KeyValuePair<string, int>>();
            antre.Enqueue(new KeyValuePair<string, int>(folder, 0));
            while (antre.Count > 0)
            {
                var d = antre.Dequeue();
                try
                {
                    foreach (var f in Directory.GetFiles(d.Key, "*.xls*"))
                        if (!Path.GetFileName(f).StartsWith("~$")) hasil.Add(new Calon { Path = f, Kedalaman = d.Value });
                    if (d.Value < KedalamanFolder)
                        foreach (var sub in Directory.GetDirectories(d.Key))
                            antre.Enqueue(new KeyValuePair<string, int>(sub, d.Value + 1));
                }
                catch (Exception ex) { CatatanLog.Tulis("Pindai folder " + d.Key + ": " + ex.Message); }
            }
            return hasil;
        }

        /// <summary>
        /// Calon terbaik bila ada beberapa file untuk bulan yang sama (mis. template + file hasil ekspor):
        /// nama depan sama dengan yang dipakai Master saat ini untuk bagian itu, lalu folder paling atas,
        /// lalu .xlsm, lalu nama.
        /// </summary>
        private static string Terbaik(List<Calon> l, string[] namaDepan)
        {
            string depan = namaDepan != null ? namaDepan[0] : null;
            return l.OrderBy(x => depan != null && Path.GetFileName(x.Path).StartsWith(depan + "_", StringComparison.OrdinalIgnoreCase) ? 0 : 1)
                    .ThenBy(x => x.Kedalaman)
                    .ThenBy(x => x.Path.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase) ? 0 : 1)
                    .ThenBy(x => x.Path, StringComparer.OrdinalIgnoreCase).First().Path;
        }

        private static readonly Regex RxAngka = new Regex("(?<![0-9])([0-9]{4,6})(?![0-9])");
        private static readonly Regex RxNamaBulan = new Regex(
            "(?<![a-z])(jan|feb|mar|apr|mei|may|jun|jul|agu|ags|agt|aug|sep|okt|oct|nov|des|dec)[a-z]*[ _.-]?([0-9]{4}|[0-9]{2})(?![0-9])",
            RegexOptions.IgnoreCase);

        /// <summary>Bulan-tahun dari nama file (MMYYYY, YYYYMM, MMYY, atau nama bulan + tahun). 0 bila tidak terbaca.</summary>
        internal static int KunciDariNama(string namaFile)
        {
            string n = Path.GetFileNameWithoutExtension(namaFile ?? "");
            foreach (Match mt in RxAngka.Matches(n))
            {
                string d = mt.Groups[1].Value;
                int a, b;
                if (d.Length == 6)
                {
                    a = int.Parse(d.Substring(0, 2)); b = int.Parse(d.Substring(2));
                    if (a >= 1 && a <= 12 && b >= 2000 && b <= 2100) return b * 100 + a;
                    a = int.Parse(d.Substring(0, 4)); b = int.Parse(d.Substring(4));
                    if (a >= 2000 && a <= 2100 && b >= 1 && b <= 12) return a * 100 + b;
                }
                else if (d.Length == 4)
                {
                    a = int.Parse(d.Substring(0, 2)); b = int.Parse(d.Substring(2));
                    if (a >= 1 && a <= 12 && b >= 10 && b <= 60) return (2000 + b) * 100 + a;
                }
            }
            var mb = RxNamaBulan.Match(n);
            if (mb.Success)
            {
                int bl = BulanDariNama(mb.Groups[1].Value);
                int th = int.Parse(mb.Groups[2].Value);
                if (th < 100) th += 2000;
                if (bl > 0 && th >= 2000 && th <= 2100) return th * 100 + bl;
            }
            return 0;
        }

        private static int BulanDariNama(string s)
        {
            switch (s.ToLowerInvariant())
            {
                case "jan": return 1; case "feb": return 2; case "mar": return 3; case "apr": return 4;
                case "mei": case "may": return 5; case "jun": return 6; case "jul": return 7;
                case "agu": case "ags": case "agt": case "aug": return 8; case "sep": return 9;
                case "okt": case "oct": return 10; case "nov": return 11; case "des": case "dec": return 12;
            }
            return 0;
        }

        // =============================================================
        // Rencana
        // =============================================================
        public static Dictionary<string, object> Rencana(Excel.Application app, string tanggal, string folderPilihan)
        {
            Excel.Workbook wb = PanelBridge.CariWorkbookAplikasi(app);
            if (wb == null) throw new InvalidOperationException("Workbook aplikasi CKPN tidak sedang terbuka.");
            Excel.Worksheet m = ParameterMaster.CariSheet(wb, "Master");
            if (m == null) throw new InvalidOperationException("Sheet 'Master' tidak ditemukan.");

            DateTime kini;
            bool adaKini = ParameterMaster.BacaTanggalLaporan(wb, out kini);
            bool setahun = TeksSel(m, "D17").IndexOf("setahun", StringComparison.OrdinalIgnoreCase) >= 0;

            DateTime baru;
            if (!DateTime.TryParseExact(tanggal ?? "", "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out baru))
                throw new InvalidOperationException("Tanggal bulan baru tidak valid: " + tanggal);
            baru = new DateTime(baru.Year, baru.Month, DateTime.DaysInMonth(baru.Year, baru.Month));

            // ---- Isi Master saat ini: file per bulan (cadangan), nama depan per bagian, folder terbanyak ----
            var dikenal = new Dictionary<int, string>();
            var namaDepan = new Dictionary<string, string[]>();          // bagian → {nama depan, ekstensi}
            var hitungFolder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            if (adaKini)
            {
                foreach (var s in SlotUntuk(kini, setahun))
                {
                    string p = TeksSel(m, s.Sel);
                    if (p.Length == 0 || p.ToLowerInvariant().Contains("pilih")) continue;
                    int k = s.Tahun * 100 + s.Bulan;
                    if (File.Exists(p) && !dikenal.ContainsKey(k)) dikenal[k] = p;
                    var nd = NamaDepan(p);
                    if (nd != null && !namaDepan.ContainsKey(s.Bagian)) namaDepan[s.Bagian] = nd;
                    string f = Path.GetDirectoryName(p);
                    if (!string.IsNullOrEmpty(f)) { int c; hitungFolder.TryGetValue(f, out c); hitungFolder[f] = c + 1; }
                }
            }

            // ---- Satu folder sumber untuk semua slot (individu, PD, LGD) ----
            string folder = (folderPilihan ?? "").Trim();
            bool folderDipilih = folder.Length > 0;
            if (!folderDipilih)
                folder = hitungFolder.Count == 0 ? "" : hitungFolder.OrderByDescending(x => x.Value).First().Key;
            if (folder.Length > 0 && !Directory.Exists(folder))
                throw new InvalidOperationException("Folder sumber tidak ditemukan: " + folder);

            // ---- Indeks isi folder: bulan → file berakhiran _MMYYYY (tepat) / file lain yang memuat bulan (longgar) ----
            var tepat = new Dictionary<int, List<Calon>>();
            var longgar = new Dictionary<int, List<Calon>>();
            int jumlahFile = 0;
            if (folder.Length > 0)
                foreach (var c in DaftarFile(folder))
                {
                    jumlahFile++;
                    string nm = Path.GetFileName(c.Path);
                    int k = KunciAkhiran(nm);
                    var tujuan = tepat;
                    if (k == 0) { k = KunciDariNama(nm); tujuan = longgar; }
                    if (k == 0) continue;
                    List<Calon> l;
                    if (!tujuan.TryGetValue(k, out l)) { l = new List<Calon>(); tujuan[k] = l; }
                    l.Add(c);
                }

            // ---- Slot untuk bulan baru ----
            string mmyyyy;
            var slot = SlotUntuk(baru, setahun);
            foreach (var s in slot)
            {
                s.PathKini = TeksSel(m, s.Sel);
                int k = s.Tahun * 100 + s.Bulan;
                mmyyyy = s.Bulan.ToString("00") + s.Tahun.ToString("0000");
                List<Calon> l;
                string kenal;
                string[] ndBagian;
                namaDepan.TryGetValue(s.Bagian, out ndBagian);
                if (tepat.TryGetValue(k, out l) && l.Count > 0)
                {
                    s.PathUsul = Terbaik(l, ndBagian);
                    s.Status = "ditemukan";
                    if (l.Count > 1) s.Catatan = l.Count + " file berakhiran _" + mmyyyy + " di folder sumber; dipilih yang nama depannya sama dengan Master saat ini" +
                                                 " (lalu folder paling atas, .xlsm). Klik baris bila keliru.";
                }
                else if (longgar.TryGetValue(k, out l) && l.Count > 0)
                {
                    s.PathUsul = Terbaik(l, ndBagian);
                    s.Status = "nama-beda";
                    s.Catatan = "Nama file tidak berakhiran _" + mmyyyy + " — pastikan isinya posisi " + BlnPanjang[s.Bulan - 1] + " " + s.Tahun + ".";
                }
                else if (dikenal.TryGetValue(k, out kenal))
                {
                    // file bulan ini sudah tercatat di Master tetapi berada di luar folder sumber
                    s.PathUsul = kenal;
                    s.Status = "nama-beda";
                    s.Catatan = "Tidak ada di folder sumber; dipakai path Master saat ini.";
                }
                else
                {
                    // usulan path: nama depan bagian ini dari Master saat ini + _MMYYYY, di folder sumber
                    s.PathUsul = folder.Length > 0 && ndBagian != null
                        ? Path.Combine(folder, ndBagian[0] + "_" + mmyyyy + ndBagian[1]) : "";
                    s.Status = "tidak-ada";
                    s.Catatan = "Tidak ada file berakhiran _" + mmyyyy + " di folder sumber.";
                }
            }

            // LGD ER / CS: periode awal tanpa file boleh dikosongkan (data hapus buku / macet belum ada)
            foreach (var bag in new[] { "lgder", "lgdcs" })
            {
                foreach (var s in slot.Where(x => x.Bagian == bag))
                {
                    if (s.Status != "tidak-ada" || !s.BolehKosong) break;
                    s.Status = "kosong";
                    s.PathUsul = "";
                    s.Catatan = "Belum ada data periode ini — baris dikosongkan.";
                }
            }

            // ---- Ringkasan ----
            var nf = slot.Where(x => x.Bagian == "netflow").ToList();
            var er = slot.Where(x => x.Bagian == "lgder").ToList();
            var cs = slot.Where(x => x.Bagian == "lgdcs").ToList();
            bool pakaiAcuan = setahun && baru.Month != 12;
            string acuanTahun = (baru.Year - 1).ToString();

            var bagian = new List<object>();
            Func<string, string, string, Dictionary<string, object>> buatBagian = (id, judul, ket) =>
            {
                var isi = slot.Where(x => x.Bagian == id).Select(x => (object)new Dictionary<string, object>
                {
                    { "sel", x.Sel }, { "label", x.Label }, { "tahun", x.Tahun }, { "bulan", x.Bulan },
                    { "pathKini", x.PathKini }, { "pathUsul", x.PathUsul }, { "status", x.Status }, { "catatan", x.Catatan },
                    { "bolehKosong", x.BolehKosong }, { "berubah", !string.Equals(x.PathKini, x.PathUsul, StringComparison.OrdinalIgnoreCase) }
                }).ToList();
                return new Dictionary<string, object> { { "id", id }, { "judul", judul }, { "keterangan", ket }, { "baris", isi } };
            };
            string ketPdLgd = pakaiAcuan
                ? "Metode PD & LGD setahun sekali: bulan ini memakai acuan Desember " + acuanTahun + "; file di bawah hanya dibaca bila Hitung penuh."
                : "";
            bagian.Add(buatBagian("individu", "CKPN Individu & ringkasan", "File template APOLLO bulan laporan (Master!D14)."));
            bagian.Add(buatBagian("netflow", "PD Net Flow · 13 bulan", ketPdLgd));
            bagian.Add(buatBagian("migration", "PD Migration · posisi awal & akhir triwulan", ketPdLgd));
            bagian.Add(buatBagian("lgder", "LGD Expected Recoveries · Desember 6 tahun", ketPdLgd));
            bagian.Add(buatBagian("lgdcs", "LGD Collateral Shortfall · Juni & Desember 6 tahun", ketPdLgd));

            var ikut = new List<string>
            {
                "Tanggal posisi (C4) → " + baru.Day + " " + BlnPanjang[baru.Month - 1] + " " + baru.Year,
                "Periode penilaian individu (C7) → " + BlnPanjang[baru.Month - 1] + " " + baru.Year,
                "Label bulan & tanggal ref PD Net Flow → " + nf[0].Label + " s.d. " + nf[nf.Count - 1].Label,
                "Tanggal triwulan PD Migration (C40:D43) dan tahun LGD ER/CS → " + er[0].Tahun + "–" + er[er.Count - 1].Tahun
            };

            int total = slot.Count, siap = slot.Count(x => x.Status == "ditemukan" || x.Status == "nama-beda" || x.Status == "kosong");
            return new Dictionary<string, object>
            {
                { "tanggal", baru.ToString("yyyy-MM-dd") },
                { "tanggalKini", adaKini ? kini.ToString("yyyy-MM-dd") : "" },
                { "setahun", setahun }, { "pakaiAcuan", pakaiAcuan },
                { "ringkas", "Net Flow memakai " + nf[0].Label + " s.d. " + nf[nf.Count - 1].Label +
                             " · LGD " + er[0].Tahun + "–" + er[er.Count - 1].Tahun + (pakaiAcuan ? " · PD & LGD acuan Des " + acuanTahun : "") },
                { "folder", folder }, { "folderDipilih", folderDipilih }, { "jumlahFile", jumlahFile },
                { "pola", ContohPola },
                { "bagian", bagian }, { "ikut", ikut },
                { "total", total }, { "siap", siap },
                { "perluCek", slot.Count(x => x.Status == "nama-beda") },
                { "tidakAda", slot.Count(x => x.Status == "tidak-ada") }
            };
        }

        // =============================================================
        // Pilih file lain untuk satu slot
        // =============================================================
        public static Dictionary<string, object> PilihFile(string judul, string awal)
        {
            using (var dlg = new System.Windows.Forms.OpenFileDialog
            {
                Title = string.IsNullOrEmpty(judul) ? "Pilih file sumber" : judul,
                Filter = "File Excel|*.xlsx;*.xlsm;*.xls|Semua file|*.*"
            })
            {
                try
                {
                    string dir = string.IsNullOrEmpty(awal) ? null : (Directory.Exists(awal) ? awal : Path.GetDirectoryName(awal));
                    if (!string.IsNullOrEmpty(dir) && Directory.Exists(dir)) dlg.InitialDirectory = dir;
                    if (!string.IsNullOrEmpty(awal) && File.Exists(awal)) dlg.FileName = Path.GetFileName(awal);
                }
                catch { }
                if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                    return new Dictionary<string, object> { { "batal", true } };
                return new Dictionary<string, object>
                {
                    { "batal", false }, { "path", dlg.FileName },
                    { "kunci", KunciAkhiran(dlg.FileName) != 0 ? KunciAkhiran(dlg.FileName) : KunciDariNama(dlg.FileName) }
                };
            }
        }

        // =============================================================
        // Pilih folder sumber (satu folder untuk semua file individu, PD, dan LGD)
        // =============================================================
        public static Dictionary<string, object> PilihFolder(string awal)
        {
            using (var dlg = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Pilih folder sumber file template (nama file berakhiran _MMYYYY, mis. …_012026.xlsm)",
                ShowNewFolderButton = false
            })
            {
                try { if (!string.IsNullOrEmpty(awal) && Directory.Exists(awal)) dlg.SelectedPath = awal; } catch { }
                if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                    return new Dictionary<string, object> { { "batal", true } };
                return new Dictionary<string, object> { { "batal", false }, { "path", dlg.SelectedPath } };
            }
        }

        // =============================================================
        // Terapkan ke sheet Master
        // =============================================================
        /// <param name="isi">[{sel, path}] — path kosong = sel dikosongkan</param>
        public static Dictionary<string, object> Terapkan(Excel.Application app, string tanggal, IEnumerable isi)
        {
            if (CKPNPipeline.SedangBerjalan || BatchGrup.Aktif)
                throw new InvalidOperationException("Perhitungan sedang berjalan. Ganti bulan setelah selesai.");
            Excel.Workbook wb = PanelBridge.CariWorkbookAplikasi(app);
            if (wb == null) throw new InvalidOperationException("Workbook aplikasi CKPN tidak sedang terbuka.");

            DateTime baru;
            if (!DateTime.TryParseExact(tanggal ?? "", "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out baru))
                throw new InvalidOperationException("Tanggal bulan baru tidak valid: " + tanggal);
            baru = new DateTime(baru.Year, baru.Month, DateTime.DaysInMonth(baru.Year, baru.Month));

            var tulis = new List<KeyValuePair<string, string>>();
            if (isi != null)
                foreach (var o in isi)
                {
                    var d = o as IDictionary<string, object>;
                    if (d == null) continue;
                    object vs, vp;
                    d.TryGetValue("sel", out vs); d.TryGetValue("path", out vp);
                    string sel = Convert.ToString(vs ?? "").Trim().ToUpperInvariant();
                    string path = Convert.ToString(vp ?? "").Trim();
                    if (!SelDiizinkan.Contains(sel)) throw new InvalidOperationException("Sel Master tidak dikenal: " + sel);
                    tulis.Add(new KeyValuePair<string, string>(sel, path));
                }

            DateTime lama;
            bool adaLama = ParameterMaster.BacaTanggalLaporan(wb, out lama);
            Protection.UbahMaster(wb, ms =>
            {
                ((Excel.Range)ms.Range["C4"]).Value2 = baru.ToOADate();
                foreach (var kv in tulis)
                {
                    Excel.Range r = (Excel.Range)ms.Range[kv.Key];
                    if (kv.Value.Length == 0) r.Value2 = null;
                    else r.Value2 = kv.Value;
                }
            });
            try { wb.Application.Calculate(); } catch { }

            string ket = "Ganti bulan pelaporan: " + (adaLama ? lama.ToString("yyyy-MM-dd") : "—") + " → " + baru.ToString("yyyy-MM-dd") +
                         " · " + tulis.Count + " path ditulis";
            CatatanLog.Tulis(ket);
            LogProses.CatatPanel(baru.ToString("yyyy-MM-dd"), "Ganti bulan pelaporan", "Master", LogProses.OK, LogProses.R()
                .Tambah("Dari", adaLama ? lama.ToString("yyyy-MM-dd") : "—").Tambah("Ke", baru.ToString("yyyy-MM-dd"))
                .Tambah("Path ditulis", tulis.Count));
            return new Dictionary<string, object> { { "tanggal", baru.ToString("yyyy-MM-dd") }, { "jumlah", tulis.Count } };
        }

        private static string TeksSel(Excel.Worksheet m, string alamat)
        {
            try
            {
                object v = ((Excel.Range)m.Range[alamat]).Value2;
                return v == null || v is int ? "" : Convert.ToString(v, CultureInfo.InvariantCulture).Trim();
            }
            catch { return ""; }
        }
    }
}
