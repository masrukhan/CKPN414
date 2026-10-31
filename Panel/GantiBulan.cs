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
    ///      (meniru rumus label Master yang bergantung pada C4 dan D17), lalu cari file sumber tiap slot:
    ///        a. pola nama yang dipelajari dari path Master saat ini (mis. "..._MMYYYY.xlsm");
    ///        b. bila tidak ada, pindai folder yang sama dan cocokkan bulan dari nama file (pola lain, mis. "jan26").
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
        // Pola nama file
        // =============================================================
        /// <summary>Teks bulan dalam berbagai pola; urutan = prioritas (yang panjang dulu).</summary>
        private static IEnumerable<KeyValuePair<string, string>> Rendering(int y, int m)
        {
            yield return new KeyValuePair<string, string>("MMYYYY", m.ToString("00") + y.ToString("0000"));
            yield return new KeyValuePair<string, string>("YYYYMM", y.ToString("0000") + m.ToString("00"));
            yield return new KeyValuePair<string, string>("bbbYYYY", Bln3[m - 1] + y.ToString("0000"));
            yield return new KeyValuePair<string, string>("bbbYY", Bln3[m - 1] + Yy(y));
            yield return new KeyValuePair<string, string>("MMYY", m.ToString("00") + Yy(y));
        }

        private static string Render(string pola, int y, int m)
        {
            foreach (var kv in Rendering(y, m)) if (kv.Key == pola) return kv.Value;
            return null;
        }

        private class Pola
        {
            public string Folder, Awalan, Akhiran, Kode;   // nama = Awalan + Render(Kode) + Akhiran
            public bool Kapital;                            // nama bulan berhuruf besar awal (Jan26)
            public string Kunci { get { return (Folder + "|" + Awalan + "|" + Kode + "|" + Akhiran).ToLowerInvariant(); } }
            public string Buat(int y, int m)
            {
                string r = Render(Kode, y, m);
                if (Kapital && r.Length > 0) r = char.ToUpperInvariant(r[0]) + r.Substring(1);
                return Path.Combine(Folder, Awalan + r + Akhiran);
            }
            public string Contoh { get { return Awalan + "{" + Kode.Replace("bbb", "Bln") + "}" + Akhiran; } }
        }

        /// <summary>Pola nama dari path yang diketahui bulannya (cocokkan teks bulan di nama file, batas bukan digit).</summary>
        private static Pola PelajariPola(string path, int y, int m)
        {
            if (string.IsNullOrEmpty(path)) return null;
            string nama = Path.GetFileName(path), folder = Path.GetDirectoryName(path) ?? "";
            foreach (var kv in Rendering(y, m))
            {
                var mt = Regex.Match(nama, "(?<![0-9])" + Regex.Escape(kv.Value) + "(?![0-9])", RegexOptions.IgnoreCase);
                if (!mt.Success) continue;
                return new Pola
                {
                    Folder = folder, Kode = kv.Key,
                    Awalan = nama.Substring(0, mt.Index), Akhiran = nama.Substring(mt.Index + mt.Length),
                    Kapital = kv.Key.StartsWith("bbb") && char.IsUpper(nama[mt.Index])
                };
            }
            return null;
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
        public static Dictionary<string, object> Rencana(Excel.Application app, string tanggal)
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

            // ---- Peta bulan → file dari isi Master saat ini (slot dihitung dari C4 saat ini) ----
            var dikenal = new Dictionary<int, string>();
            var hitungPola = new Dictionary<string, int>();
            var contohPola = new Dictionary<string, Pola>();
            var folder = new List<string>();
            if (adaKini)
            {
                foreach (var s in SlotUntuk(kini, setahun))
                {
                    string p = TeksSel(m, s.Sel);
                    if (p.Length == 0 || p.ToLowerInvariant().Contains("pilih")) continue;
                    int k = s.Tahun * 100 + s.Bulan;
                    if (File.Exists(p) && !dikenal.ContainsKey(k)) dikenal[k] = p;
                    var po = PelajariPola(p, s.Tahun, s.Bulan);
                    if (po != null)
                    {
                        int c; hitungPola.TryGetValue(po.Kunci, out c); hitungPola[po.Kunci] = c + 1;
                        if (!contohPola.ContainsKey(po.Kunci)) contohPola[po.Kunci] = po;
                    }
                    string f = Path.GetDirectoryName(p);
                    if (!string.IsNullOrEmpty(f) && !folder.Contains(f, StringComparer.OrdinalIgnoreCase)) folder.Add(f);
                }
            }
            Pola pola = hitungPola.Count == 0 ? null : contohPola[hitungPola.OrderByDescending(x => x.Value).First().Key];
            if (pola != null) { folder.Remove(pola.Folder); folder.Insert(0, pola.Folder); }

            // ---- Indeks file di folder sumber: bulan → daftar path ----
            var indeks = new Dictionary<int, List<string>>();
            foreach (var f in folder)
            {
                string[] files;
                try { files = Directory.GetFiles(f, "*.xls*"); }
                catch (Exception ex) { CatatanLog.Tulis("Pindai folder " + f + ": " + ex.Message); continue; }
                foreach (var file in files)
                {
                    string nm = Path.GetFileName(file);
                    if (nm.StartsWith("~$")) continue;
                    int k = KunciDariNama(nm);
                    if (k == 0) continue;
                    List<string> l;
                    if (!indeks.TryGetValue(k, out l)) { l = new List<string>(); indeks[k] = l; }
                    l.Add(file);
                }
            }

            // ---- Slot untuk bulan baru ----
            var slot = SlotUntuk(baru, setahun);
            foreach (var s in slot)
            {
                s.PathKini = TeksSel(m, s.Sel);
                int k = s.Tahun * 100 + s.Bulan;
                string pathPola = pola != null ? pola.Buat(s.Tahun, s.Bulan) : null;
                string dipilih = null;
                if (pathPola != null && File.Exists(pathPola)) dipilih = pathPola;
                else
                {
                    string kenal;
                    if (dikenal.TryGetValue(k, out kenal)) dipilih = kenal;
                    else
                    {
                        List<string> l;
                        if (indeks.TryGetValue(k, out l) && l.Count > 0)
                            dipilih = l.OrderBy(x => x.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase) ? 0 : 1)
                                       .ThenBy(x => x, StringComparer.OrdinalIgnoreCase).First();
                    }
                }
                if (dipilih != null)
                {
                    s.PathUsul = dipilih;
                    s.Status = pathPola != null && string.Equals(dipilih, pathPola, StringComparison.OrdinalIgnoreCase) ? "ditemukan"
                             : pola == null ? "ditemukan" : "nama-beda";
                    if (s.Status == "nama-beda") s.Catatan = "Nama file tidak mengikuti pola " + pola.Contoh + " — pastikan isinya posisi " + BlnPanjang[s.Bulan - 1] + " " + s.Tahun + ".";
                }
                else
                {
                    s.PathUsul = pathPola ?? "";
                    s.Status = "tidak-ada";
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
                { "folder", pola != null ? pola.Folder : (folder.Count > 0 ? folder[0] : "") },
                { "pola", pola != null ? pola.Contoh : "" },
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
                    { "batal", false }, { "path", dlg.FileName }, { "kunci", KunciDariNama(dlg.FileName) }
                };
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
