using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Helpers
{
    /// <summary>
    /// Log proses perhitungan CKPN (Tahap 5) — pengganti sheet "Audit Log".
    ///
    /// Sebelumnya setiap modul menambah baris ke sheet "Audit Log" di workbook
    /// aplikasi, sehingga file .xlsm terus membesar. Kini catatan yang sama ditulis
    /// ke file teks dan dibaca lewat Panel CKPN (tab Riwayat › Log proses).
    ///
    /// Lokasi  : library\logs\proses\proses_{yyyy-MM}_{user}.txt
    ///           (satu file per user per bulan kalender → tidak saling mengunci di
    ///           shared drive, dan ukuran tiap file tetap kecil)
    /// Cadangan: %LOCALAPPDATA%\CKPN414\logs\proses\ bila folder library tidak bisa ditulis
    ///
    /// Format  : teks dipisah TAB, satu kejadian per baris, baris pertama berisi
    ///           judul kolom — bisa dibuka di Notepad atau di-import ke Excel.
    ///   waktu | pengguna | sumber | jalan | periode | proses | label | status | rincian
    ///   rincian = "Nama=Nilai; Nama=Nilai; ..."
    ///
    /// Seperti CatatanLog, log TIDAK PERNAH boleh menggagalkan perhitungan:
    /// semua error ditelan.
    /// </summary>
    internal static class LogProses
    {
        // ---- Status baku ----
        public const string OK         = "OK";
        public const string Peringatan = "PERINGATAN";
        public const string Gagal      = "GAGAL";
        public const string Dilewati   = "DILEWATI";
        public const string Info       = "INFO";

        public static readonly string[] Kolom =
            { "waktu", "pengguna", "sumber", "jalan", "periode", "proses", "label", "status", "rincian" };

        private static readonly object _kunci = new object();
        internal static readonly CultureInfo Id = new CultureInfo("id-ID");

        // ---- Konteks jalan (diisi CKPNPipeline selama run dari panel) ----
        private static string _idJalan;
        private static string _sumber;
        private static string _periode;

        /// <summary>
        /// Tandai awal satu run panel. Semua catatan modul sampai <see cref="SelesaiJalan"/>
        /// memakai id jalan yang sama, sehingga bisa difilter bersama di panel.
        /// </summary>
        public static string MulaiJalan(string sumber, string periode)
        {
            _idJalan = DateTime.Now.ToString("MMdd-HHmmss", CultureInfo.InvariantCulture);
            _sumber  = sumber;
            _periode = periode;
            return _idJalan;
        }

        public static void SelesaiJalan()
        {
            _idJalan = null; _sumber = null; _periode = null;
        }

        public static string IdJalan { get { return _idJalan; } }

        // ---- Lokasi ----
        public static string FolderUtama
        {
            get
            {
                string lib = AppPaths.FolderLibrary;
                return string.IsNullOrEmpty(lib) ? "" : Path.Combine(AppPaths.FolderLogs, "proses");
            }
        }

        public static string FolderCadangan
        {
            get { return Path.Combine(Path.Combine(AppPaths.FolderLokal, "logs"), "proses"); }
        }

        /// <summary>Folder yang dipakai saat ini (untuk tombol "Buka folder").</summary>
        public static string FolderAktif
        {
            get
            {
                try
                {
                    string f = FolderUtama;
                    if (!string.IsNullOrEmpty(f) && Directory.Exists(f)) return f;
                    if (!string.IsNullOrEmpty(f) && Directory.Exists(AppPaths.FolderLogs)) return AppPaths.FolderLogs;
                }
                catch { }
                return FolderCadangan;
            }
        }

        private static string UserAman()
        {
            string user = Environment.UserName ?? "user";
            foreach (char c in Path.GetInvalidFileNameChars()) user = user.Replace(c, '_');
            return user;
        }

        private static string NamaFile(DateTime t)
        {
            return "proses_" + t.ToString("yyyy-MM", CultureInfo.InvariantCulture) + "_" + UserAman() + ".txt";
        }

        // =================================================================
        // Menulis
        // =================================================================
        /// <summary>Awal rincian baru: LogProses.R().Tambah("OS", 123).Tambah("File", path)</summary>
        public static Rincian R() { return new Rincian(); }

        public static void Catat(string proses, string label, string status, Rincian rincian)
        {
            Catat(proses, label, status, rincian == null ? "" : rincian.ToString());
        }

        /// <summary>
        /// Catatan dari modul perhitungan. Di dalam run panel memakai konteks jalan;
        /// di luar run (tombol VBA di Master) sumber = "Tombol Master" dan periode dibaca dari Master!C4.
        /// </summary>
        public static void Catat(string proses, string label, string status, string rincian)
        {
            if (_idJalan != null) Tulis(_sumber ?? "Panel", _periode ?? "", proses, label, status, rincian);
            else Tulis("Tombol Master", PeriodeMaster(), proses, label, status, rincian);
        }

        /// <summary>
        /// Catatan aksi panel (mis. Simpan grup, arsip sheet). Di dalam run memakai konteks
        /// jalan yang sedang aktif; di luar run sumber = "Panel" dengan periode yang diberikan.
        /// </summary>
        public static void CatatPanel(string periode, string proses, string label, string status, Rincian rincian)
        {
            string r = rincian == null ? "" : rincian.ToString();
            if (_idJalan != null) Tulis(_sumber ?? "Panel", _periode ?? periode ?? "", proses, label, status, r);
            else Tulis("Panel", periode ?? "", proses, label, status, r);
        }

        private static void Tulis(string sumber, string periode, string proses, string label, string status, string rincian)
        {
            try
            {
                DateTime kini = DateTime.Now;

                var sb = new StringBuilder();
                sb.Append(kini.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)).Append('\t')
                  .Append(Bersih(Environment.UserName)).Append('\t')
                  .Append(Bersih(sumber)).Append('\t')
                  .Append(Bersih(_idJalan ?? "")).Append('\t')
                  .Append(Bersih(periode)).Append('\t')
                  .Append(Bersih(proses)).Append('\t')
                  .Append(Bersih(label)).Append('\t')
                  .Append(Bersih(status)).Append('\t')
                  .Append(Bersih(rincian))
                  .Append("\r\n");

                string nama = NamaFile(kini);
                lock (_kunci)
                {
                    if (!string.IsNullOrEmpty(FolderUtama) && TulisKe(Path.Combine(FolderUtama, nama), sb.ToString())) return;
                    TulisKe(Path.Combine(FolderCadangan, nama), sb.ToString());
                }
            }
            catch { /* log tidak boleh menggagalkan proses */ }
        }

        private static bool TulisKe(string path, string baris)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                bool baru = !File.Exists(path);
                using (var fs = new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
                using (var w = new StreamWriter(fs, new UTF8Encoding(baru)))   // BOM hanya di file baru → Excel/Notepad mengenali UTF-8
                {
                    if (baru) w.Write(string.Join("\t", Kolom) + "\r\n");
                    w.Write(baris);
                }
                return true;
            }
            catch { return false; }
        }

        /// <summary>Tab/baris baru di nilai akan merusak format satu-baris-satu-kejadian.</summary>
        internal static string Bersih(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("\r\n", " / ").Replace('\n', ' ').Replace('\r', ' ').Replace('\t', ' ').Trim();
        }

        /// <summary>Periode Master!C4 untuk catatan dari tombol VBA (di luar run panel).</summary>
        private static string PeriodeMaster()
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                Excel.Workbook wb = app.ActiveWorkbook;
                DateTime t;
                if (wb != null && Panel.ParameterMaster.BacaTanggalLaporan(wb, out t))
                    return t.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }
            catch { }
            return "";
        }

        // =================================================================
        // Membaca (untuk panel)
        // =================================================================
        private const int BatasEntri = 5000;

        /// <summary>
        /// Semua catatan satu bulan kalender (semua user), terbaru di atas.
        /// </summary>
        /// <param name="bulan">"yyyy-MM"; kosong = bulan terbaru yang ada filenya</param>
        public static Dictionary<string, object> Baca(string bulan)
        {
            var semuaBulan = new SortedSet<string>(StringComparer.Ordinal);
            var file = new List<string>();
            var folder = new List<string>();
            foreach (var f in new[] { FolderUtama, FolderCadangan })
                if (!string.IsNullOrEmpty(f) && Directory.Exists(f) && !folder.Contains(f)) folder.Add(f);

            foreach (var f in folder)
                foreach (var p in Directory.GetFiles(f, "proses_*.txt"))
                {
                    string b = BulanDariNama(Path.GetFileName(p));
                    if (b != null) semuaBulan.Add(b);
                }

            var daftarBulan = new List<string>(semuaBulan);
            daftarBulan.Reverse();
            if (string.IsNullOrEmpty(bulan)) bulan = daftarBulan.Count > 0 ? daftarBulan[0] : DateTime.Now.ToString("yyyy-MM");

            var entri = new List<object>();
            long ukuran = 0;
            foreach (var f in folder)
                foreach (var p in Directory.GetFiles(f, "proses_" + bulan + "_*.txt"))
                {
                    file.Add(Path.GetFileName(p));
                    try { ukuran += new FileInfo(p).Length; } catch { }
                    BacaFile(p, entri);
                }

            // Urut terbaru di atas (kolom waktu berformat yyyy-MM-dd HH:mm:ss → urut teks = urut waktu)
            entri.Sort((a, b) => string.CompareOrdinal(
                (string)((Dictionary<string, object>)b)["waktu"],
                (string)((Dictionary<string, object>)a)["waktu"]));
            bool terpotong = entri.Count > BatasEntri;
            if (terpotong) entri.RemoveRange(BatasEntri, entri.Count - BatasEntri);

            return new Dictionary<string, object>
            {
                { "bulan", bulan }, { "daftarBulan", daftarBulan }, { "entri", entri },
                { "terpotong", terpotong }, { "file", file }, { "ukuran", ukuran },
                { "folder", FolderAktif }
            };
        }

        private static string BulanDariNama(string nama)
        {
            // proses_2026-09_user.txt
            if (nama.Length < 15 || !nama.StartsWith("proses_", StringComparison.OrdinalIgnoreCase)) return null;
            string b = nama.Substring(7, 7);
            DateTime t;
            return DateTime.TryParseExact(b, "yyyy-MM", CultureInfo.InvariantCulture, DateTimeStyles.None, out t) ? b : null;
        }

        private static void BacaFile(string path, List<object> hasil)
        {
            try
            {
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                using (var r = new StreamReader(fs, Encoding.UTF8, true))
                {
                    string baris;
                    while ((baris = r.ReadLine()) != null)
                    {
                        if (baris.Length == 0 || baris.StartsWith("waktu\t", StringComparison.Ordinal)) continue;
                        string[] k = baris.Split('\t');
                        if (k.Length < 8) continue;
                        var d = new Dictionary<string, object>();
                        for (int i = 0; i < Kolom.Length; i++) d[Kolom[i]] = i < k.Length ? k[i] : "";
                        hasil.Add(d);
                    }
                }
            }
            catch { /* file sedang ditulis / tidak bisa dibaca → lewati */ }
        }

        // =================================================================
        // Arsip sheet "Audit Log" lama (sekali jalan, dari panel)
        // =================================================================
        public const string NamaSheetLama = "Audit Log";

        /// <summary>Ukuran isi sheet Audit Log (baris terakhir terisi di kolom A–D, kolom terpakai).</summary>
        public static void UkuranSheet(Excel.Worksheet ws, out int barisAkhir, out int kolomAkhir)
        {
            barisAkhir = 0;
            foreach (var k in new[] { "A", "B", "C", "D" })
            {
                try
                {
                    var sel = (Excel.Range)ws.Cells[ws.Rows.Count, k];
                    int r = ((Excel.Range)sel.End[Excel.XlDirection.xlUp]).Row;
                    if (r > barisAkhir) barisAkhir = r;
                }
                catch { }
            }
            kolomAkhir = 1;
            try
            {
                Excel.Range ur = ws.UsedRange;
                kolomAkhir = ur.Column + ur.Columns.Count - 1;
            }
            catch { }
            if (kolomAkhir > 60) kolomAkhir = 60;     // Audit Log lama memakai s.d. ±30 kolom
            if (barisAkhir == 1)
            {
                // End(xlUp) mengembalikan 1 bila kolom kosong — pastikan A1 benar-benar berisi
                object v = null;
                try { v = ((Excel.Range)ws.Cells[1, 1]).Value2; } catch { }
                if (v == null) barisAkhir = 0;
            }
        }

        /// <summary>
        /// Salin seluruh isi sheet Audit Log ke file teks (TAB) di folder log proses.
        /// Dibaca per blok 2.000 baris agar sheet yang sangat besar tidak menghabiskan memori.
        /// </summary>
        /// <returns>path file arsip</returns>
        public static string ArsipkanSheet(Excel.Worksheet ws, out int jumlahBaris)
        {
            int barisAkhir, kolomAkhir;
            UkuranSheet(ws, out barisAkhir, out kolomAkhir);
            jumlahBaris = barisAkhir;

            string folder = !string.IsNullOrEmpty(FolderUtama) ? FolderUtama : FolderCadangan;
            try { Directory.CreateDirectory(folder); }
            catch { folder = FolderCadangan; Directory.CreateDirectory(folder); }

            string path = Path.Combine(folder,
                "arsip_sheet_audit_log_" + DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) + "_" + UserAman() + ".txt");

            string namaWb = "";
            try { namaWb = ((Excel.Workbook)ws.Parent).Name; } catch { }

            using (var w = new StreamWriter(path, false, new UTF8Encoding(true)))
            {
                w.Write("# Arsip sheet \"Audit Log\" dari " + Bersih(namaWb) +
                        " · diarsipkan " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) +
                        " oleh " + Environment.UserName + " · " + barisAkhir + " baris × " + kolomAkhir + " kolom\r\n");
                const int blok = 2000;
                for (int awal = 1; awal <= barisAkhir; awal += blok)
                {
                    int akhir = Math.Min(barisAkhir, awal + blok - 1);
                    object v = ((Excel.Range)ws.Range[ws.Cells[awal, 1], ws.Cells[akhir, kolomAkhir]]).Value2;
                    var arr = v as object[,];
                    int nBaris = arr == null ? 1 : arr.GetLength(0);
                    for (int i = 1; i <= nBaris; i++)
                    {
                        var sb = new StringBuilder();
                        for (int j = 1; j <= kolomAkhir; j++)
                        {
                            object sel = arr == null ? (j == 1 ? v : null) : arr[i, j];
                            if (j > 1) sb.Append('\t');
                            sb.Append(Bersih(TeksSel(sel, j == 1)));
                        }
                        w.Write(sb.ToString().TrimEnd('\t') + "\r\n");
                    }
                }
            }
            return path;
        }

        private static string TeksSel(object v, bool kolomWaktu)
        {
            if (v == null) return "";
            if (v is int) return "";                          // sel error (#N/A dll.)
            if (v is double)
            {
                double d = (double)v;
                // Kolom A Audit Log berisi timestamp (serial tanggal Excel)
                if (kolomWaktu && d > 30000 && d < 80000)
                {
                    try { return DateTime.FromOADate(d).ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture); }
                    catch { }
                }
                return d.ToString("0.##########", CultureInfo.InvariantCulture);
            }
            return Convert.ToString(v, CultureInfo.InvariantCulture);
        }
    }

    /// <summary>
    /// Pembentuk kolom rincian "Nama=Nilai; Nama=Nilai".
    /// Angka ditulis dengan format Indonesia (1.234.567) supaya mudah dibaca.
    /// </summary>
    internal sealed class Rincian
    {
        private readonly List<string> _isi = new List<string>();

        /// <summary>Nama null/kosong = dilewati (memudahkan isian bersyarat).</summary>
        public Rincian Tambah(string nama, object nilai)
        {
            if (string.IsNullOrEmpty(nama)) return this;
            _isi.Add(Bersih(nama).Replace("=", "-") + "=" + Bersih(Format(nilai)));
            return this;
        }

        /// <summary>Tambah hanya bila syarat terpenuhi.</summary>
        public Rincian TambahBila(bool syarat, string nama, object nilai)
        {
            return syarat ? Tambah(nama, nilai) : this;
        }

        private static string Format(object v)
        {
            if (v == null) return "";
            if (v is double)  return ((double)v).ToString("#,##0.##", LogProses.Id);
            if (v is decimal) return ((decimal)v).ToString("#,##0.##", LogProses.Id);
            if (v is float)   return ((float)v).ToString("#,##0.##", LogProses.Id);
            if (v is int || v is long) return Convert.ToInt64(v).ToString("#,##0", LogProses.Id);
            return Convert.ToString(v, CultureInfo.InvariantCulture);
        }

        // "; " dan "=" di nama, serta "; " di nilai, akan membingungkan pemisah rincian.
        private static string Bersih(string s)
        {
            return LogProses.Bersih(s).Replace("; ", ", ");
        }

        public override string ToString() { return string.Join("; ", _isi.ToArray()); }
    }
}
