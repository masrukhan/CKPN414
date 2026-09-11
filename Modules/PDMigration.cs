using System;
using System.Collections.Generic;
using CKPNLibrary.Helpers;
using CKPNLibrary.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Modules
{
    /// <summary>
    /// Konversi dari VBA modul "module PD Migration".
    /// Menghitung matriks migrasi kualitas piutang antar dua periode triwulan.
    ///
    /// Output: sheet "B2.PD-Migration" baris baseRow..baseRow+4 (5 kualitas)
    /// Staging: sheet "state PD Migration" (dibersihkan lalu diisi ulang)
    /// Log: sheet "Audit Log"
    ///
    /// ----------------------------------------------------------------
    /// REVISI 1 — Normalisasi nomor rekening (KC2900 / Write Off)
    /// ----------------------------------------------------------------
    /// Nomor rekening dinormalisasi via RekHelper sebelum dijadikan kunci,
    /// dan kolom staging B / G dipaksa bertipe Text, agar pencocokan tidak
    /// gagal karena perbedaan tipe sel (Text vs Number) antar file.
    ///
    /// ----------------------------------------------------------------
    /// REVISI 2 — Periode WO tidak lagi memfilter KC2900
    /// ----------------------------------------------------------------
    /// Filter tanggal hapus buku terhadap Master!B40:D43 dihapus. Periode
    /// sudah tersaring secara implisit lewat pasangan file awal-akhir,
    /// sehingga Master!C4 tidak lagi memengaruhi hasil perhitungan.
    ///
    /// ----------------------------------------------------------------
    /// REVISI 3 — Jejak validasi per rekening
    /// ----------------------------------------------------------------
    /// Sheet "state PD Migration" kini memuat tiga blok:
    ///   Kolom B:I  — data awal, hasil lookup, dan data akhir (seperti semula)
    ///   Kolom K:P  — jejak klasifikasi per rekening: ada di akhir? ada di
    ///                KC2900? tanggal & OS hapus buku? klasifikasi akhir?
    ///                cara pencocokan (ketat/longgar)?
    ///   Kolom R:AA — daftar CIF Top-N yang dikecualikan beserta rincian
    ///                rekeningnya
    ///
    /// Blok K:P menjawab pertanyaan "kenapa jumlah rekening di kolom Hapus
    /// Buku berbeda dengan hasil VLOOKUP manual ke KC2900": rekening yang
    /// tercatat di KC2900 TETAPI masih muncul di periode akhir dengan
    /// kualitas 1-5 tidak diklasifikasi WO, melainkan sebagai migrasi biasa.
    /// Selisih tersebut diringkas di Audit Log kolom "Di KC2900 tapi bukan WO".
    /// </summary>
    internal class PDMigration
    {
        private readonly Excel.Application _app;

        // Nama sheet wajib di workbook utama
        private const string SheetState  = "state PD Migration";
        private const string SheetTarget = "B2.PD-Migration";
        private const string SheetLog    = "Audit Log";
        private const string SheetMaster = "Master";

        // Layout KC2900: baris data mulai 3, kolom H=NoRek, I=TglHB, L=BakiDebet
        private const int    WoStartRow  = 3;

        public PDMigration(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException("app");
        }

        // ----------------------------------------------------------------
        // Entry point — dipanggil dari CKPNFunctions
        // Parameter dikirim dari VBA sebagai string
        // ----------------------------------------------------------------
        public void Hitung(
            string triwulan,      // "Triwulan I" / "II" / "III" / "IV"
            string pathAwal,      // path file periode awal
            string pathAkhir,     // path file periode akhir
            string tglAwalStr,    // label periode WO "yyyyMMdd" — Audit Log saja
            string tglAkhirStr,   // label periode WO "yyyyMMdd" — Audit Log saja
            int    topN,          // Top-N CIF yang di-skip
            string sheetKCList)   // "KC0600,KC0700,..."
        {
            var wb = _app.ActiveWorkbook;

            // Validasi sheet wajib
            var wsState  = CariSheet(wb, SheetState);
            var wsTarget = CariSheet(wb, SheetTarget);
            var wsLog    = CariSheet(wb, SheetLog);

            if (wsState  == null) throw new InvalidOperationException("Sheet '" + SheetState  + "' tidak ditemukan.");
            if (wsTarget == null) throw new InvalidOperationException("Sheet '" + SheetTarget + "' tidak ditemukan.");
            if (wsLog    == null) throw new InvalidOperationException("Sheet '" + SheetLog    + "' tidak ditemukan.");

            // Tentukan baseRow dari triwulan
            int baseRow = TriwulanKeBaseRow(triwulan);

            // Validasi file
            if (!System.IO.File.Exists(pathAwal))
                throw new System.IO.FileNotFoundException("File Awal tidak ditemukan: " + pathAwal);
            if (!System.IO.File.Exists(pathAkhir))
                throw new System.IO.FileNotFoundException("File Akhir tidak ditemukan: " + pathAkhir);

            // CATATAN: tglAwalStr / tglAkhirStr TIDAK diparsing menjadi filter.
            // Keduanya diteruskan apa adanya ke Audit Log sebagai label periode.

            // Parse daftar sheet KC
            string[] sheets = sheetKCList.Split(new[]{','}, StringSplitOptions.RemoveEmptyEntries);

            // ---- Bersihkan staging sheet ----
            BersihkanState(wsState);

            // ---- Bangun Top-N (rekening yang masuk CKPN Individu) ----
            var topNInfo = BangunTopN(pathAwal, sheets, topN);

            // ---- Baca file awal & akhir (dengan skip Top-N) ----
            AmbilDariFile(pathAwal,  wsState, "B", "C", "D", sheets, topNInfo.SetSkip);
            AmbilDariFile(pathAkhir, wsState, "G", "H", "I", sheets, topNInfo.SetSkip);

            // ---- Baca WO dari KC2900 file akhir (tanpa filter tanggal) ----
            string diagWO;
            var daftarWO = BacaKC2900(pathAkhir, out diagWO);

            // ---- Bangun indeks periode akhir dari kolom G:I staging ----
            int lastG = CariLastRow(wsState, "G", 2);
            var idxAkhir = new IndeksRekening();
            if (lastG >= 3)
            {
                Excel.Range rngAkhir = (Excel.Range)wsState.Range["G3", "I" + lastG];
                object[,] arrAkhir = (object[,])rngAkhir.Value2;
                for (int k = 1; k <= arrAkhir.GetLength(0); k++)
                {
                    string key = RekHelper.NormNoRek(arrAkhir[k, 1]);
                    if (key.Length == 0) continue;
                    idxAkhir.Tambah(
                        key,
                        new[] { ToDouble(arrAkhir[k, 2]), ToDouble(arrAkhir[k, 3]) },
                        false);   // data pertama menang, tidak diakumulasi
                }
            }

            // ---- Hitung matriks migrasi ----
            int lastB = CariLastRow(wsState, "B", 2);
            double[] totalSaldoAwal = new double[6];   // index 1..5
            double[,] matriks = new double[6, 8];      // [kualAwal 1..5][kolom 1..7]

            var stat = new StatistikKlasifikasi();

            if (lastB >= 3)
            {
                Excel.Range rngAwal = (Excel.Range)wsState.Range["B3", "D" + lastB];
                object[,] arrAwal = (object[,])rngAwal.Value2;
                int nRows = arrAwal.GetLength(0);

                object[,] efVals  = new object[nRows, 2];   // kolom E:F
                object[,] detail  = new object[nRows, 6];   // kolom K:P

                for (int k = 1; k <= nRows; k++)
                {
                    int r = k - 1;

                    string key       = RekHelper.NormNoRek(arrAwal[k, 1]);
                    int    kualAwal  = (int)ToDouble(arrAwal[k, 2]);
                    double saldoAwal = ToDouble(arrAwal[k, 3]);

                    // ---- Lookup ke KC2900 (selalu dilakukan, untuk jejak) ----
                    double[] recWO;
                    string   caraWO;
                    bool     adaWO = daftarWO.Cari(key, out recWO, out caraWO);

                    // ---- Lookup ke periode akhir ----
                    double[] recAkhir;
                    string   caraAkhir;
                    bool     adaAkhir = idxAkhir.Cari(key, out recAkhir, out caraAkhir);

                    string klasifikasi;

                    if (kualAwal >= 1 && kualAwal <= 5)
                    {
                        totalSaldoAwal[kualAwal] += saldoAwal;

                        if (adaAkhir)
                        {
                            int    kualAkhir  = (int)recAkhir[0];
                            double saldoAkhir = recAkhir[1];
                            efVals[r, 0] = kualAkhir;
                            efVals[r, 1] = saldoAkhir;

                            if (kualAkhir >= 1 && kualAkhir <= 5)
                            {
                                matriks[kualAwal, kualAkhir] += saldoAwal;
                                klasifikasi = "Kol " + kualAkhir;
                                stat.RekBertahan++;
                            }
                            else if (adaWO)
                            {
                                matriks[kualAwal, 6] += saldoAwal;
                                klasifikasi = "Hapus Buku";
                                stat.CatatWO(kualAwal, saldoAwal);
                            }
                            else
                            {
                                matriks[kualAwal, 7] += saldoAwal;
                                klasifikasi = "Lainnya";
                                stat.CatatLainnya(saldoAwal);
                            }
                        }
                        else if (adaWO)
                        {
                            // Ada di awal, hilang di akhir, tercatat di KC2900
                            // => hapus buku dalam periode berjalan
                            efVals[r, 0] = "WO";
                            efVals[r, 1] = recWO[0];
                            matriks[kualAwal, 6] += saldoAwal;
                            klasifikasi = "Hapus Buku";
                            stat.CatatWO(kualAwal, saldoAwal);
                        }
                        else
                        {
                            efVals[r, 0] = 0;
                            efVals[r, 1] = 0;
                            matriks[kualAwal, 7] += saldoAwal;
                            klasifikasi = "Lainnya";
                            stat.CatatLainnya(saldoAwal);
                        }
                    }
                    else
                    {
                        efVals[r, 0] = 0;
                        efVals[r, 1] = 0;
                        klasifikasi  = "Diabaikan (kualitas awal tidak valid)";
                        stat.RekDiabaikan++;
                    }

                    // ---- Selisih terhadap VLOOKUP manual ke KC2900 ----
                    // Rekening yang ADA di KC2900 tetapi TIDAK diklasifikasi
                    // sebagai Hapus Buku. Inilah penyebab jumlah hasil VLOOKUP
                    // manual lebih besar daripada isi kolom H PD-Migration.
                    if (adaWO && klasifikasi != "Hapus Buku")
                        stat.CatatWOTakTerpakai(kualAwal, saldoAwal);

                    // ---- Tulis jejak validasi kolom K:P ----
                    detail[r, 0] = adaAkhir ? "Ya" : "Tidak";                       // K
                    detail[r, 1] = adaWO    ? "Ya" : "Tidak";                       // L
                    detail[r, 2] = adaWO && recWO[1] > 0 ? (object)recWO[1] : "-";  // M
                    detail[r, 3] = adaWO ? (object)recWO[0] : "-";                  // N
                    detail[r, 4] = klasifikasi;                                     // O
                    detail[r, 5] = adaWO ? caraWO : "-";                            // P
                }

                // Tulis blok sekaligus — jauh lebih cepat daripada per sel
                TulisBlok(wsState, "E", "F", 3, efVals);
                TulisBlok(wsState, "K", "P", 3, detail);

                ((Excel.Range)wsState.Range["D3", "D" + lastB]).NumberFormat = "#,##0;(#,##0);-";
                ((Excel.Range)wsState.Range["F3", "F" + lastB]).NumberFormat = "#,##0;(#,##0);-";
                ((Excel.Range)wsState.Range["M3", "M" + lastB]).NumberFormat = "0";
                ((Excel.Range)wsState.Range["N3", "N" + lastB]).NumberFormat = "#,##0;(#,##0);-";
            }
            if (lastG >= 3)
                ((Excel.Range)wsState.Range["I3", "I" + lastG]).NumberFormat = "#,##0;(#,##0);-";

            // ---- Tulis rincian Top-N ke staging kolom R:AA ----
            TulisDetailTopN(wsState, topNInfo);

            // ---- Tulis matriks ke B2.PD-Migration ----
            for (int rowK = 1; rowK <= 5; rowK++)
            {
                ((Excel.Range)wsTarget.Cells[baseRow + rowK - 1, "B"]).Value2 = totalSaldoAwal[rowK];
                ((Excel.Range)wsTarget.Cells[baseRow + rowK - 1, "C"]).Value2 = matriks[rowK, 1];
                ((Excel.Range)wsTarget.Cells[baseRow + rowK - 1, "D"]).Value2 = matriks[rowK, 2];
                ((Excel.Range)wsTarget.Cells[baseRow + rowK - 1, "E"]).Value2 = matriks[rowK, 3];
                ((Excel.Range)wsTarget.Cells[baseRow + rowK - 1, "F"]).Value2 = matriks[rowK, 4];
                ((Excel.Range)wsTarget.Cells[baseRow + rowK - 1, "G"]).Value2 = matriks[rowK, 5];
                ((Excel.Range)wsTarget.Cells[baseRow + rowK - 1, "H"]).Value2 = matriks[rowK, 6];
                ((Excel.Range)wsTarget.Cells[baseRow + rowK - 1, "I"]).Value2 = matriks[rowK, 7];
            }

            // ---- Audit Log ----
            TulisAuditLog(wsLog, triwulan, tglAwalStr, tglAkhirStr,
                          pathAwal, pathAkhir,
                          lastB - 2, lastG - 2,
                          daftarWO, totalSaldoAwal, matriks,
                          topN, topNInfo, sheetKCList, diagWO, stat);

            // ---- Pesan selesai ----
            System.Windows.Forms.MessageBox.Show(
                "Selesai.\n" +
                triwulan + " -> B2.PD-Migration baris " + baseRow + ":" + (baseRow + 4) + "\n\n" +
                "Top-N        : " + topN + " diminta, " + topNInfo.DaftarCIF.Count + " CIF terpakai\n" +
                "Rek Individu : " + topNInfo.JumlahRekSkip + " rekening\n" +
                "OS Individu  : " + topNInfo.OSTopN.ToString("N0") + "\n" +
                "OS Bruto     : " + topNInfo.OSBruto.ToString("N0") + "\n\n" +
                "Baris Awal   : " + (lastB - 2) + "\n" +
                "Baris Akhir  : " + (lastG - 2) + "\n\n" +
                "Rek Bertahan : " + stat.RekBertahan + "\n" +
                "Rek WO       : " + stat.RekWO + " (" + stat.NominalWO.ToString("N0") + ")\n" +
                "Rek Lainnya  : " + stat.RekLainnya + " (" + stat.NominalLainnya.ToString("N0") + ")",
                "PD Migration", System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
        }

        // ================================================================
        // StatistikKlasifikasi: penghitung untuk Audit Log
        // ================================================================
        private class StatistikKlasifikasi
        {
            public int    RekBertahan;
            public int    RekDiabaikan;

            public int    RekWO;
            public double NominalWO;
            public readonly int[]    WOPerKol     = new int[6];
            public readonly double[] WONominalKol = new double[6];

            public int    RekLainnya;
            public double NominalLainnya;

            // Ada di KC2900 tetapi TIDAK diklasifikasi Hapus Buku
            public int    RekWOTakTerpakai;
            public double NominalWOTakTerpakai;
            public readonly int[]    TakTerpakaiPerKol     = new int[6];
            public readonly double[] TakTerpakaiNominalKol = new double[6];

            public void CatatWO(int kualAwal, double saldo)
            {
                RekWO++;
                NominalWO += saldo;
                if (kualAwal >= 1 && kualAwal <= 5)
                {
                    WOPerKol[kualAwal]++;
                    WONominalKol[kualAwal] += saldo;
                }
            }

            public void CatatLainnya(double saldo)
            {
                RekLainnya++;
                NominalLainnya += saldo;
            }

            public void CatatWOTakTerpakai(int kualAwal, double saldo)
            {
                RekWOTakTerpakai++;
                NominalWOTakTerpakai += saldo;
                if (kualAwal >= 1 && kualAwal <= 5)
                {
                    TakTerpakaiPerKol[kualAwal]++;
                    TakTerpakaiNominalKol[kualAwal] += saldo;
                }
            }

            public string RincianWO()          { return Rincian(WOPerKol, WONominalKol); }
            public string RincianWOTakTerpakai(){ return Rincian(TakTerpakaiPerKol, TakTerpakaiNominalKol); }

            private static string Rincian(int[] jml, double[] nominal)
            {
                var sb = new System.Text.StringBuilder();
                for (int i = 1; i <= 5; i++)
                {
                    if (sb.Length > 0) sb.Append("; ");
                    sb.Append("Kol").Append(i).Append("=").Append(jml[i])
                      .Append(" (").Append(nominal[i].ToString("N0")).Append(")");
                }
                return sb.ToString();
            }
        }

        // ================================================================
        // Hasil Top-N
        // ================================================================
        private class RekTopN
        {
            public string SheetKC;
            public string NoRek;
            public double OS;
        }

        private class CifTopN
        {
            public string CIF;
            public double TotalOS;
            public readonly List<RekTopN> Rekening = new List<RekTopN>();
        }

        private class HasilTopN
        {
            public readonly HashSet<string> SetSkip =
                new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            public readonly List<CifTopN> DaftarCIF = new List<CifTopN>();

            /// <summary>Total OS seluruh CIF di file awal (sebelum dikurangi Top-N).</summary>
            public double OSBruto;

            /// <summary>Total OS CIF Top-N (yang masuk CKPN Individu).</summary>
            public double OSTopN;

            /// <summary>Jumlah CIF unik di file awal.</summary>
            public int JumlahCIF;

            /// <summary>Jumlah rekening yang dikecualikan.</summary>
            public int JumlahRekSkip;
        }

        // ================================================================
        // Helpers
        // ================================================================

        private static int TriwulanKeBaseRow(string triwulan)
        {
            switch (triwulan.Trim())
            {
                case "Triwulan I":   return 4;
                case "Triwulan II":  return 23;
                case "Triwulan III": return 42;
                case "Triwulan IV":  return 61;
                default: throw new ArgumentException("Triwulan tidak dikenali: " + triwulan);
            }
        }

        // ----------------------------------------------------------------
        // BersihkanState: kosongkan seluruh blok lalu tulis ulang header
        // Sheet ini bersifat staging — dibersihkan setiap kali perhitungan.
        // ----------------------------------------------------------------
        private void BersihkanState(Excel.Worksheet ws)
        {
            BersihkanKolom(ws, "B",  "D");
            BersihkanKolom(ws, "G",  "I");
            BersihkanKolom(ws, "E",  "F");
            BersihkanKolom(ws, "K",  "P");    // jejak validasi
            BersihkanKolom(ws, "R",  "AA");   // rincian Top-N

            // ---- Kolom nomor rekening / CIF WAJIB bertipe Text ----
            // Tanpa ini, Excel meng-coerce "0060012345" menjadi 60012345
            // sehingga leading zero hilang dan pencocokan gagal.
            ((Excel.Range)ws.Range["B:B"]).NumberFormat  = "@";
            ((Excel.Range)ws.Range["G:G"]).NumberFormat  = "@";
            ((Excel.Range)ws.Range["S:S"]).NumberFormat  = "@";
            ((Excel.Range)ws.Range["Y:Y"]).NumberFormat  = "@";
            ((Excel.Range)ws.Range["Z:Z"]).NumberFormat  = "@";

            // ---- Blok 1: data awal, lookup, data akhir ----
            ((Excel.Range)ws.Range["B1"]).Value2 = "Triwulan Awal";
            ((Excel.Range)ws.Range["G1"]).Value2 = "Triwulan Akhir";
            ((Excel.Range)ws.Range["E1"]).Value2 = "VLOOKUP DATA";
            ((Excel.Range)ws.Range["B2"]).Value2 = "Nomor Rekening";
            ((Excel.Range)ws.Range["C2"]).Value2 = "Kualitas";
            ((Excel.Range)ws.Range["D2"]).Value2 = "Saldo Harga Pokok";
            ((Excel.Range)ws.Range["E2"]).Value2 = "Kualitas";
            ((Excel.Range)ws.Range["F2"]).Value2 = "Saldo Harga Pokok";
            ((Excel.Range)ws.Range["G2"]).Value2 = "Nomor Rekening";
            ((Excel.Range)ws.Range["H2"]).Value2 = "Kualitas";
            ((Excel.Range)ws.Range["I2"]).Value2 = "Saldo Harga Pokok";

            // ---- Blok 2: jejak validasi hapus buku (sejajar baris blok 1) ----
            ((Excel.Range)ws.Range["K1"]).Value2 = "VALIDASI HAPUS BUKU (KC2900)";
            ((Excel.Range)ws.Range["K2"]).Value2 = "Ada di Akhir";
            ((Excel.Range)ws.Range["L2"]).Value2 = "Ada di KC2900";
            ((Excel.Range)ws.Range["M2"]).Value2 = "Tgl Hapus Buku";
            ((Excel.Range)ws.Range["N2"]).Value2 = "OS KC2900";
            ((Excel.Range)ws.Range["O2"]).Value2 = "Klasifikasi Akhir";
            ((Excel.Range)ws.Range["P2"]).Value2 = "Cara Cocok KC2900";

            // ---- Blok 3: rincian Top-N yang dikecualikan ----
            ((Excel.Range)ws.Range["R1"]).Value2 = "TOP-N CKPN INDIVIDU (DIKECUALIKAN)";
            ((Excel.Range)ws.Range["R2"]).Value2 = "No";
            ((Excel.Range)ws.Range["S2"]).Value2 = "CIF";
            ((Excel.Range)ws.Range["T2"]).Value2 = "Total OS CIF";
            ((Excel.Range)ws.Range["U2"]).Value2 = "Jml Rekening";

            ((Excel.Range)ws.Range["W1"]).Value2 = "RINCIAN REKENING TOP-N";
            ((Excel.Range)ws.Range["W2"]).Value2 = "No CIF";
            ((Excel.Range)ws.Range["X2"]).Value2 = "Sheet KC";
            ((Excel.Range)ws.Range["Y2"]).Value2 = "CIF";
            ((Excel.Range)ws.Range["Z2"]).Value2 = "No Rekening";
            ((Excel.Range)ws.Range["AA2"]).Value2 = "OS";
        }

        private void BersihkanKolom(Excel.Worksheet ws, string kolAwal, string kolAkhir)
        {
            int lastA = CariLastRow(ws, kolAwal,  2);
            int lastB = CariLastRow(ws, kolAkhir, 2);
            int last  = Math.Max(lastA, lastB);
            if (last >= 3)
                ((Excel.Range)ws.Range[kolAwal + "3", kolAkhir + last]).ClearContents();
        }

        // ----------------------------------------------------------------
        // TulisDetailTopN: isi blok R:AA di sheet staging
        // ----------------------------------------------------------------
        private void TulisDetailTopN(Excel.Worksheet ws, HasilTopN info)
        {
            if (info.DaftarCIF.Count == 0) return;

            // Ringkasan per CIF (R:U)
            object[,] ringkas = new object[info.DaftarCIF.Count, 4];
            for (int i = 0; i < info.DaftarCIF.Count; i++)
            {
                var c = info.DaftarCIF[i];
                ringkas[i, 0] = i + 1;
                ringkas[i, 1] = c.CIF;
                ringkas[i, 2] = c.TotalOS;
                ringkas[i, 3] = c.Rekening.Count;
            }
            TulisBlok(ws, "R", "U", 3, ringkas);
            ((Excel.Range)ws.Range["T3", "T" + (2 + info.DaftarCIF.Count)])
                .NumberFormat = "#,##0;(#,##0);-";

            // Rincian per rekening (W:AA)
            int totalRek = 0;
            foreach (var c in info.DaftarCIF) totalRek += c.Rekening.Count;
            if (totalRek == 0) return;

            object[,] rinci = new object[totalRek, 5];
            int baris = 0;
            for (int i = 0; i < info.DaftarCIF.Count; i++)
            {
                var c = info.DaftarCIF[i];
                foreach (var rk in c.Rekening)
                {
                    rinci[baris, 0] = i + 1;
                    rinci[baris, 1] = rk.SheetKC;
                    rinci[baris, 2] = c.CIF;
                    rinci[baris, 3] = rk.NoRek;
                    rinci[baris, 4] = rk.OS;
                    baris++;
                }
            }
            TulisBlok(ws, "W", "AA", 3, rinci);
            ((Excel.Range)ws.Range["AA3", "AA" + (2 + totalRek)])
                .NumberFormat = "#,##0;(#,##0);-";
        }

        // Baca file sumber -> tulis ke kolom staging (setara AmbilDariFile_Dinamis VBA)
        private void AmbilDariFile(
            string filePath, Excel.Worksheet wsOut,
            string kolRek, string kolKual, string kolPok,
            string[] sheets, HashSet<string> setSkip)
        {
            Excel.Workbook wbSrc = null;
            try
            {
                wbSrc = _app.Workbooks.Open(filePath, UpdateLinks: 0, ReadOnly: true);
                int outRow = 3;

                foreach (var shName in sheets)
                {
                    SheetSpec spec;
                    try { spec = SheetSpec.Buat(shName); }
                    catch { continue; }

                    if (!ExcelHelper.SheetAda(wbSrc, shName)) continue;
                    var ws = (Excel.Worksheet)wbSrc.Worksheets[shName];

                    int headerRow = spec.HeaderRow;
                    // Gunakan ColCIF (kolom C) untuk cari lastRow — lebih konsisten
                    // karena kolom C (CIF) selalu terisi untuk setiap baris debitur
                    int lastRow   = ExcelHelper.CariLastRow(ws, spec.ColCIF, headerRow);
                    if (lastRow <= headerRow) continue;

                    int dataStart = headerRow + 1;
                    string[] namaArr  = ExcelHelper.BacaKolomString(ws, spec.ColNama,     dataStart, lastRow);
                    string[] rekArr   = ExcelHelper.BacaKolomString(ws, spec.ColNoRek,    dataStart, lastRow);
                    string[] kualArr  = ExcelHelper.BacaKolomString(ws, spec.ColKualitas, dataStart, lastRow);

                    double[] h1Arr, n1Arr;
                    double[] h2Arr = null, n2Arr = null;

                    switch (spec.Mode)
                    {
                        case SheetMode.AF:
                            n1Arr = ExcelHelper.BacaKolomDouble(ws, spec.ColOS,     dataStart, lastRow);
                            h1Arr = ExcelHelper.BacaKolomDouble(ws, spec.ColHariAF, dataStart, lastRow);
                            break;
                        case SheetMode.Tunggakan1:
                        {
                            // KC1100 (ijarah): OS/EAD staging = tunggakan pokok (AL) + ujroh (AM).
                            // Ujroh digabung ke Nom1 agar cabang osVal (Tunggakan1) tetap = n1Arr[i].
                            // Konsisten dengan EAD di CKPN Individu & PD Net Flow.
                            h1Arr = ExcelHelper.BacaKolomDouble(ws, spec.ColHariPokok, dataStart, lastRow);
                            n1Arr = ExcelHelper.BacaKolomDouble(ws, spec.ColNomPokok,  dataStart, lastRow);
                            double[] ujrohArr = ExcelHelper.BacaKolomDouble(ws, spec.ColNomUjroh, dataStart, lastRow);
                            int mUj = Math.Min(n1Arr.Length, ujrohArr.Length);
                            for (int k = 0; k < mUj; k++) n1Arr[k] += ujrohArr[k];
                            break;
                        }
                        default: // Tunggakan2
                            h1Arr = ExcelHelper.BacaKolomDouble(ws, spec.ColHariPokok, dataStart, lastRow);
                            n1Arr = ExcelHelper.BacaKolomDouble(ws, spec.ColNomPokok,  dataStart, lastRow);
                            h2Arr = ExcelHelper.BacaKolomDouble(ws, spec.ColHariBasil, dataStart, lastRow);
                            n2Arr = ExcelHelper.BacaKolomDouble(ws, spec.ColNomBasil,  dataStart, lastRow);
                            break;
                    }

                    for (int i = 0; i < namaArr.Length; i++)
                    {
                        // Filter utama: baris harus punya nama (bukan baris agunan)
                        if (string.IsNullOrEmpty(namaArr[i])) continue;

                        double osVal;
                        switch (spec.Mode)
                        {
                            case SheetMode.AF:
                            case SheetMode.Tunggakan1:
                                // KC0600-KC0900 & KC1100: OS tanpa syarat hari tunggakan
                                osVal = n1Arr[i];
                                break;
                            default: // KC1000
                                // OS = AL + AN tanpa syarat hari tunggakan
                                osVal = n1Arr[i] + (n2Arr != null ? n2Arr[i] : 0);
                                break;
                        }

                        // Normalisasi nomor rekening — tahan Text maupun Number
                        string rek = RekHelper.Bersihkan(rekArr[i]);
                        if (rek.Length == 0) continue;

                        // Skip Top-N jika berlaku. Diterapkan pada KEDUA periode
                        // agar populasi awal & akhir sebanding saat direview.
                        if (setSkip != null && setSkip.Contains(shName + "|" + rek)) continue;

                        int kual = NormKualitas(kualArr[i]);

                        ((Excel.Range)wsOut.Cells[outRow, kolRek ]).Value2 = rek;
                        ((Excel.Range)wsOut.Cells[outRow, kolKual]).Value2 = kual > 0 ? (object)kual : (object)kualArr[i];
                        ((Excel.Range)wsOut.Cells[outRow, kolPok ]).Value2 = osVal;
                        outRow++;
                    }
                }
            }
            finally
            {
                if (wbSrc != null) try { wbSrc.Close(false); } catch { }
            }
        }

        // ----------------------------------------------------------------
        // BacaKC2900: baca daftar Write Off dari file periode akhir
        //
        // CATATAN METODOLOGI — TIDAK ADA FILTER TANGGAL:
        //   Penyaringan periode sudah terjadi secara implisit lewat pasangan
        //   file awal-akhir. Sebuah rekening baru diklasifikasi WO jika:
        //     (a) punya saldo di periode awal, DAN
        //     (b) tidak ditemukan lagi di periode akhir dengan kualitas 1-5, DAN
        //     (c) tercatat di KC2900 file akhir.
        //   Kolom I (TglHB) tetap dibaca HANYA untuk ditampilkan sebagai
        //   rentang diagnostik dan jejak per rekening, bukan kriteria seleksi.
        // ----------------------------------------------------------------
        private IndeksRekening BacaKC2900(string filePath, out string diag)
        {
            var idx = new IndeksRekening();
            diag = "Sheet KC2900 tidak ditemukan di file akhir.";

            if (!System.IO.File.Exists(filePath))
            {
                diag = "File periode akhir tidak ditemukan.";
                return idx;
            }

            int  nTeks = 0, nAngka = 0, nRekKosong = 0, nTglTakValid = 0;
            long tglMin = 0, tglMax = 0;

            Excel.Workbook wbSrc = null;
            try
            {
                wbSrc = _app.Workbooks.Open(filePath, UpdateLinks: 0, ReadOnly: true);
                if (!ExcelHelper.SheetAda(wbSrc, "KC2900")) return idx;

                var ws      = (Excel.Worksheet)wbSrc.Worksheets["KC2900"];
                int lastRow = ExcelHelper.CariLastRow(ws, "H", WoStartRow - 1);
                if (lastRow < WoStartRow)
                {
                    diag = "KC2900 ada tetapi kolom H kosong.";
                    return idx;
                }

                Excel.Range rng = (Excel.Range)ws.Range["H" + WoStartRow, "L" + lastRow];
                object[,]   arr = (object[,])rng.Value2;

                for (int i = 1; i <= arr.GetLength(0); i++)
                {
                    object rawRek = arr[i, 1];
                    string noRek  = RekHelper.NormNoRek(rawRek);
                    if (noRek.Length == 0) { nRekKosong++; continue; }

                    // Statistik tipe sel — bukti validitas untuk Audit Log
                    if (rawRek is string) nTeks++; else nAngka++;

                    // Tanggal HB dibaca untuk pelaporan saja — TIDAK memfilter
                    long tglWO = RekHelper.NormYyyymmdd(arr[i, 2]);
                    if (tglWO == 0) nTglTakValid++;
                    else
                    {
                        if (tglMin == 0 || tglWO < tglMin) tglMin = tglWO;
                        if (tglWO > tglMax)                tglMax = tglWO;
                    }

                    double osVal = ToDouble(arr[i, 5]);   // kolom L = indeks 5 dari H
                    idx.Tambah(noRek, new[] { osVal, (double)tglWO }, true);
                }

                diag = "Baris dibaca=" + arr.GetLength(0) +
                       "; rek valid (teks=" + nTeks + ", angka=" + nAngka + ")" +
                       "; rek kosong=" + nRekKosong +
                       "; tgl tidak valid=" + nTglTakValid +
                       "; rentang TglHB=" + (tglMin == 0 ? "-" : tglMin + " s/d " + tglMax) +
                       "; WO unik=" + idx.Count + " (tanpa filter tanggal)";
            }
            finally
            {
                if (wbSrc != null) try { wbSrc.Close(false); } catch { }
            }
            return idx;
        }

        // ----------------------------------------------------------------
        // BangunTopN: identifikasi CIF Top-N yang dikecualikan (masuk CKPN
        // Individu) beserta seluruh rekeningnya dan nominal OS-nya.
        //
        // Menggantikan pasangan BangunSetSkipTopN + HitungOSSetSkip pada
        // versi lama — OS Individu kini dihitung dari agregasi yang sama,
        // sehingga file awal cukup dibuka SATU kali.
        //
        // Ketentuan yang dipertahankan:
        //   1. Ranking LINTAS semua sheet KC (bukan per sheet terpisah)
        //      -> konsisten dengan logika CKPNIndividu.cs
        //   2. KC1000/KC1100: masuk jika nama debitur ada, OS tanpa syarat
        //      hari tunggakan
        //   3. Nomor rekening dinormalisasi via RekHelper.Bersihkan agar
        //      kunci "sheet|rek" identik dengan yang dipakai AmbilDariFile
        // ----------------------------------------------------------------
        private HasilTopN BangunTopN(string filePath, string[] sheets, int topN)
        {
            var hasil = new HasilTopN();
            if (!System.IO.File.Exists(filePath)) return hasil;

            Excel.Workbook wbSrc = null;
            try
            {
                wbSrc = _app.Workbooks.Open(filePath, UpdateLinks: 0, ReadOnly: true);

                // -- Langkah 1: Agregasi OS per CIF LINTAS SEMUA SHEET --
                var dictCIF = new Dictionary<string, CifTopN>(StringComparer.OrdinalIgnoreCase);

                foreach (var shName in sheets)
                {
                    SheetSpec spec;
                    try { spec = SheetSpec.Buat(shName); } catch { continue; }
                    if (!ExcelHelper.SheetAda(wbSrc, shName)) continue;

                    var ws      = (Excel.Worksheet)wbSrc.Worksheets[shName];
                    int lastRow = ExcelHelper.CariLastRow(ws, spec.ColCIF, spec.HeaderRow);
                    if (lastRow <= spec.HeaderRow) continue;

                    int dataStart5 = spec.HeaderRow + 1;
                    string[] namaArr = ExcelHelper.BacaKolomString(ws, spec.ColNama,  dataStart5, lastRow);
                    string[] cifArr  = ExcelHelper.BacaKolomString(ws, spec.ColCIF,   dataStart5, lastRow);
                    string[] rekArr  = ExcelHelper.BacaKolomString(ws, spec.ColNoRek, dataStart5, lastRow);
                    double[] n1Arr, n2Arr = null;

                    switch (spec.Mode)
                    {
                        case SheetMode.AF:
                            n1Arr = ExcelHelper.BacaKolomDouble(ws, spec.ColOS,        dataStart5, lastRow);
                            break;
                        case SheetMode.Tunggakan1:
                        {
                            // KC1100 (ijarah): ranking OS = pokok (AL) + ujroh (AM) agar Top-N
                            // konsisten dengan CKPN Individu & PD Net Flow (kini pokok+ujroh).
                            n1Arr = ExcelHelper.BacaKolomDouble(ws, spec.ColNomPokok,  dataStart5, lastRow);
                            double[] ujrohArr = ExcelHelper.BacaKolomDouble(ws, spec.ColNomUjroh, dataStart5, lastRow);
                            int mUj = Math.Min(n1Arr.Length, ujrohArr.Length);
                            for (int k = 0; k < mUj; k++) n1Arr[k] += ujrohArr[k];
                            break;
                        }
                        default: // KC1000
                            n1Arr = ExcelHelper.BacaKolomDouble(ws, spec.ColNomPokok,  dataStart5, lastRow);
                            n2Arr = ExcelHelper.BacaKolomDouble(ws, spec.ColNomBasil,  dataStart5, lastRow);
                            break;
                    }

                    for (int i = 0; i < namaArr.Length; i++)
                    {
                        // Filter: harus ada nama debitur (bukan baris agunan)
                        if (string.IsNullOrEmpty(namaArr[i])) continue;

                        double osVal;
                        switch (spec.Mode)
                        {
                            case SheetMode.AF:
                            case SheetMode.Tunggakan1:
                                osVal = n1Arr[i]; break;
                            default: // KC1000
                                osVal = n1Arr[i] + (n2Arr != null ? n2Arr[i] : 0); break;
                        }

                        string cif = cifArr[i].Trim();
                        string rek = RekHelper.Bersihkan(rekArr[i]);
                        if (string.IsNullOrEmpty(cif)) continue;

                        CifTopN c;
                        if (!dictCIF.TryGetValue(cif, out c))
                        {
                            c = new CifTopN { CIF = cif, TotalOS = 0 };
                            dictCIF[cif] = c;
                        }
                        c.TotalOS += osVal;
                        c.Rekening.Add(new RekTopN { SheetKC = shName, NoRek = rek, OS = osVal });

                        hasil.OSBruto += osVal;
                    }
                }

                hasil.JumlahCIF = dictCIF.Count;
                if (dictCIF.Count == 0 || topN <= 0) return hasil;

                // -- Langkah 2: Sort descending lintas sheet, ambil Top-N CIF --
                var pairs = new List<SortHelper.CifOsPair>();
                foreach (var kv in dictCIF)
                    pairs.Add(new SortHelper.CifOsPair { CIF = kv.Key, OS = kv.Value.TotalOS });
                SortHelper.SortDescByOS(pairs, 0, pairs.Count - 1);

                int ambil = Math.Min(topN, pairs.Count);

                // -- Langkah 3: Kumpulkan rekening & nominal Top-N --
                for (int t = 0; t < ambil; t++)
                {
                    var c = dictCIF[pairs[t].CIF];
                    hasil.DaftarCIF.Add(c);
                    hasil.OSTopN += c.TotalOS;
                    foreach (var rk in c.Rekening)
                        hasil.SetSkip.Add(rk.SheetKC + "|" + rk.NoRek);
                }
                hasil.JumlahRekSkip = hasil.SetSkip.Count;
            }
            finally
            {
                if (wbSrc != null) try { wbSrc.Close(false); } catch { }
            }
            return hasil;
        }

        // ----------------------------------------------------------------
        // TulisAuditLog: catat ringkasan perhitungan ke sheet Audit Log
        //
        // tglAwalStr / tglAkhirStr dicatat sebagai LABEL periode (Master
        // B40:D43). Keduanya tidak lagi memengaruhi hasil perhitungan.
        // ----------------------------------------------------------------
        private void TulisAuditLog(
            Excel.Worksheet wsLog,
            string triwulan, string tglAwalStr, string tglAkhirStr,
            string pathAwal, string pathAkhir,
            int nAwal, int nAkhir,
            IndeksRekening daftarWO,
            double[] totalSaldoAwal, double[,] matriks,
            int topN, HasilTopN topNInfo, string sheetKCList,
            string diagWO, StatistikKlasifikasi stat)
        {
            const int summaryStart = 5;
            const int summaryMax   = 200;

            int nextRow = summaryStart;
            while (nextRow <= summaryMax)
            {
                string a = ToStr(((Excel.Range)wsLog.Cells[nextRow, 1]).Value2);
                string b = ToStr(((Excel.Range)wsLog.Cells[nextRow, 2]).Value2);
                if (string.IsNullOrEmpty(a) && string.IsNullOrEmpty(b)) break;
                nextRow++;
            }

            double ts = DateTime.Now.ToOADate();
            double totalSA = 0, totalWO = 0;
            for (int i = 1; i <= 5; i++) { totalSA += totalSaldoAwal[i]; totalWO += matriks[i, 6]; }

            // Konversi tglAwalStr / tglAkhirStr dari "yyyyMMdd" ke OADate
            double oaAwal  = ParseTglToOADate(tglAwalStr);
            double oaAkhir = ParseTglToOADate(tglAkhirStr);

            ((Excel.Range)wsLog.Cells[nextRow,  1]).Value2 = ts;
            ((Excel.Range)wsLog.Cells[nextRow,  1]).NumberFormat = "m/d/yyyy h:mm";
            ((Excel.Range)wsLog.Cells[nextRow,  2]).Value2 = "PD Migration";
            ((Excel.Range)wsLog.Cells[nextRow,  3]).Value2 = triwulan;
            ((Excel.Range)wsLog.Cells[nextRow,  4]).Value2 = oaAwal  > 0 ? oaAwal  : (object)tglAwalStr;
            ((Excel.Range)wsLog.Cells[nextRow,  4]).NumberFormat = oaAwal  > 0 ? "m/d/yyyy" : "@";
            ((Excel.Range)wsLog.Cells[nextRow,  5]).Value2 = oaAkhir > 0 ? oaAkhir : (object)tglAkhirStr;
            ((Excel.Range)wsLog.Cells[nextRow,  5]).NumberFormat = oaAkhir > 0 ? "m/d/yyyy" : "@";
            ((Excel.Range)wsLog.Cells[nextRow,  6]).Value2 = pathAwal;
            ((Excel.Range)wsLog.Cells[nextRow,  7]).Value2 = pathAkhir;
            ((Excel.Range)wsLog.Cells[nextRow,  8]).Value2 = "OK";
            ((Excel.Range)wsLog.Cells[nextRow,  9]).Value2 = nAwal;
            ((Excel.Range)wsLog.Cells[nextRow, 10]).Value2 = nAkhir;
            ((Excel.Range)wsLog.Cells[nextRow, 11]).Value2 = daftarWO.Count;
            ((Excel.Range)wsLog.Cells[nextRow, 12]).Value2 = totalSA;
            ((Excel.Range)wsLog.Cells[nextRow, 12]).NumberFormat = "#,##0;(#,##0);-";
            ((Excel.Range)wsLog.Cells[nextRow, 13]).Value2 = topN;
            ((Excel.Range)wsLog.Cells[nextRow, 14]).Value2 = topNInfo.JumlahRekSkip;
            ((Excel.Range)wsLog.Cells[nextRow, 15]).Value2 = topNInfo.OSTopN;
            ((Excel.Range)wsLog.Cells[nextRow, 15]).NumberFormat = "#,##0;(#,##0);-";
            ((Excel.Range)wsLog.Cells[nextRow, 16]).Value2 = sheetKCList;
            ((Excel.Range)wsLog.Cells[nextRow, 17]).NumberFormat = "@";
            ((Excel.Range)wsLog.Cells[nextRow, 17]).Value2 = diagWO;
            ((Excel.Range)wsLog.Cells[nextRow, 18]).NumberFormat = "@";
            ((Excel.Range)wsLog.Cells[nextRow, 18]).Value2 =
                daftarWO.CocokKetat + " / " + daftarWO.CocokLonggar;
            ((Excel.Range)wsLog.Cells[nextRow, 19]).Value2 = totalWO;
            ((Excel.Range)wsLog.Cells[nextRow, 19]).NumberFormat = "#,##0;(#,##0);-";

            // ---- Ringkasan Top-N & klasifikasi (kolom 20-30) ----
            ((Excel.Range)wsLog.Cells[nextRow, 20]).Value2 = topNInfo.DaftarCIF.Count;
            ((Excel.Range)wsLog.Cells[nextRow, 21]).Value2 = topNInfo.JumlahCIF;
            ((Excel.Range)wsLog.Cells[nextRow, 22]).Value2 = topNInfo.OSBruto;
            ((Excel.Range)wsLog.Cells[nextRow, 22]).NumberFormat = "#,##0;(#,##0);-";
            ((Excel.Range)wsLog.Cells[nextRow, 23]).Value2 = stat.RekBertahan;
            ((Excel.Range)wsLog.Cells[nextRow, 24]).Value2 = stat.RekWO;
            ((Excel.Range)wsLog.Cells[nextRow, 25]).Value2 = stat.RekLainnya;
            ((Excel.Range)wsLog.Cells[nextRow, 26]).Value2 = stat.NominalLainnya;
            ((Excel.Range)wsLog.Cells[nextRow, 26]).NumberFormat = "#,##0;(#,##0);-";
            ((Excel.Range)wsLog.Cells[nextRow, 27]).NumberFormat = "@";
            ((Excel.Range)wsLog.Cells[nextRow, 27]).Value2 = stat.RincianWO();
            ((Excel.Range)wsLog.Cells[nextRow, 28]).Value2 = stat.RekWOTakTerpakai;
            ((Excel.Range)wsLog.Cells[nextRow, 29]).Value2 = stat.NominalWOTakTerpakai;
            ((Excel.Range)wsLog.Cells[nextRow, 29]).NumberFormat = "#,##0;(#,##0);-";
            ((Excel.Range)wsLog.Cells[nextRow, 30]).NumberFormat = "@";
            ((Excel.Range)wsLog.Cells[nextRow, 30]).Value2 = stat.RincianWOTakTerpakai();

            // Header ditulis ulang setiap kali agar kolom baru selalu berlabel
            TulisHeaderAuditLog(wsLog, summaryStart - 1);
        }

        private static void TulisHeaderAuditLog(Excel.Worksheet wsLog, int rowHeader)
        {
            ((Excel.Range)wsLog.Cells[rowHeader, 13]).Value2 = "Top-N (diminta)";
            ((Excel.Range)wsLog.Cells[rowHeader, 14]).Value2 = "Rek Skip (Individu)";
            ((Excel.Range)wsLog.Cells[rowHeader, 15]).Value2 = "OS Individu (Top-N)";
            ((Excel.Range)wsLog.Cells[rowHeader, 16]).Value2 = "Ruang Lingkup KC";
            ((Excel.Range)wsLog.Cells[rowHeader, 17]).Value2 = "Diagnostik KC2900";
            ((Excel.Range)wsLog.Cells[rowHeader, 18]).Value2 = "Cocok WO (ketat/longgar)";
            ((Excel.Range)wsLog.Cells[rowHeader, 19]).Value2 = "Nominal WO";
            ((Excel.Range)wsLog.Cells[rowHeader, 20]).Value2 = "CIF Top-N (terpakai)";
            ((Excel.Range)wsLog.Cells[rowHeader, 21]).Value2 = "Jml CIF Awal";
            ((Excel.Range)wsLog.Cells[rowHeader, 22]).Value2 = "OS Awal Bruto";
            ((Excel.Range)wsLog.Cells[rowHeader, 23]).Value2 = "Rek Bertahan";
            ((Excel.Range)wsLog.Cells[rowHeader, 24]).Value2 = "Rek WO";
            ((Excel.Range)wsLog.Cells[rowHeader, 25]).Value2 = "Rek Lainnya";
            ((Excel.Range)wsLog.Cells[rowHeader, 26]).Value2 = "Nominal Lainnya";
            ((Excel.Range)wsLog.Cells[rowHeader, 27]).Value2 = "Rincian WO per Kol";
            ((Excel.Range)wsLog.Cells[rowHeader, 28]).Value2 = "Di KC2900 tapi bukan WO (rek)";
            ((Excel.Range)wsLog.Cells[rowHeader, 29]).Value2 = "Di KC2900 tapi bukan WO (nominal)";
            ((Excel.Range)wsLog.Cells[rowHeader, 30]).Value2 = "Rincian bukan-WO per Kol";
        }

        // ================================================================
        // IndeksRekening: dictionary nomor rekening dua lapis
        //
        //   Lapis 1 (ketat)   : nomor rekening hasil normalisasi apa adanya
        //   Lapis 2 (longgar) : nomor tanpa leading zero — dipakai HANYA
        //                       jika lapis 1 gagal DAN kunci longgar tidak
        //                       ambigu (tidak dimiliki dua rekening berbeda)
        // ================================================================
        private class IndeksRekening
        {
            private readonly Dictionary<string, double[]> _ketat =
                new Dictionary<string, double[]>(StringComparer.OrdinalIgnoreCase);
            private readonly Dictionary<string, double[]> _longgar =
                new Dictionary<string, double[]>(StringComparer.OrdinalIgnoreCase);
            private readonly HashSet<string> _ambigu =
                new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            /// <summary>Jumlah pencocokan yang berhasil lewat kunci ketat.</summary>
            public int CocokKetat { get; private set; }

            /// <summary>Jumlah pencocokan yang berhasil lewat fallback longgar.</summary>
            public int CocokLonggar { get; private set; }

            /// <summary>Jumlah rekening unik dalam indeks.</summary>
            public int Count { get { return _ketat.Count; } }

            /// <param name="akumulasiOS">
            /// true  = OS dijumlahkan bila rekening muncul lebih dari sekali (KC2900/WO)
            /// false = data pertama menang (staging periode akhir)
            /// </param>
            public void Tambah(string rekNorm, double[] data, bool akumulasiOS)
            {
                if (string.IsNullOrEmpty(rekNorm)) return;

                double[] rec;
                if (_ketat.TryGetValue(rekNorm, out rec))
                {
                    if (akumulasiOS) rec[0] += data[0];
                    return;   // kunci longgar untuk rekening ini sudah terdaftar
                }

                _ketat[rekNorm] = data;

                string kl = RekHelper.Longgar(rekNorm);
                if (kl.Length == 0) return;
                if (_longgar.ContainsKey(kl)) _ambigu.Add(kl);   // "00123" vs "123"
                else                          _longgar[kl] = data;
            }

            public bool Cari(string rekNorm, out double[] data)
            {
                string cara;
                return Cari(rekNorm, out data, out cara);
            }

            public bool Cari(string rekNorm, out double[] data, out string cara)
            {
                data = null;
                cara = "-";
                if (string.IsNullOrEmpty(rekNorm)) return false;

                if (_ketat.TryGetValue(rekNorm, out data))
                {
                    CocokKetat++;
                    cara = "Ketat";
                    return true;
                }

                string kl = RekHelper.Longgar(rekNorm);
                if (kl.Length > 0 && !_ambigu.Contains(kl) &&
                    _longgar.TryGetValue(kl, out data))
                {
                    CocokLonggar++;
                    cara = "Longgar";
                    return true;
                }

                data = null;
                cara = "-";
                return false;
            }
        }

        // ----------------------------------------------------------------
        // Utility methods
        // ----------------------------------------------------------------
        private static int CariLastRow(Excel.Worksheet ws, string col, int minRow)
        {
            Excel.Range lastCell = (Excel.Range)ws.Cells[ws.Rows.Count, col];
            Excel.Range endCell  = (Excel.Range)lastCell.End[Excel.XlDirection.xlUp];
            int row = (int)endCell.Row;
            return row < minRow ? minRow : row;
        }

        // Tulis blok object[,] ke range sekaligus — jauh lebih cepat
        // daripada menulis sel demi sel lewat COM.
        private static void TulisBlok(
            Excel.Worksheet ws, string kolAwal, string kolAkhir,
            int startRow, object[,] data)
        {
            int nRows = data.GetLength(0);
            if (nRows == 0) return;
            Excel.Range rng = (Excel.Range)ws.Range[
                kolAwal + startRow, kolAkhir + (startRow + nRows - 1)];
            rng.Value2 = data;
        }

        private static int NormKualitas(string val)
        {
            if (string.IsNullOrWhiteSpace(val)) return 0;
            string s = val.Trim();
            int i = 0;
            while (i < s.Length && char.IsDigit(s[i])) i++;
            if (i == 0) return 0;
            int n;
            return (int.TryParse(s.Substring(0, i), out n) && n >= 1 && n <= 5) ? n : 0;
        }

        // ----------------------------------------------------------------
        // ParseTglToOADate: konversi string "yyyyMMdd" ke OADate Excel
        // Return 0 jika parse gagal (akan fallback ke teks asli)
        // ----------------------------------------------------------------
        private static double ParseTglToOADate(string tglStr)
        {
            if (string.IsNullOrEmpty(tglStr) || tglStr.Length != 8) return 0;
            int yr, mo, dy;
            if (!int.TryParse(tglStr.Substring(0, 4), out yr)) return 0;
            if (!int.TryParse(tglStr.Substring(4, 2), out mo)) return 0;
            if (!int.TryParse(tglStr.Substring(6, 2), out dy)) return 0;
            try
            {
                return new DateTime(yr, mo, dy).ToOADate();
            }
            catch { return 0; }
        }

        private static string ToStr(object val)
        {
            if (val == null) return "";
            return val.ToString().Trim();
        }

        private static double ToDouble(object val)
        {
            if (val == null) return 0;
            if (val is double d) return d;
            if (val is int    i) return i;
            if (val is long   l) return l;
            double r;
            return double.TryParse(val.ToString(), out r) ? r : 0;
        }

        private static Excel.Worksheet CariSheet(Excel.Workbook wb, string name)
        {
            foreach (Excel.Worksheet sh in wb.Worksheets)
                if (string.Equals(sh.Name, name, StringComparison.OrdinalIgnoreCase))
                    return sh;
            return null;
        }
    }
}