using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Web.Script.Serialization;
using CKPNLibrary.Helpers;
using CKPNLibrary.Panel;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Data
{
    /// <summary>
    /// Tahap 9c — kertas kerja perhitungan CKPN per grup (segmen) dari DATABASE, bukan dari sheet Excel.
    ///
    /// Setelah "Hitung semua grup", sheet Excel hanya berisi grup terakhir, dan snapshot .xlsx tidak memuat
    /// penyesuaian yang diterapkan belakangan (koreksi panel / Terapkan penyesuaian ke versi tersimpan).
    /// Yang benar adalah versi tersimpan di database. Kelas ini mengumpulkan seluruh isi satu versi grup:
    ///   • run_grup         identitas versi, signifikansi, catatan
    ///   • ringkasan        CKPN Net Flow/Migration, LGD CS/ER/weighted, PPKA
    ///   • hasil_individu   daftar CKPN Individu per kontrak (setelah penyesuaian)
    ///   • hasil_lgdcs      rincian LGD Collateral Shortfall per rekening
    ///   • analisis_pd 'pd'        saldo bucket Net Flow 12 bulan, matriks migrasi per triwulan, PD resmi
    ///   • analisis_pd 'kolektif'  EAD/PD/LGD/CKPN per bucket & per kualitas
    ///   • analisis_pd 'lgder'     rincian LGD ER per kohort tahun hapus buku (Tahap 9d)
    /// Panel menyusunnya menjadi dokumen A4 (Preview → PDF/cetak) dan bisa mengekspornya ke Excel
    /// (<see cref="EksporExcel"/>).
    /// </summary>
    internal static class PerhitunganGrup
    {
        private static readonly JavaScriptSerializer _json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

        public static Dictionary<string, object> Data(long runId)
        {
            if (!Database.Ada) throw new InvalidOperationException("Database belum ada.");
            var h = new Dictionary<string, object>();
            using (var con = Database.Buka(false))
            {
                string tanggal = null, kode = null;
                using (var cmd = Database.Cmd(con,
                    "SELECT g.id, p.tanggal, p.status, g.kode_kc, g.versi, g.aktif, g.dihapus, g.pengguna, g.komputer, g.waktu, " +
                    "g.versi_addin, g.parameter_json, g.catatan, g.top_n, g.signifikan, g.ambang " +
                    "FROM run_grup g JOIN periode p ON p.id=g.periode_id WHERE g.id=@p0", runId))
                using (var rd = cmd.ExecuteReader())
                {
                    if (!rd.Read()) throw new InvalidOperationException("Versi grup tidak ditemukan.");
                    tanggal = rd.GetString(1); kode = rd.GetString(3);
                    string sig = rd[14] is DBNull ? Signifikansi.Ranking : rd.GetString(14);
                    int? topN = rd[13] is DBNull ? (int?)null : Convert.ToInt32(rd[13]);
                    double? ambang = rd[15] is DBNull ? (double?)null : Convert.ToDouble(rd[15]);
                    h["run"] = new Dictionary<string, object>
                    {
                        { "id", rd.GetInt64(0) }, { "tanggal", tanggal }, { "statusPeriode", Str(rd[2]) }, { "kodeKC", kode },
                        { "versi", rd.GetInt32(4) }, { "aktif", rd.GetInt32(5) == 1 }, { "dihapus", rd.GetInt32(6) == 1 },
                        { "pengguna", Str(rd[7]) }, { "komputer", Str(rd[8]) }, { "waktu", Str(rd[9]) }, { "versiAddin", Str(rd[10]) },
                        { "parameter", Urai(Str(rd[11])) }, { "catatan", Str(rd[12]) },
                        { "topN", topN }, { "signifikan", sig }, { "ambang", ambang },
                        { "labelSignifikan", Signifikansi.Label(sig, topN, ambang) }
                    };
                }

                // Nama grup & unit menurut susunan tahun periode tersebut (mode aktif)
                int sumber;
                var susunan = Periode.Susunan(con, int.Parse(tanggal.Substring(0, 4)), out sumber);
                var g = susunan == null ? null : susunan.Find(x => string.Equals(x.KodeKC, kode, StringComparison.OrdinalIgnoreCase));
                var mode = ModeHitung.Aktif;
                h["grup"] = new Dictionary<string, object>
                {
                    { "nama", g == null ? kode : g.Nama }, { "kodeKC", kode }, { "labelKode", ModeHitung.LabelKode(mode, kode) },
                    { "kc", ModeHitung.KCGrup(mode, kode) }, { "diSusunan", g != null }
                };
                var pt = Periode.Pengaturan(con, int.Parse(tanggal.Substring(0, 4)));
                h["metode"] = pt == null ? null : pt.Metode;
                h["mode"] = new Dictionary<string, object> { { "id", mode.Id }, { "nama", mode.Nama } };

                // Ringkasan nilai
                var rg = new Dictionary<string, object>();
                using (var cmd = Database.Cmd(con, "SELECT kunci, nilai FROM ringkasan WHERE run_id=@p0", runId))
                using (var rd = cmd.ExecuteReader())
                    while (rd.Read()) rg[rd.GetString(0)] = Database.Dbl(rd[1]);
                h["ringkasan"] = rg;
                double ppka = 0;
                foreach (var kc in ModeHitung.KCGrup(mode, kode))
                {
                    object v;
                    if (rg.TryGetValue("ppka_" + kc.Trim(), out v) && v != null) ppka += Convert.ToDouble(v);
                }
                h["ppkaGrup"] = ppka;

                // CKPN Individu per kontrak
                var ind = new List<object>();
                using (var cmd = Database.Cmd(con,
                    "SELECT urut, kc, cif, nama, no_kontrak, os, ada_pn, jaminan, biaya_jual, penurunan_nilai, disesuaikan " +
                    "FROM hasil_individu WHERE run_id=@p0 ORDER BY urut, rowid", runId))
                using (var rd = cmd.ExecuteReader())
                    while (rd.Read())
                        ind.Add(new Dictionary<string, object>
                        {
                            { "urut", rd[0] is DBNull ? (object)null : Convert.ToInt32(rd[0]) }, { "kc", Str(rd[1]) }, { "cif", Str(rd[2]) },
                            { "nama", Str(rd[3]) }, { "kontrak", Str(rd[4]) }, { "os", Database.Dbl(rd[5]) ?? 0 }, { "adaPN", Str(rd[6]) },
                            { "jaminan", Database.Dbl(rd[7]) ?? 0 }, { "biaya", Database.Dbl(rd[8]) ?? 0 }, { "ckpn", Database.Dbl(rd[9]) ?? 0 },
                            { "disesuaikan", !(rd[10] is DBNull) && Convert.ToInt32(rd[10]) == 1 }
                        });
                h["individu"] = ind;

                // LGD Collateral Shortfall per rekening
                var cs = new List<object>();
                using (var cmd = Database.Cmd(con,
                    "SELECT no_rek, nama, pokok_awal, nilai_agunan, thn_diserahkan, thn_eksekusi, recovery, shortfall, sumber " +
                    "FROM hasil_lgdcs WHERE run_id=@p0 ORDER BY rowid", runId))
                using (var rd = cmd.ExecuteReader())
                    while (rd.Read())
                        cs.Add(new Dictionary<string, object>
                        {
                            { "rek", Str(rd[0]) }, { "nama", Str(rd[1]) }, { "pokok", Database.Dbl(rd[2]) ?? 0 }, { "agunan", Database.Dbl(rd[3]) ?? 0 },
                            { "thnSerah", Str(rd[4]) }, { "thnEks", Str(rd[5]) }, { "recovery", Database.Dbl(rd[6]) ?? 0 },
                            { "shortfall", Database.Dbl(rd[7]) ?? 0 }, { "sumber", Str(rd[8]) }
                        });
                h["lgdcs"] = cs;

                // Analisis PD & rincian kolektif (JSON tersimpan saat Simpan grup / koreksi)
                h["pd"] = UraiObj(Database.Scalar(con, "SELECT data FROM analisis_pd WHERE run_id=@p0 AND jenis='pd'", runId));
                h["kolektif"] = UraiObj(Database.Scalar(con, "SELECT data FROM analisis_pd WHERE run_id=@p0 AND jenis=@p1", runId, RincianCkpn.Jenis));
                h["lgder"] = UraiObj(Database.Scalar(con, "SELECT data FROM analisis_pd WHERE run_id=@p0 AND jenis=@p1", runId, LgdEr.Jenis));   // Tahap 9d

                // Versi acuan PD & LGD (mode setahun): ambil label periode acuannya
                object acuanId;
                if (rg.TryGetValue("acuan_run_id", out acuanId) && acuanId != null)
                {
                    object t = Database.Scalar(con, "SELECT p.tanggal FROM run_grup g JOIN periode p ON p.id=g.periode_id WHERE g.id=@p0",
                        Convert.ToInt64(Convert.ToDouble(acuanId)));
                    h["acuanTanggal"] = t == null || t is DBNull ? null : Convert.ToString(t);
                }
            }
            return h;
        }

        // ================================================================
        // Tahap 9e: lengkapi rincian versi tersimpan dari sheet Excel
        // ================================================================
        /// <summary>
        /// Versi yang disimpan sebelum Tahap 9d punya deret Net Flow 12 bulan dan belum punya rincian LGD ER.
        /// Bila kertas kerja Excel saat ini berisi grup &amp; periode yang sama dan angka Summary-nya SAMA dengan
        /// versi aktif (tidak ada angka CKPN yang berubah), rincian PD (saldo 13 bulan, roll rate, flow to loss)
        /// dan LGD ER versi itu diperbarui dari sheet — tanpa membuat versi baru.
        /// Mode acuan PD &amp; LGD: rincian disalin dari versi acuan (Desember), bukan dari sheet.
        /// Angka CKPN, Individu, LGD CS, dan rincian kolektif per bucket tidak disentuh.
        /// </summary>
        public static Dictionary<string, object> LengkapiDariSheet(Excel.Application app, long runId)
        {
            string info;
            if (!Database.BolehMenulis(out info)) throw new InvalidOperationException(info);
            Excel.Workbook wb = PanelBridge.CariWorkbookAplikasi(app);
            if (wb == null) throw new InvalidOperationException("Workbook aplikasi CKPN tidak sedang terbuka.");

            string tanggal, kode; int versi; bool aktif; long? acuanId = null; bool koreksi = false;
            using (var con = Database.Buka(false))
            {
                using (var cmd = Database.Cmd(con,
                    "SELECT p.tanggal, g.kode_kc, g.versi, g.aktif, g.dihapus FROM run_grup g JOIN periode p ON p.id=g.periode_id WHERE g.id=@p0", runId))
                using (var rd = cmd.ExecuteReader())
                {
                    if (!rd.Read()) throw new InvalidOperationException("Versi grup tidak ditemukan.");
                    tanggal = rd.GetString(0); kode = rd.GetString(1); versi = rd.GetInt32(2);
                    aktif = rd.GetInt32(3) == 1 && rd.GetInt32(4) == 0;
                }
                object a = Database.Scalar(con, "SELECT nilai FROM ringkasan WHERE run_id=@p0 AND kunci='acuan_run_id'", runId);
                if (a != null && !(a is DBNull)) acuanId = Convert.ToInt64(Convert.ToDouble(a));
                koreksi = Database.Scalar(con, "SELECT nilai FROM ringkasan WHERE run_id=@p0 AND kunci='koreksi_dari_versi'", runId) != null;
            }
            if (!aktif) throw new InvalidOperationException("Hanya versi aktif yang dapat dilengkapi dari sheet.");

            // Kertas kerja harus berisi grup & periode yang sama, dengan angka Summary yang sama
            var cocok = LabelLingkup.Cocokkan(wb);
            var iw = (Dictionary<string, object>)cocok["info"];
            if (Convert.ToString(iw["tanggal"]) != tanggal || !string.Equals(Convert.ToString(iw["kodeKC"]), kode, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Kertas kerja Excel saat ini berisi grup/periode lain. Hitung grup ini dulu dari Hitung › Per grup, " +
                    "atau simpan sebagai versi baru setelah dihitung.");
            int beda = Convert.ToInt32(cocok["jumlahBeda"]);
            if (beda > 0)
                throw new InvalidOperationException((koreksi
                    ? "Versi v" + versi + " adalah hasil koreksi/penyesuaian panel sehingga angkanya berbeda dengan sheet. "
                    : beda + " angka Summary di sheet berbeda dengan v" + versi + ". ") +
                    "Rincian tidak dilengkapi dari sheet agar tidak mencampur hasil perhitungan berbeda. Hitung ulang grup lalu Simpan grup (versi baru).");

            string pdJson = null, erJson = null;
            if (acuanId == null)
            {
                pdJson = AnalisisPD.BacaWorkbookJson(wb);
                erJson = LgdEr.BacaWorkbookJson(wb);
            }
            int n = 0;
            Database.Cadangkan();
            using (var con = Database.Buka(true))
            using (var tx = con.BeginTransaction())
            {
                if (acuanId != null)
                {
                    Database.Exec(con, "INSERT OR REPLACE INTO analisis_pd(run_id, jenis, data) " +
                        "SELECT @p0, jenis, data FROM analisis_pd WHERE run_id=@p1 AND jenis IN ('pd', @p2)", runId, acuanId.Value, LgdEr.Jenis);
                    n = Convert.ToInt32(Database.Scalar(con, "SELECT COUNT(*) FROM analisis_pd WHERE run_id=@p0 AND jenis IN ('pd', @p1)", acuanId.Value, LgdEr.Jenis));
                }
                else
                {
                    if (!string.IsNullOrEmpty(pdJson)) { AnalisisPD.SimpanRun(con, runId, pdJson); n++; }
                    if (!string.IsNullOrEmpty(erJson)) { LgdEr.SimpanRun(con, runId, erJson); n++; }
                }
                Database.CatatAktivitas(con, tanggal, "lengkapi-rincian",
                    "v" + versi + " " + kode + ": rincian PD Net Flow & LGD ER diperbarui dari " + (acuanId != null ? "versi acuan" : "sheet Excel") +
                    " (angka CKPN tidak berubah)");
                tx.Commit();
            }
            LogProses.CatatPanel(tanggal, "Dokumen", "Lengkapi rincian v" + versi, LogProses.OK,
                LogProses.R().Tambah("Grup", kode).Tambah("Sumber", acuanId != null ? "versi acuan" : "sheet Excel").Tambah("Bagian", n));
            return new Dictionary<string, object> { { "versi", versi }, { "bagian", n }, { "acuan", acuanId != null } };
        }

        // ================================================================
        // Ekspor tabel dokumen ke workbook Excel baru (.xlsx)
        // ================================================================
        /// <summary>
        /// lembar: [{ nama, judul, blok: [{ judul, kolom: [..], baris: [[..]], format: ["teks"|"rp"|"persen"|"angka"] }] }]
        /// Nilai dikirim mentah (angka tetap angka) supaya bisa dihitung ulang di Excel.
        /// File disimpan di library\dokumen\yyyy-MM\ lalu dibiarkan terbuka.
        /// </summary>
        public static Dictionary<string, object> EksporExcel(Excel.Application app, string nama, string periode, IList lembar)
        {
            if (lembar == null || lembar.Count == 0) throw new InvalidOperationException("Tidak ada data untuk diekspor.");
            string sub = periode != null && periode.Length >= 7 ? periode.Substring(0, 7) : "lain";
            string folder = Path.Combine(PanelBridge.FolderDokumen, sub);
            Directory.CreateDirectory(folder);
            string dasar = NamaFileAman(nama);
            string path = Path.Combine(folder, dasar + ".xlsx");
            if (File.Exists(path))
            {
                try { File.Delete(path); }
                catch { path = Path.Combine(folder, dasar + "_" + DateTime.Now.ToString("HHmmss") + ".xlsx"); }
            }

            bool layar = app.ScreenUpdating;
            Excel.Workbook wb = null;
            try
            {
                app.ScreenUpdating = false;
                wb = app.Workbooks.Add();
                // Workbook baru berisi 1+ sheet bawaan: pakai yang pertama, tambah sisanya di belakang
                int i = 0;
                foreach (Dictionary<string, object> l in lembar)
                {
                    Excel.Worksheet ws;
                    if (i < wb.Worksheets.Count) ws = (Excel.Worksheet)wb.Worksheets[i + 1];
                    else ws = (Excel.Worksheet)wb.Worksheets.Add(After: wb.Worksheets[wb.Worksheets.Count]);
                    i++;
                    try { ws.Name = NamaSheetAman(Convert.ToString(Ambil(l, "nama") ?? ("Lembar " + i)), i); } catch { }
                    TulisLembar(ws, l);
                }
                // Sheet bawaan yang tidak terpakai dihapus
                app.DisplayAlerts = false;
                while (wb.Worksheets.Count > i && wb.Worksheets.Count > 1)
                    ((Excel.Worksheet)wb.Worksheets[wb.Worksheets.Count]).Delete();
                app.DisplayAlerts = true;
                ((Excel.Worksheet)wb.Worksheets[1]).Activate();
                wb.SaveAs(path, Excel.XlFileFormat.xlOpenXMLWorkbook);
            }
            catch
            {
                if (wb != null) try { wb.Close(false); } catch { }
                throw;
            }
            finally
            {
                app.DisplayAlerts = true;
                app.ScreenUpdating = layar;
            }
            try { wb.Activate(); } catch { }
            LogProses.CatatPanel(periode ?? "", "Dokumen", "Ekspor Excel " + Path.GetFileName(path), LogProses.OK, LogProses.R().Tambah("File", path));
            return new Dictionary<string, object> { { "file", path }, { "nama", Path.GetFileName(path) } };
        }

        private static void TulisLembar(Excel.Worksheet ws, Dictionary<string, object> l)
        {
            int baris = 1;
            string judul = Convert.ToString(Ambil(l, "judul") ?? "");
            if (judul.Length > 0)
            {
                Excel.Range r = (Excel.Range)ws.Cells[baris, 1];
                r.Value2 = judul;
                r.Font.Bold = true; r.Font.Size = 13;
                baris += 2;
            }
            int lebarMaks = 1;
            var blok = Ambil(l, "blok") as IList;
            if (blok == null) return;
            foreach (Dictionary<string, object> b in blok)
            {
                string jb = Convert.ToString(Ambil(b, "judul") ?? "");
                if (jb.Length > 0)
                {
                    Excel.Range r = (Excel.Range)ws.Cells[baris, 1];
                    r.Value2 = jb; r.Font.Bold = true;
                    baris++;
                }
                var kolom = Ambil(b, "kolom") as IList;
                var isi = Ambil(b, "baris") as IList;
                var format = Ambil(b, "format") as IList;
                int nKol = kolom != null ? kolom.Count : 0;
                if (isi != null) foreach (var o in isi) { var rr = o as IList; if (rr != null && rr.Count > nKol) nKol = rr.Count; }
                if (nKol == 0) { baris++; continue; }
                lebarMaks = Math.Max(lebarMaks, nKol);

                if (kolom != null && kolom.Count > 0)
                {
                    var hdr = new object[1, nKol];
                    for (int c = 0; c < kolom.Count; c++) hdr[0, c] = Convert.ToString(kolom[c]);
                    Excel.Range rh = ws.Range[ws.Cells[baris, 1], ws.Cells[baris, nKol]];
                    rh.Value2 = hdr;
                    rh.Font.Bold = true;
                    rh.Interior.Color = 0xEEF0EE;   // abu muda (BGR)
                    rh.WrapText = true;
                    baris++;
                }
                int n = isi == null ? 0 : isi.Count;
                if (n > 0)
                {
                    var data = new object[n, nKol];
                    for (int r = 0; r < n; r++)
                    {
                        var rr = isi[r] as IList;
                        if (rr == null) continue;
                        for (int c = 0; c < rr.Count && c < nKol; c++) data[r, c] = Nilai(rr[c]);
                    }
                    Excel.Range rd = ws.Range[ws.Cells[baris, 1], ws.Cells[baris + n - 1, nKol]];
                    // Kolom teks (no. rekening/CIF/kontrak) diformat teks dulu agar nol di depan tidak hilang
                    if (format != null)
                        for (int c = 0; c < format.Count && c < nKol; c++)
                        {
                            string f = Convert.ToString(format[c]);
                            Excel.Range kc = ws.Range[ws.Cells[baris, c + 1], ws.Cells[baris + n - 1, c + 1]];
                            if (f == "rp") kc.NumberFormat = "#,##0;(#,##0);-";
                            else if (f == "persen") kc.NumberFormat = "0.00%";
                            else if (f == "angka") kc.NumberFormat = "#,##0.####";
                            else kc.NumberFormat = "@";
                        }
                    rd.Value2 = data;
                    // Baris total (teks "Total" di kolom pertama) ditebalkan
                    for (int r = 0; r < n; r++)
                    {
                        var rr = isi[r] as IList;
                        if (rr != null && rr.Count > 0 && Convert.ToString(rr[0]).StartsWith("Total", StringComparison.OrdinalIgnoreCase))
                            ws.Range[ws.Cells[baris + r, 1], ws.Cells[baris + r, nKol]].Font.Bold = true;
                    }
                    baris += n;
                }
                baris++;   // baris kosong antar blok
            }
            try
            {
                Excel.Range semua = ws.Range[ws.Columns[1], ws.Columns[lebarMaks]];
                semua.ColumnWidth = 16;
                ((Excel.Range)ws.Columns[1]).ColumnWidth = 26;
            }
            catch { }
        }

        private static object Nilai(object v)
        {
            if (v == null) return null;
            if (v is decimal) return Convert.ToDouble(v);
            if (v is int || v is long || v is double || v is bool) return v;
            return Convert.ToString(v, CultureInfo.InvariantCulture);
        }

        private static string NamaSheetAman(string s, int i)
        {
            foreach (char c in new[] { '\\', '/', '?', '*', '[', ']', ':' }) s = s.Replace(c, ' ');
            s = s.Trim();
            if (s.Length == 0) s = "Lembar " + i;
            return s.Length > 31 ? s.Substring(0, 31) : s;
        }

        private static string NamaFileAman(string nama)
        {
            foreach (char c in Path.GetInvalidFileNameChars()) nama = nama.Replace(c, '_');
            nama = nama.Trim();
            return nama.Length == 0 ? "Dokumen CKPN" : (nama.Length > 120 ? nama.Substring(0, 120) : nama);
        }

        // ================================================================
        private static object UraiObj(object v)
        {
            if (v == null || v is DBNull) return null;
            try { return _json.Deserialize<Dictionary<string, object>>(Convert.ToString(v)); }
            catch { return null; }
        }

        private static object Urai(string s)
        {
            if (string.IsNullOrEmpty(s)) return null;
            try { return _json.Deserialize<Dictionary<string, object>>(s); } catch { return null; }
        }

        private static object Ambil(Dictionary<string, object> d, string k)
        {
            object v;
            return d != null && d.TryGetValue(k, out v) ? v : null;
        }

        private static string Str(object v) { return v == null || v is DBNull ? "" : Convert.ToString(v); }
    }
}
