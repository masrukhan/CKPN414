using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Web.Script.Serialization;
using CKPNLibrary.Helpers;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Data
{
    // =====================================================================
    // Model penyesuaian tersimpan
    // =====================================================================
    internal class PenyesuaianIndividu
    {
        public string NoKontrak;
        public double Jaminan, BiayaJual, JaminanSistem;
        public string Alasan, Pengguna, Waktu, Periode;
    }

    internal class PenyesuaianLgdCs
    {
        public string NoRek, Jenis;                  // ubah | hapus | tambah
        public double? NilaiAgunan, Recovery, PokokAwal;
        public string Nama, ThnDiserahkan, ThnEksekusi;
        public string Alasan, Pengguna, Waktu, Periode;
    }

    /// <summary>Baris LGD CS hasil sistem (sebelum edit user) — disimpan lokal.</summary>
    internal class DasarLgdCs
    {
        public string Rek, Nama, ThnSerah, ThnEks;
        public double Pokok, Agunan;
        public bool Dikecualikan;                    // disembunyikan oleh penyesuaian 'hapus'
    }

    internal class DasarIndividu
    {
        public string Kontrak;
        public double JaminanSistem;
        public int    Baris;
    }

    /// <summary>
    /// Menerapkan penyesuaian manual tersimpan ke hasil modul, dan mencatat
    /// hasil sistem ("data dasar") untuk dibandingkan saat Simpan grup.
    ///
    /// Dipanggil dari DALAM modul CKPNIndividu dan LGDCollateralShortfall
    /// (lihat PATCH_MODUL_3A.md), sehingga tombol VBA dan panel memberi hasil
    /// yang sama.
    ///
    /// ATURAN KESELAMATAN: kelas ini tidak boleh menggagalkan perhitungan.
    /// Bila database tidak bisa dibaca, perhitungan tetap jalan tanpa
    /// penyesuaian dan masalahnya dicatat di log + info review.
    /// </summary>
    internal static class Penyesuaian
    {
        /// <summary>
        /// Terapkan penyesuaian tersimpan? Default true (tombol VBA selalu menerapkan).
        /// Panel dapat mematikannya untuk satu run guna melihat hasil sistem murni.
        /// </summary>
        public static bool Aktif = true;

        // Warna penanda di sheet (format .Color Excel = BGR)
        private const int WarnaDariPenyesuaian = 0xC07000;   // biru  RGB(0,112,192)
        private const int WarnaRealisasi       = 0x008000;   // hijau RGB(0,128,0)
        private const int WarnaManual          = 0xA03070;   // ungu  RGB(112,48,160)

        private static readonly JavaScriptSerializer _json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

        // ---------------- state run berjalan ----------------
        private static Dictionary<string, PenyesuaianIndividu> _ovInd;
        private static Dictionary<string, PenyesuaianLgdCs>    _ovCs;
        private static List<DasarIndividu> _dasarInd;
        private static List<DasarLgdCs>    _dasarCs;
        private static bool _indDiterapkan, _csDiterapkan;
        private static string _masalahDb;

        // ---------------- info review (dibaca panel) ----------------
        private static Dictionary<string, object> _reviewInd;
        private static Dictionary<string, object> _reviewCs;

        // =================================================================
        // CKPN INDIVIDU
        // =================================================================
        public static void MulaiIndividu()
        {
            _dasarInd = new List<DasarIndividu>();
            _ovInd = new Dictionary<string, PenyesuaianIndividu>(StringComparer.OrdinalIgnoreCase);
            _indDiterapkan = Aktif;
            _masalahDb = null;
            if (!Aktif) return;
            try { _ovInd = MuatIndividu(); }
            catch (Exception ex) { CatatMasalah("Penyesuaian Individu tidak dapat dibaca: " + ex.Message); }
        }

        /// <summary>Dipanggil setelah TulisBaris untuk setiap kontrak.</summary>
        public static void TerapkanIndividu(Excel.Worksheet ws, int baris, string noKontrak, double jaminanSistem)
        {
            try
            {
                if (_dasarInd == null) return;   // modul dipanggil tanpa MulaiIndividu
                string k = (noKontrak ?? "").Trim();
                _dasarInd.Add(new DasarIndividu { Kontrak = k, JaminanSistem = jaminanSistem, Baris = baris });

                PenyesuaianIndividu ov;
                if (!_indDiterapkan || !_ovInd.TryGetValue(k, out ov)) return;

                ((Excel.Range)ws.Cells[baris, "I"]).Value2 = ov.Jaminan;
                if (Math.Abs(ov.BiayaJual) > 0.0001)
                    ((Excel.Range)ws.Cells[baris, "J"]).Value2 = ov.BiayaJual;
            }
            catch (Exception ex) { CatatMasalah("Terapkan penyesuaian Individu " + noKontrak + ": " + ex.Message); }
        }

        /// <summary>Dipanggil setelah tabel selesai di-style (warna penanda tidak tertimpa).</summary>
        public static void SelesaiIndividu(Excel.Worksheet ws)
        {
            try
            {
                if (_dasarInd == null) return;
                var dikenal = KontrakPernahDisimpan();
                var berubah = new List<object>();
                var baru    = new List<object>();
                int diterapkan = 0;

                foreach (var d in _dasarInd)
                {
                    PenyesuaianIndividu ov;
                    if (_indDiterapkan && _ovInd.TryGetValue(d.Kontrak, out ov))
                    {
                        diterapkan++;
                        Excel.Range ij = (Excel.Range)ws.Range["I" + d.Baris, "J" + d.Baris];
                        ij.Font.Color = WarnaDariPenyesuaian;
                        ij.Font.Bold  = true;

                        // Nilai agunan dari sistem berubah sejak penyesuaian disimpan
                        if (Math.Abs(ov.JaminanSistem - d.JaminanSistem) > 0.5)
                            berubah.Add(new Dictionary<string, object>
                            {
                                { "kunci", d.Kontrak }, { "baris", d.Baris },
                                { "sistemLama", ov.JaminanSistem }, { "sistemBaru", d.JaminanSistem },
                                { "dipakai", ov.Jaminan }
                            });
                    }
                    else if (dikenal != null && !dikenal.Contains(d.Kontrak))
                    {
                        baru.Add(new Dictionary<string, object> { { "kunci", d.Kontrak }, { "baris", d.Baris } });
                    }
                }

                SimpanDasar("individu.json", new Dictionary<string, object>
                {
                    { "waktu", Database.Sekarang() },
                    { "penyesuaianDiterapkan", _indDiterapkan },
                    { "baris", _dasarInd }
                });

                _reviewInd = new Dictionary<string, object>
                {
                    { "waktu", Database.Sekarang() },
                    { "penyesuaianDiterapkan", _indDiterapkan },
                    { "jumlahKontrak", _dasarInd.Count },
                    { "diterapkan", diterapkan },
                    { "berubah", berubah },
                    { "baru", baru },
                    { "dikenalTersedia", dikenal != null },
                    { "masalah", _masalahDb }
                };
            }
            catch (Exception ex) { CatatMasalah("Selesai Individu: " + ex.Message); }
            finally { _dasarInd = null; }
        }

        // =================================================================
        // LGD COLLATERAL SHORTFALL
        // =================================================================
        public static void MulaiLgdCs()
        {
            _dasarCs = new List<DasarLgdCs>();
            _ovCs = new Dictionary<string, PenyesuaianLgdCs>(StringComparer.OrdinalIgnoreCase);
            _csDiterapkan = Aktif;
            _masalahDb = null;
            if (!Aktif) return;
            try { _ovCs = MuatLgdCs(); }
            catch (Exception ex) { CatatMasalah("Penyesuaian LGD CS tidak dapat dibaca: " + ex.Message); }
        }

        /// <summary>
        /// Catat baris hasil sistem. Return false bila baris ini dikecualikan
        /// oleh penyesuaian 'hapus' (modul lalu melewatinya).
        /// </summary>
        public static bool CatatBarisLgdCs(string noRek, double pokok, double agunan,
                                           int thnSerah, int thnEks, string nama)
        {
            try
            {
                if (_dasarCs == null) return true;
                string r = (noRek ?? "").Trim();
                PenyesuaianLgdCs ov;
                bool kecuali = _csDiterapkan && _ovCs.TryGetValue(r, out ov) && ov.Jenis == "hapus";
                _dasarCs.Add(new DasarLgdCs
                {
                    Rek = r, Nama = nama, Pokok = pokok, Agunan = agunan,
                    ThnSerah = thnSerah > 0 ? thnSerah.ToString() : "",
                    ThnEks   = thnEks   > 0 ? thnEks.ToString()   : "",
                    Dikecualikan = kecuali
                });
                return !kecuali;
            }
            catch (Exception ex) { CatatMasalah("Catat baris LGD CS " + noRek + ": " + ex.Message); return true; }
        }

        /// <summary>Dipanggil setelah TulisBaris: terapkan nilai agunan / realisasi tersimpan.</summary>
        public static void TerapkanLgdCs(Excel.Worksheet ws, int row, string noRek)
        {
            try
            {
                PenyesuaianLgdCs ov;
                if (_dasarCs == null || !_csDiterapkan || !_ovCs.TryGetValue((noRek ?? "").Trim(), out ov)) return;
                if (ov.Jenis != "ubah") return;

                if (ov.NilaiAgunan.HasValue)
                {
                    Excel.Range d = (Excel.Range)ws.Cells[row, "D"];
                    d.Value2 = ov.NilaiAgunan.Value;
                    d.Font.Color = WarnaDariPenyesuaian;
                }
                if (ov.Recovery.HasValue)
                {
                    // Nilai realisasi menggantikan rumus prefill di kolom G
                    Excel.Range g = (Excel.Range)ws.Cells[row, "G"];
                    g.Value2 = ov.Recovery.Value;
                    g.Font.Color = WarnaRealisasi;
                    try { g.Comment.Delete(); } catch { }
                    try
                    {
                        g.AddComment("Realisasi dari penyesuaian tersimpan\n" +
                                     "oleh " + ov.Pengguna + " (" + ov.Waktu + ")" +
                                     (string.IsNullOrEmpty(ov.Alasan) ? "" : "\n" + ov.Alasan));
                    }
                    catch { }
                }
            }
            catch (Exception ex) { CatatMasalah("Terapkan penyesuaian LGD CS " + noRek + ": " + ex.Message); }
        }

        /// <summary>
        /// Tambahkan baris manual tersimpan (jenis 'tambah') yang tidak ada di hasil
        /// sistem. Return baris kosong berikutnya.
        /// </summary>
        public static int TambahBarisManualLgdCs(Excel.Worksheet ws, int outRow)
        {
            try
            {
                if (_dasarCs == null || !_csDiterapkan) return outRow;
                var adaDiSistem = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var d in _dasarCs) adaDiSistem.Add(d.Rek);

                var daftar = new List<PenyesuaianLgdCs>();
                foreach (var ov in _ovCs.Values)
                    if (ov.Jenis == "tambah" && !adaDiSistem.Contains(ov.NoRek)) daftar.Add(ov);
                daftar.Sort((a, b) => string.Compare(a.NoRek, b.NoRek, StringComparison.OrdinalIgnoreCase));

                foreach (var ov in daftar)
                {
                    int r = outRow;
                    Excel.Range b = (Excel.Range)ws.Cells[r, "B"];
                    b.NumberFormat = "@";
                    b.Value2 = ov.NoRek;
                    ((Excel.Range)ws.Cells[r, "C"]).Value2 = ov.PokokAwal ?? 0;
                    ((Excel.Range)ws.Cells[r, "D"]).Value2 = ov.NilaiAgunan ?? 0;
                    ((Excel.Range)ws.Cells[r, "E"]).NumberFormat = "@";
                    ((Excel.Range)ws.Cells[r, "E"]).Value2 = ov.ThnDiserahkan ?? "";
                    ((Excel.Range)ws.Cells[r, "F"]).NumberFormat = "@";
                    ((Excel.Range)ws.Cells[r, "F"]).Value2 = ov.ThnEksekusi ?? "";

                    Excel.Range g = (Excel.Range)ws.Cells[r, "G"];
                    if (ov.Recovery.HasValue) g.Value2 = ov.Recovery.Value;
                    else g.Formula = "=MIN(D" + r + "*(1-Master!$C$70),C" + r + ")";   // sama dengan prefill modul

                    ((Excel.Range)ws.Cells[r, "H"]).Formula = "=IF(G" + r + ">C" + r + ",0,C" + r + "-G" + r + ")";

                    Excel.Range i = (Excel.Range)ws.Cells[r, "I"];
                    i.Value2 = ov.Nama ?? "";
                    i.Font.Italic = true;

                    Excel.Range baris = (Excel.Range)ws.Range["B" + r, "G" + r];
                    baris.Font.Color = WarnaManual;
                    try
                    {
                        b.AddComment("Baris manual dari penyesuaian tersimpan\noleh " + ov.Pengguna + " (" + ov.Waktu + ")");
                    }
                    catch { }
                    outRow++;
                }
                _reviewCsManual = daftar.Count;
            }
            catch (Exception ex) { CatatMasalah("Tambah baris manual LGD CS: " + ex.Message); }
            return outRow;
        }
        private static int _reviewCsManual;

        public static void SelesaiLgdCs()
        {
            try
            {
                if (_dasarCs == null) return;
                var dikenal = RekPernahDisimpan();
                int kecuali = 0, diubah = 0;
                var baru = new List<object>();

                foreach (var d in _dasarCs)
                {
                    if (d.Dikecualikan) { kecuali++; continue; }
                    PenyesuaianLgdCs ov;
                    if (_csDiterapkan && _ovCs.TryGetValue(d.Rek, out ov))
                    {
                        if (ov.Jenis == "ubah") diubah++;
                    }
                    else if (dikenal != null && !dikenal.Contains(d.Rek))
                    {
                        baru.Add(new Dictionary<string, object> { { "kunci", d.Rek }, { "nama", d.Nama ?? "" } });
                    }
                }

                SimpanDasar("lgdcs.json", new Dictionary<string, object>
                {
                    { "waktu", Database.Sekarang() },
                    { "penyesuaianDiterapkan", _csDiterapkan },
                    { "baris", _dasarCs }
                });

                _reviewCs = new Dictionary<string, object>
                {
                    { "waktu", Database.Sekarang() },
                    { "penyesuaianDiterapkan", _csDiterapkan },
                    { "jumlahSistem", _dasarCs.Count },
                    { "dikecualikan", kecuali },
                    { "diubah", diubah },
                    { "ditambah", _csDiterapkan ? _reviewCsManual : 0 },
                    { "baru", baru },
                    { "dikenalTersedia", dikenal != null },
                    { "masalah", _masalahDb }
                };
            }
            catch (Exception ex) { CatatMasalah("Selesai LGD CS: " + ex.Message); }
            finally { _dasarCs = null; _reviewCsManual = 0; }
        }

        /// <summary>
        /// Tahap 5e: LGD CS tidak dihitung pada run ini (mode setahun sekali, LGD memakai acuan
        /// Desember). Info review LGD CS dari run sebelumnya — mungkin milik grup lain — diganti
        /// supaya tidak muncul sebagai temuan.
        /// </summary>
        public static void LewatiLgdCs(string keterangan)
        {
            _reviewCs = new Dictionary<string, object>
            {
                { "waktu", Database.Sekarang() },
                { "lewati", true }, { "keterangan", keterangan },
                { "penyesuaianDiterapkan", Aktif },
                { "jumlahSistem", 0 }, { "dikecualikan", 0 }, { "diubah", 0 }, { "ditambah", 0 },
                { "baru", new List<object>() },
                { "dikenalTersedia", true },
                { "masalah", null }
            };
        }

        // =================================================================
        // Info review untuk panel
        // =================================================================
        public static Dictionary<string, object> InfoReview()
        {
            return new Dictionary<string, object>
            {
                { "individu", _reviewInd },
                { "lgdcs", _reviewCs }
            };
        }

        // =================================================================
        // Data dasar lokal (dibaca StagingGrup saat Simpan grup)
        // =================================================================
        public static Dictionary<string, object> BacaDasar(string nama)
        {
            try
            {
                string f = Path.Combine(AppPaths.FolderDasar, nama);
                if (!File.Exists(f)) return null;
                return _json.Deserialize<Dictionary<string, object>>(File.ReadAllText(f));
            }
            catch { return null; }
        }

        private static void SimpanDasar(string nama, object isi)
        {
            try
            {
                Directory.CreateDirectory(AppPaths.FolderDasar);
                File.WriteAllText(Path.Combine(AppPaths.FolderDasar, nama), _json.Serialize(isi));
            }
            catch (Exception ex) { CatatMasalah("Simpan data dasar " + nama + ": " + ex.Message); }
        }

        // =================================================================
        // Baca penyesuaian dari database (read-only)
        // =================================================================
        public static Dictionary<string, PenyesuaianIndividu> MuatIndividu()
        {
            var hasil = new Dictionary<string, PenyesuaianIndividu>(StringComparer.OrdinalIgnoreCase);
            if (!Database.Ada) return hasil;
            using (var con = Database.Buka(false))
            using (var cmd = Database.Cmd(con,
                "SELECT no_kontrak, jaminan, biaya_jual, jaminan_sistem, alasan, pengguna, waktu, periode FROM penyesuaian_individu"))
            using (var rd = cmd.ExecuteReader())
            {
                while (rd.Read())
                {
                    var p = new PenyesuaianIndividu
                    {
                        NoKontrak     = rd.GetString(0),
                        Jaminan       = Database.Dbl(rd[1]) ?? 0,
                        BiayaJual     = Database.Dbl(rd[2]) ?? 0,
                        JaminanSistem = Database.Dbl(rd[3]) ?? 0,
                        Alasan = Str(rd[4]), Pengguna = Str(rd[5]), Waktu = Str(rd[6]), Periode = Str(rd[7])
                    };
                    hasil[p.NoKontrak] = p;
                }
            }
            return hasil;
        }

        public static Dictionary<string, PenyesuaianLgdCs> MuatLgdCs()
        {
            var hasil = new Dictionary<string, PenyesuaianLgdCs>(StringComparer.OrdinalIgnoreCase);
            if (!Database.Ada) return hasil;
            using (var con = Database.Buka(false))
            using (var cmd = Database.Cmd(con,
                "SELECT no_rek, jenis, nilai_agunan, recovery, nama, pokok_awal, thn_diserahkan, thn_eksekusi, " +
                "alasan, pengguna, waktu, periode FROM penyesuaian_lgdcs"))
            using (var rd = cmd.ExecuteReader())
            {
                while (rd.Read())
                {
                    var p = new PenyesuaianLgdCs
                    {
                        NoRek = rd.GetString(0), Jenis = rd.GetString(1),
                        NilaiAgunan = Database.Dbl(rd[2]), Recovery = Database.Dbl(rd[3]),
                        Nama = Str(rd[4]), PokokAwal = Database.Dbl(rd[5]),
                        ThnDiserahkan = Str(rd[6]), ThnEksekusi = Str(rd[7]),
                        Alasan = Str(rd[8]), Pengguna = Str(rd[9]), Waktu = Str(rd[10]), Periode = Str(rd[11])
                    };
                    hasil[p.NoRek] = p;
                }
            }
            return hasil;
        }

        // Kontrak / rekening yang pernah masuk staging (untuk menandai "baru").
        // Null bila database belum ada → penanda "baru" tidak ditampilkan.
        private static HashSet<string> KontrakPernahDisimpan()
        {
            return KumpulanDistinct("SELECT DISTINCT no_kontrak FROM hasil_individu");
        }

        private static HashSet<string> RekPernahDisimpan()
        {
            return KumpulanDistinct("SELECT DISTINCT no_rek FROM hasil_lgdcs");
        }

        private static HashSet<string> KumpulanDistinct(string sql)
        {
            try
            {
                if (!Database.Ada) return null;
                var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                using (var con = Database.Buka(false))
                using (var cmd = Database.Cmd(con, sql))
                using (var rd = cmd.ExecuteReader())
                    while (rd.Read()) if (!(rd[0] is DBNull)) set.Add(Convert.ToString(rd[0]).Trim());
                return set;
            }
            catch (Exception ex) { CatatMasalah("Baca riwayat staging: " + ex.Message); return null; }
        }

        private static void CatatMasalah(string pesan)
        {
            _masalahDb = pesan;
            CatatanLog.Tulis("[Penyesuaian] " + pesan);
        }

        private static string Str(object v) { return v == null || v is DBNull ? "" : Convert.ToString(v); }
    }
}
