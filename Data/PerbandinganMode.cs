using System;
using System.Collections.Generic;
using CKPNLibrary.Helpers;
using CKPNLibrary.Panel;

namespace CKPNLibrary.Data
{
    /// <summary>
    /// Tahap 8 — pembanding total CKPN satu periode antar mode perhitungan (per akad vs per kantor).
    ///
    /// Setiap mode dibaca dari databasenya sendiri dengan aturan konsolidasi yang sama (metode tahunan
    /// mode itu, CKPN ABA dari kiriman terbaru). Hanya untuk informasi: angka resmi dan jurnal tetap dari
    /// mode resmi tahun tersebut.
    /// </summary>
    internal static class PerbandinganMode
    {
        public static Dictionary<string, object> Data(string tanggal)
        {
            if (string.IsNullOrEmpty(tanggal) || tanggal.Length < 4) throw new InvalidOperationException("Periode belum dipilih.");
            int tahun = int.Parse(tanggal.Substring(0, 4));
            var resmi = ModeHitung.ModeResmi(tahun);
            string aktif = ModeHitung.Aktif.Id;

            var baris = new List<object>();
            int adaDb = 0;
            foreach (var m in ModeHitung.Semua)
            {
                var b = new Dictionary<string, object>
                {
                    { "id", m.Id }, { "nama", m.Nama }, { "namaPendek", m.NamaPendek },
                    { "aktif", m.Id == aktif }, { "resmi", m.Id == resmi.Mode }, { "adaDatabase", Database.AdaMode(m) }
                };
                baris.Add(b);
                if (!Database.AdaMode(m)) { b["catatan"] = "Mode ini belum pernah dipakai."; continue; }
                adaDb++;
                try
                {
                    var k = Periode.Konsolidasi(tanggal, m);
                    var grup = (List<object>)k["grup"];
                    var belum = (List<string>)k["belum"];
                    double ind = 0, kol = 0;
                    foreach (Dictionary<string, object> g in grup)
                    {
                        ind += Convert.ToDouble(g["ind"]);
                        kol += Convert.ToDouble(g["kol"]);
                    }
                    var total = (Dictionary<string, object>)k["total"];
                    var aba = (Dictionary<string, object>)k["aba"];
                    b["status"] = k["status"];
                    b["metode"] = k["metode"];
                    b["jumlahGrup"] = grup.Count + belum.Count;
                    b["grupTersimpan"] = grup.Count;
                    b["lengkap"] = k["lengkap"];
                    b["individu"] = ind;
                    b["kolektif"] = kol;
                    b["aba"] = aba["ckpn"];
                    b["ckpn"] = total["ckpn"];
                    b["ppka"] = total["ppka"];
                    var rincian = new List<object>();
                    foreach (Dictionary<string, object> g in grup)
                        rincian.Add(new Dictionary<string, object>
                        {
                            { "nama", g["nama"] }, { "labelKode", g["labelKode"] }, { "ckpn", g["ckpn"] }, { "ppka", g["ppka"] },
                            { "ind", g["ind"] }, { "kol", g["kol"] }
                        });
                    b["grup"] = rincian;
                    if (belum.Count > 0) b["catatan"] = "Belum tersimpan: " + string.Join(", ", belum.ToArray()) + ".";
                }
                catch (Exception ex) { b["catatan"] = ex.Message; }
            }
            return new Dictionary<string, object>
            {
                { "tanggal", tanggal }, { "tahun", tahun }, { "resmi", resmi.Mode }, { "aktif", aktif },
                { "mode", baris }, { "jumlahDatabase", adaDb }
            };
        }
    }
}
