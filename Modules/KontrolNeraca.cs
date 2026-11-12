using System;
using System.Collections.Generic;
using CKPNLibrary.Helpers;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Modules
{
    /// <summary>
    /// Kontrol OS neraca (GB0200) vs OS segmen (KC0600–KC1100) — dipakai Analisis › Kontrol.
    ///
    /// Relasi pos neraca ↔ kolom nominatif mengikuti rumus KC0200/GB0200 di template LBBPRS:
    ///   E8  Piutang Murabahah            = KC0600 Σ AE (Jumlah = pokok + margin)
    ///   E9  Margin Murabahah ditangguhkan = KC0600 Σ AD
    ///   E10 / E11 Istishna               = KC0700 Σ AE / Σ AD
    ///   E12 / E13 Multijasa              = KC0800 Σ AE / Σ AD
    ///   E14 Piutang Qardh                = KC0900 Σ AC (Jumlah)
    ///   E17 Mudarabah                    = KC1000 Σ AI, jenis akad (U) &lt; 29
    ///   E18 Musyarakah                   = KC1000 Σ AI, jenis akad (U) &gt; 29
    ///   E19 Bagi hasil lainnya           = KC1000 Σ AI, jenis akad lain/kosong
    ///   E25 Aset Ijarah                  = KC1100 Σ X (harga perolehan)
    ///   E26 Akumulasi penyusutan         = KC1100 Σ AI
    ///   E15 Piutang Sewa                 = KC1100 Σ AL (tunggakan pokok) — bagian baki debet ijarah
    ///   E27 CKPN aset ijarah             = KC1100 Σ AJ — informasi, tidak mengurangi OS
    ///   Tunggakan ujroh AM               = tidak ada pos neraca, bukan OS (objek CKPN KC1100 kualitas 1–2)
    ///   E21 Salam                        = tidak ada sheet segmen CKPN → selisih tetap
    ///
    /// OS neraca = (E7 − E15) + E16 + E21 + (E25 − E26 + E15) — baki ijarah = sandi 181 − 182 + 160 (sama dengan DataOverviewBuilder).
    /// OS segmen = OS per KC yang dipakai Ringkasan (baris berdebitur, kualitas 1–5):
    ///   KC0600–0800 AC · KC0900 AA · KC1000 AG · KC1100 X − AI + AL (baki debet ijarah, = GB0200 E25 − E26 + E15).
    ///
    /// Selisih total dipecah menjadi empat bagian yang jumlahnya PERSIS sama dengan selisih:
    ///   a. pos neraca tanpa segmen (Salam)
    ///   b. selisih sumber   : nilai GB0200 ≠ jumlah kolom nominatif KC (tarikan berbeda / rekening di luar nominatif)
    ///   c. selisih cakupan  : nominatif KC ≠ OS segmen (baris tanpa nama/kualitas, atau kolom OS beda dengan kolom neraca)
    ///   d. selisih subtotal : subtotal GB0200 (E7, E16) ≠ penjumlahan rinciannya
    /// </summary>
    internal static class KontrolNeraca
    {
        private const int BatasContoh = 25;   // baris penyebab yang disimpan per segmen (snapshot tetap kecil)

        // ================================================================
        // Model per segmen (diisi saat DataOverviewBuilder membaca sheet KC)
        // ================================================================
        internal class Segmen
        {
            public string Kode;
            public bool   Ada;
            public string Pesan = "";
            /// <summary>Σ kolom nominatif seluruh baris (tanpa filter), kunci = nama kolom / sub-pos.</summary>
            public Dictionary<string, double> Kolom = new Dictionary<string, double>();
            public double NetoKC;          // kolom neraca neto, seluruh baris
            public double OSSegmen;        // = total OS segmen Ringkasan
            public double Dikecualikan;    // nilai neto baris yang tidak masuk OS segmen
            public int    NDikecualikan;
            public double BedaKolom;       // selisih kolom neraca vs kolom OS pada baris yang masuk
            public int    NBedaKolom;
            public List<Dictionary<string, object>> Contoh = new List<Dictionary<string, object>>();
        }

        private static readonly string[] Urutan = { "KC0600", "KC0700", "KC0800", "KC0900", "KC1000", "KC1100" };

        private static string NamaSegmen(string kode)
        {
            switch (kode)
            {
                case "KC0600": return "Murabahah";
                case "KC0700": return "Istishna";
                case "KC0800": return "Multijasa";
                case "KC0900": return "Qardh";
                case "KC1000": return "Bagi hasil (Mudarabah/Musyarakah)";
                case "KC1100": return "Ijarah";
                default:       return kode;
            }
        }

        private static string KolomOS(string kode)
        {
            switch (kode)
            {
                case "KC0900": return "AA (Nominal)";
                case "KC1000": return "AG (Nominal)";
                case "KC1100": return "X − AI + AL (harga perolehan − akumulasi penyusutan + tunggakan pokok)";
                default:       return "AC (Saldo harga pokok)";
            }
        }

        private static string KolomNeto(string kode)
        {
            switch (kode)
            {
                case "KC0900": return "AC (Jumlah)";
                case "KC1000": return "AI (Jumlah)";
                case "KC1100": return "X − AI + AL (harga perolehan − penyusutan + tunggakan pokok)";
                default:       return "AE − AD (jumlah − margin ditangguhkan)";
            }
        }

        // ================================================================
        // 1. Baca sheet KC — dipanggil dari DataOverviewBuilder.BacaSegmen
        //    Syarat "masuk OS segmen" SAMA dengan loop Overview: nama tidak kosong & kualitas 1–5.
        // ================================================================
        public static Segmen BacaSegmen(Excel.Worksheet ws, string kode, int dataStart, int lastRow,
                                        string[] namaArr, string[] kualArr, string[] rekArr, double[] osArr)
        {
            var k = new Segmen { Kode = kode, Ada = true };
            int n = namaArr.Length;
            Func<string, double[]> kol = c => ExcelHelper.BacaKolomDouble(ws, c, dataStart, lastRow);

            double[] neto = new double[n];
            switch (kode)
            {
                case "KC0600":
                case "KC0700":
                case "KC0800":
                {
                    double[] ae = kol("AE"), ad = kol("AD");
                    for (int i = 0; i < n; i++) neto[i] = ae[i] - ad[i];
                    k.Kolom["AE"] = Jumlah(ae);
                    k.Kolom["AD"] = Jumlah(ad);
                    break;
                }
                case "KC0900":
                {
                    double[] ac = kol("AC");
                    Array.Copy(ac, neto, n);
                    k.Kolom["AC"] = Jumlah(ac);
                    break;
                }
                case "KC1000":
                {
                    // Pemisahan Mudarabah/Musyarakah mengikuti rumus KC0200: jenis akad (U) < 29 / > 29
                    double[] ai = kol("AI");
                    string[] akad = ExcelHelper.BacaKolomString(ws, "U", dataStart, lastRow);
                    double mud = 0, mus = 0, lain = 0;
                    for (int i = 0; i < n; i++)
                    {
                        neto[i] = ai[i];
                        int u;
                        string s = i < akad.Length ? (akad[i] ?? "").Trim() : "";
                        if (int.TryParse(s, out u) && u < 29) mud += ai[i];
                        else if (int.TryParse(s, out u) && u > 29) mus += ai[i];
                        else lain += ai[i];
                    }
                    k.Kolom["AI_MUD"] = mud;
                    k.Kolom["AI_MUS"] = mus;
                    k.Kolom["AI_LAIN"] = lain;
                    k.Kolom["AI"] = mud + mus + lain;
                    break;
                }
                case "KC1100":
                {
                    double[] x = kol("X"), ai = kol("AI"), aj = kol("AJ"), al = kol("AL"), am = kol("AM");
                    for (int i = 0; i < n; i++) neto[i] = x[i] - ai[i] + al[i];   // = sandi 181 − 182 + 160
                    k.Kolom["X"] = Jumlah(x);
                    k.Kolom["AI"] = Jumlah(ai);
                    k.Kolom["AJ"] = Jumlah(aj);
                    k.Kolom["AL"] = Jumlah(al);   // piutang sewa (tunggakan pokok)
                    k.Kolom["AM"] = Jumlah(am);   // tunggakan ujroh — tidak ada pos neraca
                    break;
                }
            }

            var kandidat = new List<Dictionary<string, object>>();
            for (int i = 0; i < n; i++)
            {
                k.NetoKC += neto[i];
                bool adaNama = !string.IsNullOrEmpty(namaArr[i]);
                int kual = NormKualitas(i < kualArr.Length ? kualArr[i] : null);
                bool masuk = adaNama && kual >= 1 && kual <= 5;
                if (masuk)
                {
                    double os = i < osArr.Length ? osArr[i] : 0;
                    k.OSSegmen += os;
                    double beda = neto[i] - os;
                    if (Math.Abs(beda) >= 1)
                    {
                        k.BedaKolom += beda;
                        k.NBedaKolom++;
                        kandidat.Add(Contoh(dataStart + i, rekArr, namaArr, kualArr, i, beda,
                            "kolom neraca " + KolomNeto(kode).Split(' ')[0] + " ≠ kolom OS " + KolomOS(kode).Split(' ')[0]));
                    }
                }
                else if (Math.Abs(neto[i]) >= 1)
                {
                    // baris tanpa nama (dianggap baris agunan) / kualitas kosong tapi bernilai → masuk neraca, tidak masuk OS segmen
                    k.Dikecualikan += neto[i];
                    k.NDikecualikan++;
                    string raw = i < kualArr.Length ? (kualArr[i] ?? "").Trim() : "";
                    kandidat.Add(Contoh(dataStart + i, rekArr, namaArr, kualArr, i, neto[i],
                        !adaNama ? "tanpa nama debitur (dibaca sebagai baris agunan)"
                                 : "kualitas " + (raw.Length == 0 ? "kosong" : "'" + raw + "' tidak valid")));
                }
            }
            // simpan contoh terbesar saja
            kandidat.Sort((a, b) => Math.Abs((double)b["nilai"]).CompareTo(Math.Abs((double)a["nilai"])));
            if (kandidat.Count > BatasContoh) kandidat.RemoveRange(BatasContoh, kandidat.Count - BatasContoh);
            k.Contoh = kandidat;
            return k;
        }

        private static Dictionary<string, object> Contoh(int baris, string[] rek, string[] nama, string[] kual,
                                                         int i, double nilai, string alasan)
        {
            return new Dictionary<string, object>
            {
                { "baris", baris },
                { "rek", i < rek.Length ? (rek[i] ?? "").Trim() : "" },
                { "nama", i < nama.Length ? (nama[i] ?? "").Trim() : "" },
                { "kualitas", i < kual.Length ? (kual[i] ?? "").Trim() : "" },
                { "nilai", nilai },
                { "alasan", alasan }
            };
        }

        // ================================================================
        // 2. Baca GB0200 (nilai sel E3..E40)
        // ================================================================
        public static Dictionary<int, double> BacaGB0200(Excel.Worksheet ws)
        {
            var g = new Dictionary<int, double>();
            object[,] raw = ExcelHelper.BacaRangeAman((Excel.Range)ws.Range["E1", "E40"], 40);
            for (int r = 1; r <= 40; r++) g[r] = ToDouble(raw[r, 1]);
            return g;
        }

        // ================================================================
        // 3. Susun hasil untuk panel (masuk ke snapshot overview → bisa dibaca ulang tanpa template)
        // ================================================================
        public static Dictionary<string, object> Susun(Dictionary<int, double> g, double osNeraca,
                                                       Dictionary<string, Segmen> segmen, Dictionary<string, string> pesanSegmen)
        {
            Func<int, double> E = r => g.ContainsKey(r) ? g[r] : 0;
            var daftar = new List<object>();
            double totNeto = 0, totKC = 0, totOS = 0, totSumber = 0, totCakupan = 0;

            foreach (string kode in Urutan)
            {
                Segmen k;
                segmen.TryGetValue(kode, out k);
                bool ada = k != null && k.Ada;
                Func<string, double> K = c => ada && k.Kolom.ContainsKey(c) ? k.Kolom[c] : 0;

                var pos = new List<object>();
                var info = new List<object>();
                double netoNeraca;
                switch (kode)
                {
                    case "KC0600":
                        pos.Add(Pos("E8", 131, "Piutang Murabahah", E(8), "AE", K("AE"), false));
                        pos.Add(Pos("E9", 132, "Pendapatan margin Murabahah ditangguhkan", E(9), "AD", K("AD"), true));
                        netoNeraca = E(8) - E(9);
                        break;
                    case "KC0700":
                        pos.Add(Pos("E10", 141, "Piutang Istishna", E(10), "AE", K("AE"), false));
                        pos.Add(Pos("E11", 142, "Pendapatan margin Istishna ditangguhkan", E(11), "AD", K("AD"), true));
                        netoNeraca = E(10) - E(11);
                        break;
                    case "KC0800":
                        pos.Add(Pos("E12", 143, "Piutang Multijasa", E(12), "AE", K("AE"), false));
                        pos.Add(Pos("E13", 144, "Pendapatan margin Multijasa ditangguhkan", E(13), "AD", K("AD"), true));
                        netoNeraca = E(12) - E(13);
                        break;
                    case "KC0900":
                        pos.Add(Pos("E14", 150, "Piutang Qardh", E(14), "AC", K("AC"), false));
                        netoNeraca = E(14);
                        break;
                    case "KC1000":
                        pos.Add(Pos("E17", 171, "Mudarabah", E(17), "AI (akad < 29)", K("AI_MUD"), false));
                        pos.Add(Pos("E18", 172, "Musyarakah", E(18), "AI (akad > 29)", K("AI_MUS"), false));
                        pos.Add(Pos("E19", 173, "Bagi hasil lainnya", E(19), "AI (akad lain)", K("AI_LAIN"), false));
                        netoNeraca = E(17) + E(18) + E(19);
                        break;
                    default: // KC1100
                        pos.Add(Pos("E25", 181, "Aset Ijarah", E(25), "X", K("X"), false));
                        pos.Add(Pos("E26", 182, "Akumulasi penyusutan/amortisasi", E(26), "AI", K("AI"), true));
                        pos.Add(Pos("E15", 160, "Piutang sewa (tunggakan pokok)", E(15), "AL", K("AL"), false));
                        netoNeraca = E(25) - E(26) + E(15);   // baki debet ijarah = sandi 181 − 182 + 160
                        // CKPN aset ijarah & tunggakan ujroh: informasi, bukan OS
                        info.Add(new Dictionary<string, object>
                        {
                            { "sel", "E27" }, { "sandi", 183 }, { "nama", "CKPN aset Ijarah" },
                            { "neraca", E(27) }, { "kolomKC", "AJ" }, { "kc", K("AJ") }, { "selisih", E(27) - K("AJ") },
                            { "ket", "Cadangan, tidak mengurangi baki debet ijarah." }
                        });
                        info.Add(new Dictionary<string, object>
                        {
                            { "sel", "" }, { "sandi", null }, { "nama", "Tunggakan ujroh/imbalan" },
                            { "neraca", null }, { "kolomKC", "AM" }, { "kc", K("AM") }, { "selisih", null },
                            { "ket", "Tidak ada pos di neraca. Menjadi objek CKPN (EAD) KC1100 hanya untuk kualitas 1–2." }
                        });
                        break;
                }

                double netoKC = ada ? k.NetoKC : 0, os = ada ? k.OSSegmen : 0;
                double selSumber = netoNeraca - netoKC, selCakupan = netoKC - os;
                totNeto += netoNeraca; totKC += netoKC; totOS += os; totSumber += selSumber; totCakupan += selCakupan;

                string pesan;
                pesanSegmen.TryGetValue(kode, out pesan);
                daftar.Add(new Dictionary<string, object>
                {
                    { "kode", kode }, { "nama", NamaSegmen(kode) }, { "ada", ada }, { "pesan", pesan ?? (ada ? "" : "tidak terbaca") },
                    { "netoNeraca", netoNeraca }, { "netoKC", netoKC }, { "osSegmen", os },
                    { "selisih", netoNeraca - os }, { "selSumber", selSumber }, { "selCakupan", selCakupan },
                    { "kolomNeto", KolomNeto(kode) }, { "kolomOS", KolomOS(kode) },
                    { "pos", pos }, { "info", info },
                    { "dikecualikan", ada ? k.Dikecualikan : 0 }, { "nDikecualikan", ada ? k.NDikecualikan : 0 },
                    { "bedaKolom", ada ? k.BedaKolom : 0 }, { "nBedaKolom", ada ? k.NBedaKolom : 0 },
                    { "contoh", ada ? (object)k.Contoh : new List<object>() }
                });
            }

            // Pos neraca yang masuk OS neraca tetapi tidak punya sheet segmen CKPN
            var tanpa = new List<object>
            {
                new Dictionary<string, object> { { "sel", "E21" }, { "sandi", 200 }, { "nama", "Salam" }, { "nilai", E(21) } }
            };
            double totTanpa = E(21);

            // Subtotal GB0200 vs rinciannya (biasanya nol; berbeda bila template diisi manual)
            var subtotal = new List<object>();
            double rinciE7 = (E(8) - E(9)) + (E(10) - E(11)) + (E(12) - E(13)) + E(14) + E(15);
            double rinciE16 = E(17) + E(18) + E(19);
            subtotal.Add(Subtotal("E7", "Piutang", E(7), rinciE7, "E8 − E9 + E10 − E11 + E12 − E13 + E14 + E15"));
            subtotal.Add(Subtotal("E16", "Pembiayaan bagi hasil", E(16), rinciE16, "E17 + E18 + E19"));
            double totSubtotal = osNeraca - (totNeto + totTanpa);

            return new Dictionary<string, object>
            {
                { "rumusNeraca", "(E7 − E15) + E16 + E21 + (E25 − E26 + E15)" },
                { "osNeraca", osNeraca }, { "netoNeraca", totNeto }, { "netoKC", totKC }, { "osSegmen", totOS },
                { "selisih", osNeraca - totOS },
                { "komponen", new Dictionary<string, object>
                    {
                        { "tanpaSegmen", totTanpa }, { "sumber", totSumber },
                        { "cakupan", totCakupan }, { "subtotal", totSubtotal }
                    } },
                { "segmen", daftar }, { "tanpaSegmen", tanpa }, { "subtotal", subtotal }
            };
        }

        private static Dictionary<string, object> Pos(string sel, int sandi, string nama, double neraca,
                                                      string kolomKC, double kc, bool pengurang)
        {
            return new Dictionary<string, object>
            {
                { "sel", sel }, { "sandi", sandi }, { "nama", nama }, { "neraca", neraca },
                { "kolomKC", kolomKC }, { "kc", kc }, { "selisih", neraca - kc }, { "pengurang", pengurang }
            };
        }

        private static Dictionary<string, object> Subtotal(string sel, string nama, double nilai, double rinci, string rumus)
        {
            return new Dictionary<string, object>
            {
                { "sel", sel }, { "nama", nama }, { "nilai", nilai }, { "rinci", rinci },
                { "selisih", nilai - rinci }, { "rumus", rumus }
            };
        }

        // ================================================================
        // Util (sama dengan DataOverviewBuilder supaya hasilnya identik)
        // ================================================================
        private static double Jumlah(double[] a) { double t = 0; foreach (var v in a) t += v; return t; }

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

        private static double ToDouble(object val)
        {
            if (val == null) return 0;
            if (val is double d) return d;
            if (val is int i) return i;
            if (val is long l) return l;
            double r;
            return double.TryParse(val.ToString(), out r) ? r : 0;
        }
    }
}
