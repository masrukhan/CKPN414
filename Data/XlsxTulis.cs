using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace CKPNLibrary.Data
{
    /// <summary>
    /// Penulis .xlsx sederhana TANPA Excel (zip + XML SpreadsheetML).
    ///
    /// Dipakai kertas kerja versi (<see cref="KertasKerjaVersi"/>): file dibuat langsung dari database, jadi
    /// tidak membuka/mengganggu workbook aplikasi, tidak bergantung pada thread Excel, dan cukup cepat untuk
    /// dibuat otomatis setiap kali versi baru lahir dari koreksi panel.
    ///
    /// Yang didukung hanya yang diperlukan kertas kerja: teks, angka, rumus (dengan nilai cache), gaya sel
    /// dari daftar tetap, lebar kolom, freeze pane, dan gabung sel. Rumus ditulis TANPA tanda "=" di XML
    /// (aturan SpreadsheetML); workbook ditandai fullCalcOnLoad sehingga Excel menghitung ulang saat dibuka.
    /// Nilai cache tetap ditulis agar angka terlihat juga di penampil yang tidak menghitung rumus.
    /// </summary>
    internal sealed class XlsxTulis
    {
        // Urutan HARUS sama dengan <cellXfs> di GayaXml()
        public enum Gaya
        {
            Biasa = 0,
            Tebal = 1,
            Kepala = 2,        // judul kolom: tebal putih, latar hijau tua, rata tengah, wrap, border
            Angka = 3,         // #,##0 (hasil rumus)
            AngkaTebal = 4,
            Persen = 5,        // 0.00%
            PersenTebal = 6,
            Input = 7,         // #,##0 biru — data dari database / isian (boleh diubah user)
            Judul = 8,         // tebal 14
            Catatan = 9,       // miring abu
            Teks = 10,         // format teks (@) — no. rekening / kontrak
            Total = 11,        // #,##0 tebal, latar hijau muda, border
            PersenInput = 12,  // 0.00% biru
            Pd = 13,           // 0.0000% (hasil rumus)
            PdInput = 14,      // 0.0000% biru
            Nilai = 15,        // #,##0 latar kuning — angka tersimpan dipakai apa adanya (rumus tidak cocok/tidak tersedia)
            PersenNilai = 16,  // 0.0000% latar kuning
            Subjudul = 17,     // tebal hijau tua
            TotalPersen = 18,  // 0.00% tebal, latar hijau muda, border
            TotalLabel = 19,   // teks tebal, latar hijau muda, border
            Bungkus = 20,      // teks wrap, rata atas
            Peringatan = 21    // teks tebal merah
        }

        internal sealed class Sel
        {
            public string Teks;       // isi teks (inlineStr) bila bukan angka/rumus
            public double? Angka;
            public string Rumus;      // tanpa "="
            public object Cache;      // double / string / null
            public Gaya Gaya;
        }

        public sealed class Lembar
        {
            public readonly string Nama;
            internal readonly SortedDictionary<int, SortedDictionary<int, Sel>> Baris = new SortedDictionary<int, SortedDictionary<int, Sel>>();
            internal readonly SortedDictionary<int, double> LebarKolom = new SortedDictionary<int, double>();
            internal readonly List<string> Gabung = new List<string>();
            internal int BekuBaris, BekuKolom;
            internal string Warna;   // warna tab (ARGB) atau null

            internal Lembar(string nama) { Nama = nama; }

            private Sel Ambil(int r, int c)
            {
                if (r < 1 || c < 1) throw new ArgumentOutOfRangeException("Sel " + r + "," + c);
                SortedDictionary<int, Sel> baris;
                if (!Baris.TryGetValue(r, out baris)) { baris = new SortedDictionary<int, Sel>(); Baris[r] = baris; }
                Sel s;
                if (!baris.TryGetValue(c, out s)) { s = new Sel(); baris[c] = s; }
                return s;
            }

            public void Teks(int r, int c, string teks, Gaya g = Gaya.Biasa)
            {
                var s = Ambil(r, c);
                s.Teks = teks ?? ""; s.Angka = null; s.Rumus = null; s.Gaya = g;
            }

            /// <summary>Angka; null → sel kosong bergaya (tetap tampil sebagai sel berformat).</summary>
            public void Angka(int r, int c, double? nilai, Gaya g = Gaya.Angka)
            {
                var s = Ambil(r, c);
                s.Teks = null; s.Angka = nilai.HasValue && !double.IsNaN(nilai.Value) && !double.IsInfinity(nilai.Value) ? nilai : null;
                s.Rumus = null; s.Gaya = g;
            }

            /// <summary>Rumus (boleh diawali "="). cache: nilai hasil rumus menurut perhitungan aplikasi.</summary>
            public void Rumus(int r, int c, string rumus, object cache, Gaya g = Gaya.Angka)
            {
                var s = Ambil(r, c);
                s.Teks = null; s.Angka = null;
                s.Rumus = rumus != null && rumus.StartsWith("=") ? rumus.Substring(1) : rumus;
                if (cache is double && (double.IsNaN((double)cache) || double.IsInfinity((double)cache))) cache = null;
                s.Cache = cache; s.Gaya = g;
            }

            public void Gayakan(int r, int c, Gaya g) { Ambil(r, c).Gaya = g; }

            public void Lebar(int c, double lebar) { LebarKolom[c] = lebar; }
            public void Beku(int baris, int kolom) { BekuBaris = baris; BekuKolom = kolom; }
            public void GabungSel(int r1, int c1, int r2, int c2) { Gabung.Add(Ref(r1, c1) + ":" + Ref(r2, c2)); }
            public void WarnaTab(string argb) { Warna = argb; }
        }

        private readonly List<Lembar> _lembar = new List<Lembar>();

        public Lembar Tambah(string nama)
        {
            var l = new Lembar(nama);
            _lembar.Add(l);
            return l;
        }

        // ================================================================
        // Util referensi sel
        // ================================================================
        public static string Kol(int c)
        {
            var sb = new StringBuilder();
            while (c > 0) { int m = (c - 1) % 26; sb.Insert(0, (char)('A' + m)); c = (c - 1) / 26; }
            return sb.ToString();
        }
        public static string Ref(int r, int c) { return Kol(c) + r; }
        /// <summary>'Nama Sheet'!$A$1 — referensi mutlak ke sheet lain.</summary>
        public static string RefLuar(string sheet, int r, int c) { return "'" + sheet.Replace("'", "''") + "'!$" + Kol(c) + "$" + r; }

        // ================================================================
        // Simpan
        // ================================================================
        public void Simpan(string path)
        {
            if (_lembar.Count == 0) throw new InvalidOperationException("Workbook tanpa sheet.");
            string tmp = path + ".tmp";
            if (File.Exists(tmp)) File.Delete(tmp);
            using (var fs = new FileStream(tmp, FileMode.CreateNew))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Create))
            {
                Tulis(zip, "[Content_Types].xml", ContentTypes());
                Tulis(zip, "_rels/.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                    "</Relationships>");
                Tulis(zip, "xl/workbook.xml", Workbook());
                Tulis(zip, "xl/_rels/workbook.xml.rels", WorkbookRels());
                Tulis(zip, "xl/styles.xml", GayaXml());
                for (int i = 0; i < _lembar.Count; i++)
                    Tulis(zip, "xl/worksheets/sheet" + (i + 1) + ".xml", SheetXml(_lembar[i], i == 0));
            }
            // Ganti file lama hanya setelah file baru lengkap (file lama yang sedang dibuka → gagal di sini, bukan setengah jadi)
            try
            {
                if (File.Exists(path)) File.Delete(path);
                File.Move(tmp, path);
            }
            catch
            {
                try { File.Delete(tmp); } catch { }
                throw;
            }
        }

        private static void Tulis(ZipArchive zip, string nama, string isi)
        {
            var e = zip.CreateEntry(nama, CompressionLevel.Optimal);
            using (var s = e.Open())
            {
                var b = new UTF8Encoding(false).GetBytes(isi);
                s.Write(b, 0, b.Length);
            }
        }

        private string ContentTypes()
        {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
            sb.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
            sb.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
            sb.Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
            sb.Append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
            for (int i = 0; i < _lembar.Count; i++)
                sb.Append("<Override PartName=\"/xl/worksheets/sheet" + (i + 1) +
                          ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
            sb.Append("</Types>");
            return sb.ToString();
        }

        private string Workbook()
        {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
                      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            sb.Append("<bookViews><workbookView activeTab=\"0\"/></bookViews><sheets>");
            for (int i = 0; i < _lembar.Count; i++)
                sb.Append("<sheet name=\"" + Esc(_lembar[i].Nama) + "\" sheetId=\"" + (i + 1) + "\" r:id=\"rId" + (i + 1) + "\"/>");
            sb.Append("</sheets>");
            // Hitung ulang penuh saat dibuka: rumus hidup selalu menampilkan hasil terkini
            sb.Append("<calcPr calcId=\"191029\" fullCalcOnLoad=\"1\"/>");
            sb.Append("</workbook>");
            return sb.ToString();
        }

        private string WorkbookRels()
        {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            for (int i = 0; i < _lembar.Count; i++)
                sb.Append("<Relationship Id=\"rId" + (i + 1) + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" " +
                          "Target=\"worksheets/sheet" + (i + 1) + ".xml\"/>");
            sb.Append("<Relationship Id=\"rId" + (_lembar.Count + 1) + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
            sb.Append("</Relationships>");
            return sb.ToString();
        }

        private static string GayaXml()
        {
            // numFmt 164 = 0.0000% (PD). Bawaan: 3 = #,##0, 10 = 0.00%, 49 = @
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
                "<numFmts count=\"1\"><numFmt numFmtId=\"164\" formatCode=\"0.0000%\"/></numFmts>" +
                "<fonts count=\"8\">" +
                "<font><sz val=\"10\"/><name val=\"Calibri\"/><family val=\"2\"/></font>" +                                  // 0 biasa
                "<font><b/><sz val=\"10\"/><name val=\"Calibri\"/><family val=\"2\"/></font>" +                              // 1 tebal
                "<font><b/><sz val=\"10\"/><color rgb=\"FFFFFFFF\"/><name val=\"Calibri\"/><family val=\"2\"/></font>" +     // 2 tebal putih
                "<font><sz val=\"10\"/><color rgb=\"FF1F4E9A\"/><name val=\"Calibri\"/><family val=\"2\"/></font>" +         // 3 biru (input)
                "<font><i/><sz val=\"9\"/><color rgb=\"FF6B6B6B\"/><name val=\"Calibri\"/><family val=\"2\"/></font>" +      // 4 catatan
                "<font><b/><sz val=\"14\"/><color rgb=\"FF1B5E20\"/><name val=\"Calibri\"/><family val=\"2\"/></font>" +     // 5 judul
                "<font><b/><sz val=\"11\"/><color rgb=\"FF1B5E20\"/><name val=\"Calibri\"/><family val=\"2\"/></font>" +     // 6 subjudul
                "<font><b/><sz val=\"10\"/><color rgb=\"FFB71C1C\"/><name val=\"Calibri\"/><family val=\"2\"/></font>" +     // 7 peringatan
                "</fonts>" +
                "<fills count=\"5\">" +
                "<fill><patternFill patternType=\"none\"/></fill>" +
                "<fill><patternFill patternType=\"gray125\"/></fill>" +
                "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF2E7D32\"/><bgColor indexed=\"64\"/></patternFill></fill>" +   // 2 hijau tua
                "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFE8F5E9\"/><bgColor indexed=\"64\"/></patternFill></fill>" +   // 3 hijau muda
                "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFFF3C4\"/><bgColor indexed=\"64\"/></patternFill></fill>" +   // 4 kuning
                "</fills>" +
                "<borders count=\"2\">" +
                "<border><left/><right/><top/><bottom/><diagonal/></border>" +
                "<border><left style=\"thin\"><color rgb=\"FFBDBDBD\"/></left><right style=\"thin\"><color rgb=\"FFBDBDBD\"/></right>" +
                "<top style=\"thin\"><color rgb=\"FFBDBDBD\"/></top><bottom style=\"thin\"><color rgb=\"FFBDBDBD\"/></bottom><diagonal/></border>" +
                "</borders>" +
                "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>" +
                "<cellXfs count=\"22\">" +
                Xf(0, 0, 0, 0) +                                                      // 0 Biasa
                Xf(0, 1, 0, 0) +                                                      // 1 Tebal
                Xf(0, 2, 2, 1, "<alignment horizontal=\"center\" vertical=\"center\" wrapText=\"1\"/>") + // 2 Kepala
                Xf(3, 0, 0, 1) +                                                      // 3 Angka
                Xf(3, 1, 0, 1) +                                                      // 4 AngkaTebal
                Xf(10, 0, 0, 1) +                                                     // 5 Persen
                Xf(10, 1, 0, 1) +                                                     // 6 PersenTebal
                Xf(3, 3, 0, 1) +                                                      // 7 Input
                Xf(0, 5, 0, 0) +                                                      // 8 Judul
                Xf(0, 4, 0, 0) +                                                      // 9 Catatan
                Xf(49, 0, 0, 1) +                                                     // 10 Teks
                Xf(3, 1, 3, 1) +                                                      // 11 Total
                Xf(10, 3, 0, 1) +                                                     // 12 PersenInput
                Xf(164, 0, 0, 1) +                                                    // 13 Pd
                Xf(164, 3, 0, 1) +                                                    // 14 PdInput
                Xf(3, 0, 4, 1) +                                                      // 15 Nilai
                Xf(164, 0, 4, 1) +                                                    // 16 PersenNilai
                Xf(0, 6, 0, 0) +                                                      // 17 Subjudul
                Xf(10, 1, 3, 1) +                                                     // 18 TotalPersen
                Xf(0, 1, 3, 1) +                                                      // 19 TotalLabel
                Xf(0, 0, 0, 0, "<alignment vertical=\"top\" wrapText=\"1\"/>") +      // 20 Bungkus
                Xf(0, 7, 0, 0) +                                                      // 21 Peringatan
                "</cellXfs>" +
                "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>" +
                "</styleSheet>";
        }

        private static string Xf(int numFmt, int font, int fill, int border, string align = null)
        {
            string a = "<xf numFmtId=\"" + numFmt + "\" fontId=\"" + font + "\" fillId=\"" + fill + "\" borderId=\"" + border + "\" xfId=\"0\"" +
                       (numFmt != 0 ? " applyNumberFormat=\"1\"" : "") + (font != 0 ? " applyFont=\"1\"" : "") +
                       (fill != 0 ? " applyFill=\"1\"" : "") + (border != 0 ? " applyBorder=\"1\"" : "");
            return align == null ? a + "/>" : a + " applyAlignment=\"1\">" + align + "</xf>";
        }

        private static string SheetXml(Lembar l, bool aktif)
        {
            var sb = new StringBuilder(64 * 1024);
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
                      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            if (l.Warna != null) sb.Append("<sheetPr><tabColor rgb=\"" + l.Warna + "\"/></sheetPr>");
            sb.Append("<sheetViews><sheetView workbookViewId=\"0\" zoomScale=\"90\"" + (aktif ? " tabSelected=\"1\"" : "") + ">");
            if (l.BekuBaris > 0 || l.BekuKolom > 0)
            {
                string kiriAtas = Ref(l.BekuBaris + 1, l.BekuKolom + 1);
                string pane = l.BekuBaris > 0 && l.BekuKolom > 0 ? "bottomRight" : l.BekuBaris > 0 ? "bottomLeft" : "topRight";
                sb.Append("<pane" + (l.BekuKolom > 0 ? " xSplit=\"" + l.BekuKolom + "\"" : "") + (l.BekuBaris > 0 ? " ySplit=\"" + l.BekuBaris + "\"" : "") +
                          " topLeftCell=\"" + kiriAtas + "\" activePane=\"" + pane + "\" state=\"frozen\"/>");
                sb.Append("<selection pane=\"" + pane + "\" activeCell=\"" + kiriAtas + "\" sqref=\"" + kiriAtas + "\"/>");
            }
            sb.Append("</sheetView></sheetViews>");
            sb.Append("<sheetFormatPr defaultRowHeight=\"14\"/>");
            if (l.LebarKolom.Count > 0)
            {
                sb.Append("<cols>");
                foreach (var kv in l.LebarKolom)
                    sb.Append("<col min=\"" + kv.Key + "\" max=\"" + kv.Key + "\" width=\"" + Ang(kv.Value) + "\" customWidth=\"1\"/>");
                sb.Append("</cols>");
            }
            sb.Append("<sheetData>");
            foreach (var br in l.Baris)
            {
                sb.Append("<row r=\"" + br.Key + "\">");
                foreach (var kc in br.Value)
                {
                    var s = kc.Value;
                    string r = Ref(br.Key, kc.Key);
                    int g = (int)s.Gaya;
                    string st = g == 0 ? "" : " s=\"" + g + "\"";
                    if (s.Rumus != null)
                    {
                        if (s.Cache is string)
                            sb.Append("<c r=\"" + r + "\"" + st + " t=\"str\"><f>" + Esc(s.Rumus) + "</f><v>" + Esc((string)s.Cache) + "</v></c>");
                        else if (s.Cache is double)
                            sb.Append("<c r=\"" + r + "\"" + st + "><f>" + Esc(s.Rumus) + "</f><v>" + Ang((double)s.Cache) + "</v></c>");
                        else
                            sb.Append("<c r=\"" + r + "\"" + st + "><f>" + Esc(s.Rumus) + "</f></c>");
                    }
                    else if (s.Angka.HasValue)
                        sb.Append("<c r=\"" + r + "\"" + st + "><v>" + Ang(s.Angka.Value) + "</v></c>");
                    else if (!string.IsNullOrEmpty(s.Teks))
                        sb.Append("<c r=\"" + r + "\"" + st + " t=\"inlineStr\"><is><t xml:space=\"preserve\">" + Esc(s.Teks) + "</t></is></c>");
                    else
                        sb.Append("<c r=\"" + r + "\"" + st + "/>");
                }
                sb.Append("</row>");
            }
            sb.Append("</sheetData>");
            if (l.Gabung.Count > 0)
            {
                sb.Append("<mergeCells count=\"" + l.Gabung.Count + "\">");
                foreach (var m in l.Gabung) sb.Append("<mergeCell ref=\"" + m + "\"/>");
                sb.Append("</mergeCells>");
            }
            sb.Append("<pageMargins left=\"0.5\" right=\"0.5\" top=\"0.6\" bottom=\"0.6\" header=\"0.3\" footer=\"0.3\"/>");
            sb.Append("</worksheet>");
            return sb.ToString();
        }

        private static string Ang(double v) { return v.ToString("R", CultureInfo.InvariantCulture); }

        /// <summary>Escape XML + buang karakter kontrol yang tidak sah di XML 1.0.</summary>
        private static string Esc(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            var sb = new StringBuilder(s.Length + 16);
            foreach (char ch in s)
            {
                switch (ch)
                {
                    case '&': sb.Append("&amp;"); break;
                    case '<': sb.Append("&lt;"); break;
                    case '>': sb.Append("&gt;"); break;
                    case '"': sb.Append("&quot;"); break;
                    default:
                        if (ch < 0x20 && ch != '\t' && ch != '\n' && ch != '\r') continue;
                        if (ch == 0xFFFE || ch == 0xFFFF) continue;
                        sb.Append(ch);
                        break;
                }
            }
            return sb.ToString();
        }
    }
}
