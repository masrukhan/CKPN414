using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Styling
{
    internal static class StyleHelper
    {
        private const int WarnaPrimary    = 0x436E1F; // RGB(31,110,67)
        private const int WarnaRowAlt     = 0xF4F8F1; // RGB(241,248,244)
        private const int WarnaRowNormal  = 0xFFFFFF;
        private const int WarnaInputKuning= 0xE7FDFF; // RGB(255,253,231)
        private const int WarnaFontPutih  = 0xFFFFFF;
        private const int WarnaFontHitam  = 0x000000;

        public static void StyleTabelCKPN(Excel.Worksheet ws, int totalRows)
        {
            const int dataStart = 6;
            int dataEnd  = dataStart + totalRows - 1;
            int totalRow = dataEnd + 1;

            // Banded rows berdasarkan nilai kolom B (nomor urut CIF)
            object prevCif = "";
            int toggle = 0;
            for (int r = dataStart; r <= dataEnd; r++)
            {
                // Cast eksplisit ke Excel.Range — wajib
                Excel.Range cellB    = (Excel.Range)ws.Cells[r, "B"];
                object      curCif   = cellB.Value2;

                if (r == dataStart)
                    toggle = 0;
                else if ((curCif?.ToString() ?? "") != (prevCif?.ToString() ?? ""))
                    toggle = 1 - toggle;

                Excel.Range rowRange = (Excel.Range)ws.Range["B" + r, "K" + r];
                rowRange.Interior.Color = toggle == 0 ? WarnaRowNormal : WarnaRowAlt;
                rowRange.Font.Name      = "Calibri";
                rowRange.Font.Size      = 11;
                rowRange.Font.Color     = WarnaFontHitam;
                prevCif = curCif;
            }

            // Alignment
            ((Excel.Range)ws.Range["B" + dataStart, "D" + dataEnd]).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ((Excel.Range)ws.Range["E" + dataStart, "E" + dataEnd]).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            ((Excel.Range)ws.Range["F" + dataStart, "F" + dataEnd]).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ((Excel.Range)ws.Range["H" + dataStart, "H" + dataEnd]).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            // Format angka
            string fmtAngka = "#,##0;(#,##0);-";
            ((Excel.Range)ws.Range["G" + dataStart, "G" + dataEnd]).NumberFormat = fmtAngka;
            ((Excel.Range)ws.Range["I" + dataStart, "K" + dataEnd]).NumberFormat = fmtAngka;

            // Kolom D & F teks (leading-zero)
            ((Excel.Range)ws.Range["D" + dataStart, "D" + dataEnd]).NumberFormat = "@";
            ((Excel.Range)ws.Range["F" + dataStart, "F" + dataEnd]).NumberFormat = "@";

            // Kolom J input user: kuning lembut
            ((Excel.Range)ws.Range["J" + dataStart, "J" + dataEnd]).Interior.Color = WarnaInputKuning;

            // Border data
            Excel.Borders borders = ((Excel.Range)ws.Range["B" + dataStart, "K" + dataEnd]).Borders;
            borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            borders.Color     = WarnaPrimary;
            borders.Weight    = Excel.XlBorderWeight.xlThin;

            // Baris TOTAL
            Excel.Range totalRange = (Excel.Range)ws.Range["B" + totalRow, "K" + totalRow];
            totalRange.Interior.Color      = WarnaPrimary;
            totalRange.Font.Color          = WarnaFontPutih;
            totalRange.Font.Bold           = true;
            totalRange.Font.Size           = 11;
            totalRange.Font.Name           = "Calibri";
            totalRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            totalRange.Borders.LineStyle   = Excel.XlLineStyle.xlContinuous;
            totalRange.Borders.Color       = WarnaPrimary;
            totalRange.Borders.Weight      = Excel.XlBorderWeight.xlMedium;

            ws.Application.DisplayAlerts = false;
            ((Excel.Range)ws.Range["B" + totalRow, "F" + totalRow]).Merge();
            ws.Application.DisplayAlerts = true;
            ((Excel.Range)ws.Cells[totalRow, "B"]).Value2 = "TOTAL";

            string fmtSum = "#,##0;(#,##0);-";
            ((Excel.Range)ws.Cells[totalRow, "G"]).Formula = "=SUM(G" + dataStart + ":G" + dataEnd + ")";
            ((Excel.Range)ws.Cells[totalRow, "I"]).Formula = "=SUM(I" + dataStart + ":I" + dataEnd + ")";
            ((Excel.Range)ws.Cells[totalRow, "J"]).Formula = "=SUM(J" + dataStart + ":J" + dataEnd + ")";
            ((Excel.Range)ws.Cells[totalRow, "K"]).Formula = "=SUM(K" + dataStart + ":K" + dataEnd + ")";
            ((Excel.Range)ws.Range["G" + totalRow, "K" + totalRow]).NumberFormat = fmtSum;
        }
    }
}