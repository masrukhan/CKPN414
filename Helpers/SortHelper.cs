using System.Collections.Generic;
using CKPNLibrary.Models;

namespace CKPNLibrary.Helpers
{
    /// <summary>
    /// Utilitas pengurutan — setara dengan SortDescByCol4 dan QSortDescCol2 di VBA.
    /// </summary>
    internal static class SortHelper
    {
        // ----------------------------------------------------------------
        // Sort List<CIFData> descending by TotalOS (in-place QuickSort)
        // Setara SortDescByCol4 di VBA.
        // ----------------------------------------------------------------
        public static void SortDescByTotalOS(List<CIFData> list, int lo, int hi)
        {
            if (lo >= hi) return;
            int i = lo, j = hi;
            double pivot = list[(lo + hi) / 2].TotalOS;
            while (i <= j)
            {
                while (list[i].TotalOS > pivot) i++;
                while (list[j].TotalOS < pivot) j--;
                if (i <= j)
                {
                    var tmp = list[i]; list[i] = list[j]; list[j] = tmp;
                    i++; j--;
                }
            }
            if (lo < j) SortDescByTotalOS(list, lo, j);
            if (i < hi) SortDescByTotalOS(list, i, hi);
        }

        // ----------------------------------------------------------------
        // Sort array of (cif, os) pairs descending by os.
        // Setara QSortDescCol2 di VBA.
        // Digunakan untuk bangun Top-N di PD Net Flow.
        // ----------------------------------------------------------------
        public static void SortDescByOS(List<CifOsPair> list, int lo, int hi)
        {
            if (lo >= hi) return;
            int i = lo, j = hi;
            double pivot = list[(lo + hi) / 2].OS;
            while (i <= j)
            {
                while (list[i].OS > pivot) i++;
                while (list[j].OS < pivot) j--;
                if (i <= j)
                {
                    var tmp = list[i]; list[i] = list[j]; list[j] = tmp;
                    i++; j--;
                }
            }
            if (lo < j) SortDescByOS(list, lo, j);
            if (i < hi) SortDescByOS(list, i, hi);
        }

        /// <summary>Pasangan CIF + total OS untuk keperluan ranking Top-N.</summary>
        public struct CifOsPair
        {
            public string CIF;
            public double OS;
        }
    }
}