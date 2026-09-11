namespace CKPNLibrary.Models
{
    /// <summary>
    /// Mode pembacaan sheet KC — menentukan kolom mana yang dipakai
    /// dan bagaimana OS serta filter hari tunggakan dihitung.
    /// </summary>
    public enum SheetMode
    {
        /// <summary>KC0600, KC0700, KC0800, KC0900: filter pakai kolom AF (hari tunggakan).</summary>
        AF,

        /// <summary>
        /// KC1100 (LOGIKA BARU): masuk jika ada nama debitur (kolom D tidak kosong).
        /// OS/EAD = tunggakan pokok (AL) + tunggakan ujroh/imbalan (AM) tanpa syarat hari.
        /// Hari tunggakan gabungan pokok & ujroh ada di AK (ColHariPokok) →
        /// dipakai sebagai basis bucket PD Net Flow.
        /// Baris tanpa nama = baris agunan, di-skip.
        /// </summary>
        Tunggakan1,

        /// <summary>
        /// KC1000 (LOGIKA BARU): masuk jika ada nama debitur (kolom D tidak kosong).
        /// OS = jumlah seluruh AL (tunggakan pokok) + AN (tunggakan basil) tanpa syarat hari.
        /// Baris tanpa nama = baris agunan, di-skip.
        /// </summary>
        Tunggakan2
    }

    /// <summary>
    /// Spesifikasi kolom satu sheet KC.
    /// Setara dengan spec(idx) = Array(shName, mode, colOS, colHari, colHariBasil, colNomBasil) di VBA.
    /// </summary>
    public class SheetSpec
    {
        public string    SheetName    { get; set; }
        public SheetMode Mode         { get; set; }

        // Kolom-kolom yang dibaca (nama kolom Excel: "C", "AC", dll.)
        public string ColCIF          { get; set; }   // selalu "C"
        public string ColNama         { get; set; }   // selalu "D"
        public string ColNoRek        { get; set; }   // selalu "J"
        public string ColOS           { get; set; }   // OS / NomPokok
        public string ColJaminan      { get; set; }   // kolom nilai jaminan
        public string ColKualitas     { get; set; }   // kolom kualitas

        // Khusus mode AF: kolom hari tunggakan
        public string ColHariAF       { get; set; }   // mis. "AF"

        // Khusus Tunggakan1 & Tunggakan2: hari & nominal pokok
        public string ColHariPokok    { get; set; }   // mis. "AK"
        public string ColNomPokok     { get; set; }   // mis. "AL"

        // Khusus Tunggakan2 saja: hari & nominal basil
        public string ColHariBasil    { get; set; }   // mis. "AM"
        public string ColNomBasil     { get; set; }   // mis. "AN"

        // Khusus Tunggakan1 (KC1100 ijarah): nominal tunggakan ujroh/imbalan.
        // Aging-nya menyatu dengan pokok di ColHariPokok (AK) — tidak butuh kolom hari sendiri.
        public string ColNomUjroh     { get; set; }   // mis. "AM"

        // Header row sumber (baris data mulai di HEADER_ROW + 1)
        public int HeaderRow          { get; set; } = 4;

        // ----------------------------------------------------------------
        // Factory: buat spec berdasarkan nama sheet
        // ----------------------------------------------------------------
        public static SheetSpec Buat(string sheetName)
        {
            switch (sheetName)
            {
                case "KC0600":
                case "KC0700":
                case "KC0800":
                    return new SheetSpec
                    {
                        SheetName  = sheetName,
                        Mode       = SheetMode.AF,
                        ColCIF     = "C", ColNama = "D", ColNoRek = "J",
                        ColOS      = "AC",
                        ColJaminan = "AS",
                        ColKualitas= "Y",
                        ColHariAF  = "AF"
                    };

                case "KC0900":
                    return new SheetSpec
                    {
                        SheetName  = sheetName,
                        Mode       = SheetMode.AF,
                        ColCIF     = "C", ColNama = "D", ColNoRek = "J",
                        ColOS      = "AA",
                        ColJaminan = "AQ",
                        ColKualitas= "Y",
                        ColHariAF  = "AF"
                    };

                case "KC1000":
                    // Filter: baris dengan nama debitur (kolom D tidak kosong)
                    // OS: jumlah AL (pokok) + AN (basil) — semua, tanpa syarat hari
                    // ColHariPokok & ColHariBasil tetap ada untuk PD Net Flow (bucket hari)
                    return new SheetSpec
                    {
                        SheetName   = sheetName,
                        Mode        = SheetMode.Tunggakan2,
                        ColCIF      = "C", ColNama = "D", ColNoRek = "J",
                        ColJaminan  = "AY",
                        ColKualitas = "AE",
                        ColHariPokok= "AK", ColNomPokok = "AL",
                        ColHariBasil= "AM", ColNomBasil = "AN"
                    };

                case "KC1100":
                    // Filter: baris dengan nama debitur (kolom D tidak kosong)
                    // OS/EAD: tunggakan pokok (AL) + tunggakan ujroh/imbalan (AM) — semua, tanpa syarat hari
                    // ColHariPokok (AK) = hari tunggakan gabungan pokok & ujroh → basis bucket PD Net Flow
                    // CATATAN: di sini "AM" = NOMINAL ujroh. Berbeda peran dgn KC1000 (Tunggakan2)
                    //          yang memakai AM sebagai ColHariBasil (hari tunggakan basil).
                    // PERBAIKAN: ColKualitas = AF (bukan AE)
                    return new SheetSpec
                    {
                        SheetName   = sheetName,
                        Mode        = SheetMode.Tunggakan1,
                        ColCIF      = "C", ColNama = "D", ColNoRek = "J",
                        ColJaminan  = "AY",
                        ColKualitas = "AF",
                        ColHariPokok= "AK", ColNomPokok = "AL",
                        ColNomUjroh = "AM"
                    };

                default:
                    throw new System.ArgumentException("Sheet tidak dikenal: " + sheetName);
            }
        }
    }
}