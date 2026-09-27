using System;
using CKPNLibrary.Helpers;
using ExcelDna.Integration;

namespace CKPNLibrary.Panel
{
    /// <summary>
    /// Siklus hidup bagian panel. Excel-DNA memanggil AutoOpen/AutoClose pada
    /// SEMUA kelas yang mengimplementasikan IExcelAddIn, jadi kelas ini berdiri
    /// sendiri dan CKPNAddIn di CKPNFunctions.cs tidak perlu diubah.
    /// </summary>
    public class PanelAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            // Siapkan folder library\data, config, backup, logs dan folder lokal.
            // Tidak melempar error: kegagalan folder tidak boleh menggagalkan add-in.
            AppPaths.SiapkanFolder();
        }

        public void AutoClose()
        {
            try { PanelManager.TutupSemua(); } catch { }
        }
    }

    /// <summary>
    /// Perintah Excel untuk panel — mengikuti pola [ExcelCommand] di CKPNFunctions.
    /// </summary>
    public static class PanelCommands
    {
        // ----------------------------------------------------------------
        // Dipanggil VBA (tombol "Panel CKPN" di sheet Master):
        //   Application.Run "CKPN_BukaPanel"
        // Selalu MENAMPILKAN (bukan toggle) supaya tombol di sheet tidak
        // malah menutup panel yang sedang terbuka.
        // ----------------------------------------------------------------
        [ExcelCommand(Name = "CKPN_BukaPanel")]
        public static void BukaPanel()
        {
            try { PanelManager.Tampilkan(); }
            catch (Exception ex)
            {
                PanelHost.CatatLog("CKPN_BukaPanel: " + ex);
                System.Windows.Forms.MessageBox.Show(
                    "Panel CKPN gagal dibuka:\n" + ex.Message, "CKPNLibrary",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
    }
}
