using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;

namespace CKPNLibrary.Panel
{
    /// <summary>
    /// Tab ribbon "CKPN 414" dengan tombol pembuka Panel CKPN.
    /// Excel-DNA otomatis memuat kelas turunan ExcelRibbon yang public dan
    /// ComVisible saat .xll dimuat (RegisterXLL dari VBA atau Workbook_Open).
    /// </summary>
    [ComVisible(true)]
    public class CKPNRibbon : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            return
@"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='tabCKPN414' label='CKPN 414'>
        <group id='grpPanelCKPN' label='Panel'>
          <button id='btnPanelCKPN'
                  label='Panel CKPN'
                  size='large'
                  imageMso='ChartTypeAllInsertDialog'
                  screentip='Panel CKPN'
                  supertip='Buka atau tutup Panel CKPN: jalankan perhitungan, lihat periode, riwayat, dan analisis.'
                  onAction='OnPanelCKPN' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void OnPanelCKPN(IRibbonControl control)
        {
            try { PanelManager.Toggle(); }
            catch (System.Exception ex)
            {
                PanelHost.CatatLog("Ribbon OnPanelCKPN: " + ex);
                System.Windows.Forms.MessageBox.Show(
                    "Panel CKPN gagal dibuka:\n" + ex.Message, "CKPNLibrary",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
    }
}
