using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace SAPOLEStatement
{
    public partial class SignatureRibbon
    {
        private void SignatureRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }


        private void group1_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            // Show or hide task pane
            // Globals.ThisAddIn.ToggleTaskPaneDisplay();
        }

        internal void SetToggleSignatureLock(bool isSAPOL)
        {
            btnToggleSignature.Enabled = isSAPOL;
            this.RibbonUI.Invalidate();
        }

        //internal void SetLocks(bool isReadOnly,string docPath)
        //{

        //    // Ensure checkbox is accurate
        //    btnLock.Enabled = !isReadOnly;
        //    chkFinalRevision.Checked = isReadOnly;
        //    btnFinilise.Enabled = !string.IsNullOrEmpty(docPath);

        //    this.RibbonUI.Invalidate();
        //}

        //private void btnFinalise_Click(object sender, RibbonControlEventArgs e)
        //{
        //    Globals.ThisAddIn.DisplayPane(PaneMode.LockDisplayPane);
        //}

        private void btnToggleSignature_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.DisplaySignaturePane();
        }

        
        //private void btnLock_Click(object sender, RibbonControlEventArgs e)
        //{
        //    Globals.ThisAddIn.DisplayLockPane();

        //}


    }

}
