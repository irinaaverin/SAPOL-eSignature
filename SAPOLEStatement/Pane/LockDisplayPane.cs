using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SAPOLEStatement
{
    public partial class LockDisplayPane : UserControl
    {
        private TemplateDetails documentEditSignatures = null;

        public LockDisplayPane(TemplateDetails documentEditSignatures)
            : this()
        {
            // Store reference to edit times
            this.documentEditSignatures = documentEditSignatures;
        }

        public LockDisplayPane()
        {
            InitializeComponent();
            if (this.documentEditSignatures == null)
                this.documentEditSignatures = new TemplateDetails();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtPassword.Text))
            {
                MessageBox.Show("The password cannot be blank", "SA Police Add-Ins");
                return;
            }
            else if (!txtPassword.Text.Equals(txtPasswordRepeat.Text))
            {
                MessageBox.Show("The password confirmation does not match", "SA Police Add-Ins");
                return;
            }
            documentEditSignatures.Password = txtPassword.Text;
            Globals.ThisAddIn.LockDocument();
        }

    }

}
