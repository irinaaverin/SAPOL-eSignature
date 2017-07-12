using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using Microsoft.Ink;
using System.Drawing.Imaging;
using System.Configuration;
using NLog;

namespace SAPOLEStatement
{
    public partial class SignatureDisplayPane : UserControl
    {
        private const string INC_CTRL_NAME = "inkSignature";
        private const string LBL_CTRL_NAME = "lblSignature";
        private ImageSaver imageSaverHelper = new ImageSaver();
        private TemplateDetails documentEditProperties;
        
        public SignatureDisplayPane(TemplateDetails documentEditProperties, Logger logger)
            : this()
        {
            try
            {
                this.documentEditProperties = documentEditProperties;
                this.documentEditProperties.PDFFullName = string.Empty;
                
                UpdateControl();
            }
            catch (Exception ex)
            {

                logger.LogException(LogLevel.Error, "SignatureDisplayPane", ex);
            }
        }

        private void UpdateControl()
        {
            lblInfo.Text = string.Format("All signatures will be embedded into the word document, which will be converted to PDF format.{0}Signatures will not be saved in the word document.", Environment.NewLine);
            int signatures = 0;
            //Select only unique tags - from this list only 2 will be selected (SIGNATUREF1,SIGNATUREF2))
            //F2:SignatureF1, SIGNATUREF1, F2 | 
            //F2:SignatureF2, SIGNATUREF2, F2 | 
            //Deponent Footer, SIGNATUREF1, F1 | 
            //Witnessed By Footer, SIGNATUREF2, F1 | 

            List<SignatureInfo> uniqueList = documentEditProperties.SignatureInfoList.GroupBy(x => x.Tag).Select(y => y.Where(x=>!x.Title.Contains(":")).First()).ToList();

            this.pnlSignatures.SuspendLayout();
            Size signatureSize = GetSignatureBoxSize();
            Size signatureLabelSize = GetSignatureLabelsSize();
            int firstControlStep = 20;
             Int32.TryParse(ConfigurationManager.AppSettings["StepGap"].ToString(), out firstControlStep);
            foreach (SignatureInfo item in uniqueList)    
            {
                signatures++;
                item.Number = signatures;
                
                Label signatureLbl = new System.Windows.Forms.Label();
                signatureLbl.Name = string.Concat(LBL_CTRL_NAME, signatures);
                signatureLbl.Location = new Point(13, firstControlStep + 54 + (signatureSize.Height + 20) * (signatures - 1));
                signatureLbl.Size = signatureLabelSize;
                signatureLbl.Text = string.Concat(item.Title, ":");

                Microsoft.Ink.InkPicture inkSignature = new Microsoft.Ink.InkPicture();
                inkSignature.Name = string.Concat(INC_CTRL_NAME, signatures);
                inkSignature.Location = new Point(signatureLbl.Location.X + signatureLabelSize.Width + 20, firstControlStep + 54 + (signatureSize.Height + 20) * (signatures - 1));
                inkSignature.Size = signatureSize;
                inkSignature.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;

                this.pnlSignatures.Controls.Add(signatureLbl);
                this.pnlSignatures.Controls.Add(inkSignature);
            }
            this.pnlSignatures.ResumeLayout(false);
            this.pnlSignatures.PerformLayout();

            numSignatures.Value = signatures;
            pnlSignatures.Visible = (signatures > 0);
            btnDeleteImage.Visible = (signatures > 0);

            if (signatures == 0)
            {
                pnlButtons.Location = new Point(pnlButtons.Location.X, pnlSignatures.Location.Y);
            }

        }
        private int GetPanelWidth()
        {
            double widthPerc = 30;
            try
            {
                Double.TryParse(ConfigurationManager.AppSettings["SignaturePaneWidthRatioPercent"].ToString(), out widthPerc);

            }
            catch (Exception)
            {
                //
            }
            int width = Convert.ToInt32((Screen.PrimaryScreen.Bounds.Width / 100) * widthPerc);

            return width;
        }

        private int GetValueFromPercentage(string setting)
        {
            double value = 0;
            try
            {
                Double.TryParse(ConfigurationManager.AppSettings[setting].ToString(), out value);

            }
            catch (Exception)
            {
                //
            }
            int result = Convert.ToInt32((GetPanelWidth() / 100) * value);

            return result;
        }
        private Size GetSignatureLabelsSize()
        {
            int labelWidth = 82;            
            int height = 53;
            try
            {
                height = GetValueFromPercentage("SignatureBoxHeight");
                labelWidth = GetValueFromPercentage("LabelWidth");
                //Int32.TryParse(ConfigurationManager.AppSettings["SignatureBoxHeight"].ToString(), out height);
                //Int32.TryParse(ConfigurationManager.AppSettings["LabelWidth"].ToString(), out labelWidth);

            }
            catch (Exception)
            {
                //
            }
            return new Size(labelWidth == 0 ? 246 : labelWidth, height == 0 ? 53 : height);
        }
        private Size GetSignatureBoxSize()
        {
            int labelWidth = 82;
            int width = 246;
            int height = 53;
            try
            {
                labelWidth = GetValueFromPercentage("LabelWidth");
                width = GetValueFromPercentage("SignatureBoxWidth");
                height = GetValueFromPercentage("SignatureBoxHeight");


                //Int32.TryParse(ConfigurationManager.AppSettings["SignatureBoxWidth"].ToString(), out width);
                //Int32.TryParse(ConfigurationManager.AppSettings["SignatureBoxHeight"].ToString(), out height);
                //Int32.TryParse(ConfigurationManager.AppSettings["LabelWidth"].ToString(), out labelWidth);

            }
            catch (Exception)
            {
                //
            }
            return new Size(width == 0 ? 246 : width, height == 0 ? 53 : height);
        }
        public SignatureDisplayPane()
        {
            InitializeComponent();
            ResizeControl();
        }
        public void AdjustSize(int height)
        {
            this.pnlButtons.Dock = DockStyle.Bottom;
            this.pnlSignatures.Size = new Size(365, height - this.pnlButtons.Size.Height - 50);
        }
        private void ResizeControl()
        {
            foreach (Control ctrl in this.pnlSignatures.Controls)
            {
                int index = 1;
                if (ctrl.Name.Contains(INC_CTRL_NAME) || ctrl.Name.Contains(LBL_CTRL_NAME))
                {
                    index = Convert.ToInt16(ctrl.Name.Substring(ctrl.Name.Length - 1, 1));
                    ctrl.Visible = index <= numSignatures.Value;
                }

                if (ctrl is Microsoft.Ink.InkPicture)
                {
                    if (index == Convert.ToInt16(numSignatures.Value))
                    {
                        pnlSignatures.Size = new Size(pnlSignatures.Width, ctrl.Location.Y + ctrl.Size.Height + 5);
                        pnlButtons.Location = new Point(pnlButtons.Location.X, ctrl.Location.Y + ctrl.Size.Height + 20);
                    }
                }
            }
        }

        private void numSignatures_ValueChanged(object sender, EventArgs e)
        {
            ResizeControl();
        }

        private bool IsValid()
        {
            bool valid = true;            

            if (string.IsNullOrEmpty(this.documentEditProperties.PDFFullName))
            {
                MessageBox.Show("Please enter name of the PDF file", "File Name is Required");
                return false;
            }
            int nonSigned = 0;
            for (int i = 1; i <= numSignatures.Value; i++)
            {
                InkPicture inkPictureCtrl = this.Controls.Find(string.Concat(INC_CTRL_NAME, i), true).FirstOrDefault() as InkPicture;
                if (inkPictureCtrl.Ink.Strokes.Count == 0)
                    nonSigned++;
            }

            if (nonSigned > 0)
            {
                if (MessageBox.Show(string.Format("{0} of {1} signature boxes are not populated.{2} Do you want to continue?", nonSigned, numSignatures.Value, Environment.NewLine), "Missing signatures", MessageBoxButtons.OKCancel) == DialogResult.Cancel)
                    valid = false;
            }

            return valid;
        }

        private void btnSaveImage_Click(object sender, EventArgs e)
        {
            if (!IsValid())
                return;
            for (int i = 1; i <= numSignatures.Value; i++)
            {
                InkPicture inkPictureCtrl = this.Controls.Find(string.Concat(INC_CTRL_NAME, i), true).FirstOrDefault() as InkPicture;
                string fileToSaveJpeg = string.Empty;
                SignatureInfo item = documentEditProperties.SignatureInfoList.Where(x => x.Number == i).FirstOrDefault();

                if (!imageSaverHelper.SaveSignatureToArray(inkPictureCtrl, ref item))
                    break;

            }
            PopulateNonDisplayed();
            Globals.ThisAddIn.AddSignatures();
        }

        private void PopulateNonDisplayed()
        {
            List<SignatureInfo> unpopulated = documentEditProperties.SignatureInfoList.Where(x => string.IsNullOrEmpty(x.PicturePath)).ToList();

            foreach (SignatureInfo item in unpopulated)
            {
                SignatureInfo populated =
                documentEditProperties.SignatureInfoList.Where(x => x.Tag.Equals(item.Tag) && !string.IsNullOrEmpty(x.PicturePath)).FirstOrDefault();
                if (populated != null)
                {
                    item.PicturePath = populated.PicturePath;
                    item.PictureByteArray = populated.PictureByteArray;
                }
            }
        }

        private void btnDeleteImage_Click(object sender, EventArgs e)
        {
            for (int i = 1; i <= numSignatures.Value; i++)
            {
                InkPicture inkPictureCtrl = this.Controls.Find(string.Concat(INC_CTRL_NAME, i), true).FirstOrDefault() as InkPicture;
                imageSaverHelper.ClearInkPicture(inkPictureCtrl);
            }

            ClearSignatureInfo();
        }
        private void ClearSignatureInfo()
        {
            foreach (SignatureInfo item in documentEditProperties.SignatureInfoList)
            {
                item.PicturePath = string.Empty;
            }
        }

        private void btnPath_Click(object sender, EventArgs e)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.Filter = "PDF| *.pdf";
            dlg.Title = "Save document as PDF file";
            dlg.CheckPathExists = true;
            dlg.ValidateNames = true;
            if (!string.IsNullOrEmpty(this.documentEditProperties.DocumentPath))
                dlg.InitialDirectory = this.documentEditProperties.DocumentPath;
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                this.documentEditProperties.PDFFullName = dlg.FileName;
                this.documentEditProperties.DocumentPath = Path.GetDirectoryName(this.documentEditProperties.PDFFullName);

                txtPath.Text = this.documentEditProperties.PDFFullName;
            }

        }

        private void pnlSignatures_Resize(object sender, EventArgs e)
        {
            //AdjustSize();
        }

        private void btnLockImage_Click(object sender, EventArgs e)
        {

        }

    }

}
