using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Tools = Microsoft.Office.Tools;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Configuration;
using NLog;
using NLog.Config;
using NLog.Targets;
using System.Reflection;

namespace SAPOLEStatement
{
    public enum PaneMode
    {
        SignatureDisplayPane,
        LockDisplayPane,
        None
    }
    public partial class ThisAddIn
    {
        private Dictionary<string, TemplateDetails> documentEditProperties;
        private List<Tools.CustomTaskPane> SAPOLDisplayPanes;
        private PaneMode paneMode = PaneMode.None;
        private const string TEMPLATE_EXT = ".DOT";
        private Logger logger = LogManager.GetCurrentClassLogger();

        private void ThisAddIn_Startup (object sender, System.EventArgs e)
        {
            logger = LogHelper.InitialiseNLog();

            // Initialize timers and display panels
            documentEditProperties = new Dictionary<string, TemplateDetails>();
            SAPOLDisplayPanes = new List<Microsoft.Office.Tools.CustomTaskPane>();

            // Add event handlers
            Word.ApplicationEvents4_Event eventInterface = this.Application;
            eventInterface.DocumentOpen += new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentOpenEventHandler(eventInterface_DocumentOpen);
            eventInterface.NewDocument += new Microsoft.Office.Interop.Word.ApplicationEvents4_NewDocumentEventHandler(eventInterface_NewDocument);
            eventInterface.DocumentBeforeClose += new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentBeforeCloseEventHandler(eventInterface_DocumentBeforeClose);
            //eventInterface.WindowActivate += new Microsoft.Office.Interop.Word.ApplicationEvents4_WindowActivateEventHandler(eventInterface_WindowActivate);
            //eventInterface.DocumentBeforeSave += new ApplicationEvents4_DocumentBeforeSaveEventHandler(eventInterface_DocumentBeforeSave);

            // Start monitoring active document
            MonitorDocument(this.Application.ActiveDocument);
        }

        #region eventInterfaceEvents
        private void eventInterface_DocumentBeforeSave (Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            try
            {
                //this is done as a substitution for AfterSave event!
                Cancel = true;
                SaveAsUI = false;

                string oldName = Doc.Name;
                Doc.Saved = false;

                DocContent.UpdateDateModifiedCtrls(Doc);

                Doc.Save();

                ResetDocumentInDictionary(Doc, oldName);
            }
            catch (Exception e)
            {
                logger.LogException(LogLevel.Error, "eventInterface_DocumentBeforeSave", e);
            }
        }

        private void eventInterface_DocumentBeforeClose (Word.Document Doc, ref bool Cancel)
        {
            try
            {
                //if (Doc.Name.ToUpper().Contains(TEMPLATE_EXT))
                //    return;
                if (!Doc.Name.ToUpper().Contains(TEMPLATE_EXT) && documentEditProperties.ContainsKey(Doc.Name) && documentEditProperties[Doc.Name].IsSAPOL)
                    DocContent.UpdateDateModifiedCtrls(Doc);
                if (documentEditProperties.ContainsKey(Doc.Name))
                    ReleaseSignatureInfo(documentEditProperties[Doc.Name]);

                // Remove task pane
                RemoveTaskPaneFromWindow(Doc.ActiveWindow);
            }
            catch (Exception e)
            {
                logger.LogException(LogLevel.Error, "eventInterface_DocumentBeforeClose", e);
            }
        }

        private void eventInterface_NewDocument (Word.Document Doc)
        {
            // Monitor new doc
            MonitorDocument(Doc);
            //10/05/2017
            LockSignatures(Doc);
        }

        private void eventInterface_DocumentOpen (Word.Document Doc)
        {
            if (!documentEditProperties.ContainsKey(Doc.Name))
            {
                // Monitor new doc
                MonitorDocument(Doc);
            }
            if (!Doc.Name.ToUpper().Contains(TEMPLATE_EXT) && documentEditProperties[Doc.Name].IsSAPOL)
                Doc.TrackRevisions = true;
            //10/05/2017
            LockSignatures(Doc);
        }
        //10/05/2017
        private void LockSignatures (Word.Document doc)
        {
            List<Microsoft.Office.Interop.Word.ContentControl> signatures = DocContent.GetAllContentControls(doc);
            foreach (Microsoft.Office.Interop.Word.ContentControl item in signatures)
            {
                item.LockContentControl = true;
                item.LockContents = true;
            }
        }
        #endregion

        #region ReleaseSignatureInfo
        private void ReleaseSignatureInfo (TemplateDetails detailes)
        {
            foreach (SignatureInfo item in detailes.SignatureInfoList)
            {
                if (Marshal.IsComObject(item.SignatureControl))
                    Marshal.ReleaseComObject(item.SignatureControl);
            }
            detailes = null;
        }
        #endregion

        #region Manage documentEditProperties
        private void ResetDocumentInDictionary (Word.Document Doc, string oldName)
        {
            string newName = Doc.Name;
            if (documentEditProperties.ContainsKey(oldName))
            {
                TemplateDetails td = documentEditProperties[oldName];
                documentEditProperties.Remove(oldName);
                documentEditProperties.Add(newName, td);
            }
        }


        private void MonitorDocument (Word.Document Doc)
        {
            // Monitor doc
            if (!documentEditProperties.ContainsKey(Doc.Name))
            {
                bool isSAPOL = DocContent.IsSAPOL(Doc);
                documentEditProperties.Add(Doc.Name, new TemplateDetails(isSAPOL));
                //Globals.Ribbons.SignatureRibbon.SetToggleSignatureLock(isSAPOL);
            }
        }
        #endregion

        #region RemoveTaskPane
        private Tools.CustomTaskPane GetPaneByName ()
        {
            Tools.CustomTaskPane docPane = null;
            foreach (Tools.CustomTaskPane pane in SAPOLDisplayPanes)
            {
                if (pane.Control.Name == paneMode.ToString())
                {
                    docPane = pane;
                    break;
                }
            }
            return docPane;
        }

        private Tools.CustomTaskPane GetPaneByWindow (Word.Window Wn)
        {
            Tools.CustomTaskPane docPane = null;
            foreach (Tools.CustomTaskPane pane in SAPOLDisplayPanes)
            {
                if (pane.Window == Wn)
                {
                    docPane = pane;
                    break;
                }
            }
            return docPane;
        }

        private void RemoveTaskPaneFromWindow ()
        {
            Tools.CustomTaskPane docPane = GetPaneByName();
            // Remove document task pane
            if (docPane != null)
            {
                this.CustomTaskPanes.Remove(docPane);
                SAPOLDisplayPanes.Remove(docPane);
            }
        }

        private void RemoveTaskPaneFromWindow (Word.Window Wn)
        {
            // Check for task pane in window
            Tools.CustomTaskPane docPane = GetPaneByWindow(Wn);

            // Remove document task pane
            if (docPane != null)
            {
                this.CustomTaskPanes.Remove(docPane);
                SAPOLDisplayPanes.Remove(docPane);
            }
        }
        #endregion

        #region DisplaySignaturePane
        private int GetWidth ()
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

        internal void DisplaySignaturePane ()
        {
            paneMode = PaneMode.SignatureDisplayPane;
            Word.Window window = this.Application.ActiveDocument.ActiveWindow;

            RemoveTaskPaneFromWindow(window);

            // Check for task pane in window
            Tools.CustomTaskPane docPane = null;
            Tools.CustomTaskPane paneToRemove = null;
            foreach (Tools.CustomTaskPane pane in SAPOLDisplayPanes)
            {
                try
                {
                    if (pane.Window == window)
                    {
                        docPane = pane;
                        break;
                    }
                }
                catch (ArgumentNullException)
                {
                    // pane.Window is null, so document1 has been unloaded.
                    paneToRemove = pane;
                }
            }

            // Remove pane if necessary
            SAPOLDisplayPanes.Remove(paneToRemove);

            // Add task pane to doc
            if (docPane == null)
            {
                Tools.CustomTaskPane pane;
                //09/05/2017
                try
                {
                    this.Application.ActiveDocument.Unprotect("NEOPHYTE");
                }
                catch (Exception)
                {
                    //
                }
                //PopulateDocumentPath();
                PopulateSignatureFileds();
                SignatureDisplayPane userCtrl = new SignatureDisplayPane(documentEditProperties[this.Application.ActiveDocument.Name], logger);
                pane = this.CustomTaskPanes.Add(
                       userCtrl,
                       "Add Signature(s) to Document",
                       window);

                pane.Width = GetWidth();
                userCtrl.AdjustSize(pane.Height);
                SAPOLDisplayPanes.Add(pane);
                pane.Visible = true;
            }
        }

        private void PopulateDocumentPath ()
        {
            string docPath = this.Application.ActiveDocument.Path;
            if (string.IsNullOrEmpty(docPath))
            {
                Word.Template template = (Word.Template)this.Application.ActiveDocument.get_AttachedTemplate();
                docPath = template.Path;
            }

            documentEditProperties[this.Application.ActiveDocument.Name].DocumentPath = docPath;
            //do not prepopulate file name - let user do this
            //documentEditProperties[this.Application.ActiveDocument.Name].PDFFullName = Path.Combine(
            //    docPath,string.Concat(Path.GetFileNameWithoutExtension(this.Application.ActiveDocument.Name), ".pdf"));

        }

        private void PopulateSignatureFileds ()
        {
            Word.Document doc = this.Application.ActiveDocument;
            documentEditProperties[this.Application.ActiveDocument.Name].SignatureInfoList = DocContent.BuildSignatureList(doc);
        }
        #endregion

        #region Add Signatures to PDF
        private void SaveAsPDF ()
        {
            try
            {
                Word.Document doc = this.Application.ActiveDocument;

                DocContent.UpdateDateModifiedCtrls(doc);
                DocContent.SaveDocumentAsPDF(doc,
                    documentEditProperties[doc.Name].PDFFullName);
            }
            catch (System.IO.IOException ex)
            {
                MessageBox.Show(string.Format("The following error occurs while saving the document into PDF format: {0}", ex.Message), "Error on saving as PDF");
            }
            catch (Exception ex)
            {
                logger.LogException(LogLevel.Error, "SaveAsPDF", ex);
            }
        }
        //09/05/2017 - SaveDocument before adding signatures
        private void SaveAsWordDocumentWithoutSignatures (string wordVersion)
        {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            if (File.Exists(wordVersion))
                File.Delete(wordVersion);
            //11/05/2017 - protect before saving, but then unprotect after saving 
            //as the process of saving Signature images requires the unptorected document
            object noReset = false;
            object password = "NEOPHYTE";
            object useIRM = false;
            object enforceStyleLock = false;

            doc.Protect(Word.WdProtectionType.wdAllowOnlyReading, ref noReset,
                ref password, ref useIRM, ref enforceStyleLock);

            var persistFile = (System.Runtime.InteropServices.ComTypes.IPersistFile)doc;

            persistFile.Save(wordVersion, false);

            //11/05/2017 
            doc.Unprotect(password);
        }
        //10/05/2017
        private void PasswordProtectDocument (string fileName)
        {
            Word.Document doc = null;
            object missing = System.Reflection.Missing.Value;
            object readOnly = false;
            object visible = true;
            object fileToOpen = fileName;
            try
            {
                doc = Globals.ThisAddIn.Application.Documents.Open(ref fileToOpen, ref missing, ref readOnly, ref missing, ref missing,
                                                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                ref visible, ref visible, ref missing, ref missing, ref missing);

                try
                {
                    doc.Activate();

                    object noReset = false;
                    object password = "NEOPHYTE";
                    object useIRM = false;
                    object enforceStyleLock = false;

                    doc.Protect(Word.WdProtectionType.wdAllowOnlyReading, ref noReset,
                        ref password, ref useIRM, ref enforceStyleLock);

                    doc.SaveAs(ref fileToOpen, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                      ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                }
                catch (Exception)
                {
                    //
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                doc.Close(ref missing, ref missing, ref missing);

            }
        }
        internal void AddSignatures ()
        {

            TemplateDetails details = documentEditProperties[this.Application.ActiveDocument.Name];
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            //Save document using SaveAs function
            SaveAsWordDocumentWithoutSignatures();

            //12/03/2017 - unlock controls to insert signatures
            DocContent.UnlockHeaderSignatures(doc);
            DocContent.UnlockBodySignatures(doc);
            DocContent.UnlockFooterSignatures(doc);

            int selectionStart = 0;

            //Add Signatures to document
            foreach (SignatureInfo item in details.SignatureInfoList)
            {
                try
                {
                    DocContent.ByteArrayToImageFilebyMemoryStream(item);
                    if (!string.IsNullOrEmpty(item.PicturePath))
                    {
                        DocContent.AddPictureToControl(doc, item.SignatureControl, item.PicturePath, ref selectionStart);

                        if (File.Exists(item.PicturePath))
                            File.Delete(item.PicturePath);

                    }
                }
                catch (Exception ex)
                {
                    logger.LogException(LogLevel.Error, "AddSignatures1", ex);
                }

            }
            //Save as pdf            
            SaveAsPDF();

            try
            {
                ////09/05/2017                
                ////Remove pictures
                //foreach (SignatureInfo item in details.SignatureInfoList)
                //{
                ////  09/05/2017 - ira
                //    item.SignatureControl.Delete();
                //    DocContent.DeletePictureFromControl(doc, item.SignatureControl);
                //}
                ////logger.Log(LogLevel.Warn, "B4 SaveAsWordDocumentWithoutSignatures");
                ////Save as Word Document without the signature
                //SaveAsWordDocumentWithoutSignatures();
                ////logger.Log(LogLevel.Warn, "After SaveAsWordDocumentWithoutSignatures");

                RemoveTaskPaneFromWindow();
            }
            catch (Exception ex)
            {
                logger.LogException(LogLevel.Error, "AddSignatures2", ex);
            }

            object saveOption = Word.WdSaveOptions.wdDoNotSaveChanges;
            object originalFormat = Word.WdOriginalFormat.wdOriginalDocumentFormat;
            object routeDocument = false;
            doc.Close(ref saveOption, ref originalFormat, ref routeDocument);

        }
        
        private void SaveAsWordDocumentWithoutSignatures ()
        {
            try
            {
                Word.Document doc = this.Application.ActiveDocument;

                string oldName = doc.Name;
                DocContent.SaveDocumentAsWordWithoutSignatures(doc,
                    documentEditProperties[doc.Name].PDFFullName);

                ResetDocumentInDictionary(doc, oldName);
            }
            catch (System.IO.IOException ex)
            {
                MessageBox.Show(string.Format("The following error occurs while saving the document into word document format: {0}", ex.Message), "Error on saving as Word Document Without the Signatures");
            }
            catch (Exception ex)
            {
                logger.LogException(LogLevel.Error, "SaveAsWordDocumentWithoutSignatures", ex);
            }
        }
        #endregion

        #region FinaliseDocument
        internal void LockDocument ()
        {

            try
            {
                TemplateDetails details = documentEditProperties[this.Application.ActiveDocument.Name];

                object noReset = false;
                object password = details.Password;
                object useIRM = false;
                object enforceStyleLock = false;

                this.Application.ActiveDocument.Protect(Word.WdProtectionType.wdAllowOnlyReading, ref noReset,
                    ref password, ref useIRM, ref enforceStyleLock);

                RemoveTaskPaneFromWindow();
            }
            catch (Exception ex)
            {
                logger.LogException(LogLevel.Error, "LockDocument", ex);
            }
        }
        #endregion

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup ()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        private void ThisAddIn_Shutdown (object sender, System.EventArgs e)
        {
            try
            {
                foreach (KeyValuePair<string, TemplateDetails> item in documentEditProperties)
                {
                    ReleaseSignatureInfo(item.Value);
                }
            }
            catch (Exception ex)
            {
                logger.LogException(LogLevel.Error, "ThisAddIn_Shutdown", ex);
            }
        }
        #endregion
    }
}

