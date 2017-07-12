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
using System.Runtime.InteropServices;
using System.Configuration;
using NLog;
using System.Reflection;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

namespace SAPOLEStatement
{
    internal static class DocContent
    {
        static readonly string SIGNATURE_FIELD;
        static readonly string DATEMODIFIED_CONTROL;
        static readonly string SAPOLMARKER_CONTROL;
        private static Logger logger = LogManager.GetCurrentClassLogger();

        static DocContent ()
        {
            logger = LogHelper.InitialiseNLog();
            try
            {
                SIGNATURE_FIELD = ConfigurationManager.AppSettings["SignatureFieldName"].ToString().ToUpper();
                DATEMODIFIED_CONTROL = ConfigurationManager.AppSettings["DateModifiedFieldName"].ToString().ToUpper();
                SAPOLMARKER_CONTROL = ConfigurationManager.AppSettings["SAPOLFieldName"].ToString().ToUpper();
            }
            catch (Exception ex)
            {
                logger.LogException(LogLevel.Error, "DocContent", ex);
                SIGNATURE_FIELD = "Signature".ToUpper();
                DATEMODIFIED_CONTROL = "DateModified".ToUpper();
                SAPOLMARKER_CONTROL = "SAPOLMarker".ToUpper();
            }
        }

        internal static bool IsSAPOL (Word.Document doc)
        {
            return GetContentControl(doc, SAPOLMARKER_CONTROL) != null;
        }

        internal static Word.ContentControl GetContentControl (Word.Document doc, string name)
        {
            Word.ContentControl contentControl = GetBodyContentControl(doc, name);
            if (contentControl == null)
                contentControl = GetFooterContentControl(doc, name);
            if (contentControl == null)
                contentControl = GetHeaderContentControl(doc, name);
            return contentControl;
        }

        internal static List<SignatureInfo> BuildSignatureList (Word.Document doc)
        {
            //09/05/2017 - better version of collection of signature controls, but it's too late now to
            //make any changes
            //List<Microsoft.Office.Interop.Word.ContentControl> signatures = GetAllContentControls(doc);

            List<SignatureInfo> signaturesCtrls = new List<SignatureInfo>();
            try
            {
                //AnalyseDoc(doc);

                //Tag will include section and name in the format:
                //B:Signature1
                //F1:Signature2
                //F2:Signature2
                //header
                List<SignatureInfo> headerCtrls = GetHeaderSignatureCtrls(doc);
                signaturesCtrls.AddRange(headerCtrls.GroupBy(x => new { x.Tag, x.SignatureControlPosition }).Select(y => y.First()).ToList());
                //signaturesCtrls.AddRange(GetHeaderSignatureCtrls(doc));

                //body
                List<SignatureInfo> bodyCtrls = GetBodySignatureCtrls(doc);
                signaturesCtrls.AddRange(bodyCtrls.GroupBy(x => new { x.Tag, x.SignatureControlPosition }).Select(y => y.First()).ToList());

                //footer
                List<SignatureInfo> footerCtrls = GetFooterSignatureCtrls(doc);
                signaturesCtrls.AddRange(footerCtrls.GroupBy(x => new { x.Tag, x.SignatureControlPosition }).Select(y => y.First()).ToList());
                //signaturesCtrls.AddRange(GetFooterSignatureCtrls(doc));
                foreach (SignatureInfo item in signaturesCtrls)
                {
                    //10/05/2017
                    item.SignatureControl.LockContentControl = true;
                    item.SignatureControl.LockContents = true;
                    logger.Log(LogLevel.Info, string.Format("{0}, {1}, {2}", item.Title, item.Tag, item.SignatureControlPosition));
                }
            }
            catch (Exception ex)
            {
                logger.LogException(LogLevel.Error, "BuildSignatureList", ex);
            }
            return signaturesCtrls;
        }

        internal static void AddPictureToControl (Word.Document doc, Word.ContentControl contentControl, string fileName, ref int selectionStart)
        {
            try
            {
                if (contentControl != null)
                {
                    contentControl.SetPlaceholderText(null, null, String.Empty);
                    //selectionStart = contentControl.Range.End + 1;
                    //int start = contentControl.Range.Start;
                    //int end = contentControl.Range.End;

                    Object oMissed = contentControl.Range; //the position you want to insert
                    Object oLinkToFile = false;  //default
                    Object oSaveWithDocument = true;//default

                    //08/05/2017 - Remove all typed text
                    try
                    {
                        contentControl.LockContents = false;
                        contentControl.LockContentControl = false;
                        ((contentControl).Range).Text = "";
                    }
                    catch (Exception ex)
                    {
                        string A = ex.Message;
                        //
                    }

                    var shape = doc.InlineShapes.AddPicture(fileName, ref  oLinkToFile, ref  oSaveWithDocument, ref  oMissed);
                    //shape.Width = 300;
                    //shape.Height = 25;

                    if (Marshal.IsComObject(oMissed))
                        Marshal.ReleaseComObject(oMissed);
                }
            }
            catch (Exception ex)
            {
                logger.LogException(LogLevel.Error, "AddPictureToControl", ex);
            }
        }
        //09/05/2017
        //internal static void DeletePictureFromControl (Word.Document doc, Word.ContentControl contentControl)
        //{
        //    try
        //    {
        //        if (contentControl != null && contentControl.Tag != null)
        //        {
        //            //09/05/2017 - ira
        //            contentControl.Delete();

        //            //contentControl.SetPlaceholderText(null, null, String.Empty);
        //            //Range range = contentControl.Range;
        //            //foreach (InlineShape oIShape in range.InlineShapes)
        //            //{
        //            //    oIShape.Delete();
        //            //    if (Marshal.IsComObject(oIShape))
        //            //        Marshal.ReleaseComObject(oIShape);
        //            //}
        //            //if (Marshal.IsComObject(range))
        //            //    Marshal.ReleaseComObject(range);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        logger.LogException(LogLevel.Error, "DeletePictureFromControl", ex);
        //    }

        //}
        //09/05/2017
        //internal static void DeletePictureFromControl (Word.Document doc, string tag)
        //{
        //    Word.ContentControl contentControl = DocContent.GetContentControl(doc, tag);
        //    DeletePictureFromControl(doc, contentControl);
        //}

        internal static Word.Bookmark GetBookmark (Word.Document Doc, string name)
        {
            Word.Bookmark bookmark = null;
            try
            {
                if (Doc.Bookmarks.Exists(name))
                {
                    object firstHalfName = name;
                    bookmark = Doc.Bookmarks.get_Item(ref firstHalfName);
                }
            }
            catch (Exception ex)
            {
                logger.LogException(LogLevel.Error, "GetBookmark", ex);
            }
            return bookmark;
        }

        internal static void UpdateDateModifiedCtrls (Word.Document doc)
        {
            List<Word.ContentControl> dateModifiedControl = new List<Word.ContentControl>();
            dateModifiedControl.AddRange(GetBodyCtrls(doc, DATEMODIFIED_CONTROL));
            dateModifiedControl.AddRange(GetHeaderCtrls(doc, DATEMODIFIED_CONTROL));
            dateModifiedControl.AddRange(GetFooterCtrls(doc, DATEMODIFIED_CONTROL));
            string dateModified = FormatDate(DateTime.Now);
            try
            {
                foreach (Word.ContentControl control in dateModifiedControl)
                {
                    control.Range.Text = dateModified;

                    if (Marshal.IsComObject(control))
                        Marshal.ReleaseComObject(control);
                }
            }
            catch (Exception ex)
            {
                logger.LogException(LogLevel.Error, "UpdateDateModifiedCtrls", ex);
            }
        }

        private static string FormatDate (DateTime date)
        {
            return string.Format("{0} {1}, {2} {3}", date.ToString("MMMM"), date.Day.AsOrdinal(), date.ToString("yyyy"), date.ToString("HH:mm"));
        }

        private static List<Word.ContentControl> GetFooterCtrls (Word.Document doc, string tag, bool contains = true)
        {
            List<Microsoft.Office.Interop.Word.ContentControl> footerCtrls = new List<Microsoft.Office.Interop.Word.ContentControl>();
            try
            {
                foreach (Microsoft.Office.Interop.Word.Section wordSection in doc.Sections)
                {
                    foreach (Microsoft.Office.Interop.Word.HeaderFooter wordFooter in wordSection.Footers)
                    {
                        Microsoft.Office.Interop.Word.Range docRange = wordFooter.Range;

                        ContentControls footerControls = docRange.ContentControls;

                        if (footerControls != null)
                        {
                            foreach (Microsoft.Office.Interop.Word.ContentControl control in footerControls)
                            {
                                if (contains)
                                {
                                    if (control.Tag.ToUpper().Contains(tag.ToUpper()))
                                    {
                                        footerCtrls.Add(control);
                                    }
                                }
                                else
                                {
                                    if (control.Tag.ToUpper().Equals(tag.ToUpper()))
                                    {
                                        footerCtrls.Add(control);
                                    }
                                }
                            }
                        }

                        if (Marshal.IsComObject(docRange))
                            Marshal.ReleaseComObject(docRange);
                        if (Marshal.IsComObject(footerControls))
                            Marshal.ReleaseComObject(footerControls);
                        if (Marshal.IsComObject(wordFooter))
                            Marshal.ReleaseComObject(wordFooter);
                    }

                    if (Marshal.IsComObject(wordSection))
                        Marshal.ReleaseComObject(wordSection);
                }
            }
            catch (Exception ex)
            {
                logger.LogException(LogLevel.Error, "GetFooterCtrls", ex);
            }
            return footerCtrls;
        }

        private static List<Word.ContentControl> GetHeaderCtrls (Word.Document doc, string tag, bool contains = true)
        {
            List<Microsoft.Office.Interop.Word.ContentControl> headerCtrls = new List<Microsoft.Office.Interop.Word.ContentControl>();
            try
            {
                foreach (Microsoft.Office.Interop.Word.Section wordSection in doc.Sections)
                {
                    foreach (Microsoft.Office.Interop.Word.HeaderFooter wordHeader in wordSection.Headers)
                    {
                        Microsoft.Office.Interop.Word.Range docRange = wordHeader.Range;

                        ContentControls headerControls = docRange.ContentControls;

                        if (headerControls != null)
                        {
                            foreach (Microsoft.Office.Interop.Word.ContentControl control in headerControls)
                            {
                                if (contains)
                                {
                                    if (control.Tag.ToUpper().Contains(tag.ToUpper()))
                                    {
                                        headerCtrls.Add(control);
                                    }
                                }
                                else
                                {
                                    if (control.Tag.ToUpper().Equals(tag.ToUpper()))
                                    {
                                        headerCtrls.Add(control);
                                    }
                                }
                            }
                        }

                        if (Marshal.IsComObject(docRange))
                            Marshal.ReleaseComObject(docRange);
                        if (Marshal.IsComObject(headerControls))
                            Marshal.ReleaseComObject(headerControls);
                        if (Marshal.IsComObject(wordHeader))
                            Marshal.ReleaseComObject(wordHeader);
                    }

                    if (Marshal.IsComObject(wordSection))
                        Marshal.ReleaseComObject(wordSection);
                }
            }
            catch (Exception ex)
            {
                logger.LogException(LogLevel.Error, "GetHeaderCtrls", ex);
            }
            return headerCtrls;
        }

        private static List<SignatureInfo> GetFooterSignatureCtrls (Word.Document doc)
        {
            List<SignatureInfo> signaturesFooter = new List<SignatureInfo>();
            try
            {
                foreach (Microsoft.Office.Interop.Word.Section wordSection in doc.Sections)
                {
                    foreach (Microsoft.Office.Interop.Word.HeaderFooter wordFooter in wordSection.Footers)
                    {
                        Microsoft.Office.Interop.Word.Range docRange = wordFooter.Range;

                        ContentControls footerControls = docRange.ContentControls;

                        if (footerControls != null)
                        {
                            int counter = 1;
                            foreach (Microsoft.Office.Interop.Word.ContentControl control in footerControls)
                            {
                                //MessageBox.Show("footerControls:" + control.Tag);
                                if (control.Tag.ToUpper().Contains(SIGNATURE_FIELD))
                                {

                                    //12/03/2017 - Do not allow typing in this field from inside the document
                                    //control.LockContents = false;
                                    //control.LockContentControl = false;
                                    Microsoft.Office.Interop.Word.Range ctrlRange = control.Range;
                                    int start = ctrlRange.Start;
                                    int end = ctrlRange.End;
                                    //Tag will include section and name in the format:
                                    //B:Signature1
                                    //F1:Signature2
                                    //F2:Signature2
                                    string[] parts = control.Tag.ToUpper().Split(':');
                                    SignatureInfo signatureInfo = new SignatureInfo
                                    {
                                        Title = control.Title,
                                        Tag = parts.Length > 1 ? parts[1] : control.Tag,
                                        SignatureControl = control,
                                        SignatureControlPosition = parts.Length > 1 ? parts[0] : "Footer"

                                    };
                                    counter++;
                                    signaturesFooter.Add(signatureInfo);
                                }
                            }
                        }

                        if (Marshal.IsComObject(docRange))
                            Marshal.ReleaseComObject(docRange);
                        if (Marshal.IsComObject(footerControls))
                            Marshal.ReleaseComObject(footerControls);

                        if (Marshal.IsComObject(wordFooter))
                            Marshal.ReleaseComObject(wordFooter);

                    }

                    if (Marshal.IsComObject(wordSection))
                        Marshal.ReleaseComObject(wordSection);
                }
            }
            catch (Exception ex)
            {
                logger.LogException(LogLevel.Error, "GetFooterSignatureCtrls", ex);
            }
            return signaturesFooter;
        }
        private static void AnalyseDoc (Word.Document doc)
        {
            int sections = 0;
            foreach (Microsoft.Office.Interop.Word.Section wordSection in doc.Sections)
            {
                sections++;
                int header = 0;
                int footer = 0;
                int controlsH = 0;
                int controlsF = 0;
                foreach (Microsoft.Office.Interop.Word.HeaderFooter fh in wordSection.Headers)
                {
                    Microsoft.Office.Interop.Word.Range docRange = fh.Range;
                    logger.Log(LogLevel.Info, "Header: " + docRange.Text);
                    controlsH += docRange.ContentControls.Count;
                    header++;
                }
                foreach (Microsoft.Office.Interop.Word.HeaderFooter fh in wordSection.Footers)
                {
                    Microsoft.Office.Interop.Word.Range docRange = fh.Range;
                    logger.Log(LogLevel.Info, "Footer: " + docRange.Text);
                    controlsF += docRange.ContentControls.Count;
                    footer++;
                }
                logger.Log(LogLevel.Info, string.Format("body controls:", doc.ContentControls.Count));
                logger.Log(LogLevel.Info, string.Format("section:{0}, headers:{1} - controls:{3}, footers:{2} - controls:{4} ", sections, header, footer, controlsH, controlsF));
            }
        }
        private static List<SignatureInfo> GetHeaderSignatureCtrls (Word.Document doc)
        {

            List<SignatureInfo> signaturesHeader = new List<SignatureInfo>();
            try
            {
                foreach (Microsoft.Office.Interop.Word.Section wordSection in doc.Sections)
                {
                    foreach (Microsoft.Office.Interop.Word.HeaderFooter wordHeader in wordSection.Headers)
                    {
                        Microsoft.Office.Interop.Word.Range docRange = wordHeader.Range;
                        ContentControls headerControls = docRange.ContentControls;

                        if (headerControls != null)
                        {
                            int counter = 1;
                            foreach (Microsoft.Office.Interop.Word.ContentControl control in headerControls)
                            {
                                //MessageBox.Show("headerControls:" + control.Tag);
                                if (control.Tag.ToUpper().Contains(SIGNATURE_FIELD))
                                {
                                    //12/03/2017 - Do not allow typing in this field from inside the document
                                    //control.LockContents = false;
                                    //control.LockContentControl = false;

                                    //Tag will include section and name in the format:
                                    //B:Signature1
                                    //F1:Signature2
                                    //F2:Signature2
                                    string[] parts = control.Tag.ToUpper().Split(':');
                                    SignatureInfo signatureInfo = new SignatureInfo
                                    {
                                        Title = control.Title,
                                        Tag = parts.Length > 1 ? parts[1] : control.Tag,
                                        SignatureControl = control,
                                        SignatureControlPosition = parts.Length > 1 ? parts[0] : "Header"
                                    };
                                    counter++;
                                    signaturesHeader.Add(signatureInfo);
                                }
                            }
                        }

                        if (Marshal.IsComObject(docRange))
                            Marshal.ReleaseComObject(docRange);
                        if (Marshal.IsComObject(headerControls))
                            Marshal.ReleaseComObject(headerControls);

                        if (Marshal.IsComObject(wordHeader))
                            Marshal.ReleaseComObject(wordHeader);

                    }

                    if (Marshal.IsComObject(wordSection))
                        Marshal.ReleaseComObject(wordSection);
                }
            }
            catch (Exception ex)
            {
                logger.LogException(LogLevel.Error, "GetHeaderSignatureCtrls", ex);
            }
            return signaturesHeader;
        }
        //12/03/2017 - unlock controls to insert signatures
        public static void UnlockFooterSignatures (Word.Document doc)
        {
            try
            {
                foreach (Microsoft.Office.Interop.Word.Section wordSection in doc.Sections)
                {
                    foreach (Microsoft.Office.Interop.Word.HeaderFooter wordFooter in wordSection.Footers)
                    {
                        Microsoft.Office.Interop.Word.Range docRange = wordFooter.Range;

                        ContentControls footerControls = docRange.ContentControls;

                        if (footerControls != null)
                        {
                            foreach (Microsoft.Office.Interop.Word.ContentControl control in footerControls)
                            {
                                if (control.Tag.ToUpper().Contains(SIGNATURE_FIELD))
                                {
                                    control.LockContents = false;
                                    control.LockContentControl = false;

                                }
                            }
                        }

                        if (Marshal.IsComObject(docRange))
                            Marshal.ReleaseComObject(docRange);
                        if (Marshal.IsComObject(footerControls))
                            Marshal.ReleaseComObject(footerControls);

                        if (Marshal.IsComObject(wordFooter))
                            Marshal.ReleaseComObject(wordFooter);

                    }

                    if (Marshal.IsComObject(wordSection))
                        Marshal.ReleaseComObject(wordSection);
                }
            }
            catch (Exception ex)
            {
                logger.LogException(LogLevel.Error, "GetFooterSignatureCtrls", ex);
            }
        }
        //12/03/2017 - unlock controls to insert signatures
        public static void UnlockHeaderSignatures (Word.Document doc)
        {
            try
            {
                foreach (Microsoft.Office.Interop.Word.Section wordSection in doc.Sections)
                {
                    foreach (Microsoft.Office.Interop.Word.HeaderFooter wordHeader in wordSection.Headers)
                    {
                        Microsoft.Office.Interop.Word.Range docRange = wordHeader.Range;
                        ContentControls headerControls = docRange.ContentControls;

                        if (headerControls != null)
                        {

                            foreach (Microsoft.Office.Interop.Word.ContentControl control in headerControls)
                            {
                                if (control.Tag.ToUpper().Contains(SIGNATURE_FIELD))
                                {
                                    control.LockContents = false;
                                    control.LockContentControl = false;
                                }
                            }
                        }

                        if (Marshal.IsComObject(docRange))
                            Marshal.ReleaseComObject(docRange);
                        if (Marshal.IsComObject(headerControls))
                            Marshal.ReleaseComObject(headerControls);

                        if (Marshal.IsComObject(wordHeader))
                            Marshal.ReleaseComObject(wordHeader);

                    }

                    if (Marshal.IsComObject(wordSection))
                        Marshal.ReleaseComObject(wordSection);
                }
            }
            catch (Exception ex)
            {
                logger.LogException(LogLevel.Error, "UnlockHeaderSignatures", ex);
            }
        }
        //09/05/2017 
        public static List<Microsoft.Office.Interop.Word.ContentControl> GetAllContentControls (Microsoft.Office.Interop.Word.Document wordDocument)
        {
            if (null == wordDocument)
                throw new ArgumentNullException("wordDocument");

            List<Microsoft.Office.Interop.Word.ContentControl> ccList = new List<Microsoft.Office.Interop.Word.ContentControl>();

            // The code below search content controls in all
            // word document stories see http://word.mvps.org/faqs/customization/ReplaceAnywhere.htm
            Range rangeStory;
            foreach (Range range in wordDocument.StoryRanges)
            {
                rangeStory = range;
                do
                {
                    try
                    {
                        foreach (Microsoft.Office.Interop.Word.ContentControl cc in rangeStory.ContentControls)
                        {
                            if (cc.Tag.ToUpper().Contains(SIGNATURE_FIELD))
                            {
                                logger.Log(LogLevel.Info, "List: " + cc.Tag.ToUpper());
                                ccList.Add(cc);
                            }
                        }
                        foreach (Shape shapeRange in rangeStory.ShapeRange)
                        {
                            foreach (Microsoft.Office.Interop.Word.ContentControl cc in shapeRange.TextFrame.TextRange.ContentControls)
                            {
                                if (cc.Tag.ToUpper().Contains(SIGNATURE_FIELD))
                                {
                                    logger.Log(LogLevel.Info, "List: " + cc.Tag.ToUpper());
                                    ccList.Add(cc);
                                }
                            }
                        }
                    }
                    catch (COMException) { }
                    rangeStory = rangeStory.NextStoryRange;

                }
                while (rangeStory != null);
            }
            return ccList;
        }
        //12/03/2017 - unlock controls to insert signatures
        public static void UnlockBodySignatures (Word.Document doc)
        {
            //Body
            try
            {
                foreach (Microsoft.Office.Interop.Word.ContentControl control in doc.ContentControls)
                {
                    if (control.Tag.ToUpper().Contains(SIGNATURE_FIELD))
                    {
                        control.LockContents = false;
                        control.LockContentControl = false;
                    }
                }
            }
            catch (Exception ex)
            {
                logger.LogException(LogLevel.Error, "UnlockBodySignatures", ex);
            }

        }
        private static List<SignatureInfo> GetBodySignatureCtrls (Word.Document doc)
        {
            List<SignatureInfo> signaturesBody = new List<SignatureInfo>();

            int counter = 1;
            try
            {
                foreach (Microsoft.Office.Interop.Word.ContentControl control in doc.ContentControls)
                {
                    //MessageBox.Show("body:"+control.Tag);
                    if (control.Tag.ToUpper().Contains(SIGNATURE_FIELD))
                    {

                        //12/03/2017 - Do not allow typing in this field from inside the document
                        //control.LockContents = false;
                        //control.LockContentControl = false;
                        //Tag will include section and name in the format:
                        //B:Signature1
                        //F1:Signature2
                        //F2:Signature2
                        string[] parts = control.Tag.ToUpper().Split(':');
                        SignatureInfo signatureInfo = new SignatureInfo
                        {
                            Title = control.Title,
                            Tag = parts.Length > 1 ? parts[1] : control.Tag,
                            SignatureControl = control,
                            SignatureControlPosition = parts.Length > 1 ? parts[0] : "Body"

                        };
                        counter++;
                        signaturesBody.Add(signatureInfo);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.LogException(LogLevel.Error, "GetBodySignatureCtrls", ex);
            }
            return signaturesBody;
        }

        private static Word.ContentControl GetHeaderContentControl (Word.Document doc, string name)
        {
            Word.ContentControl contentControl = null;
            try
            {
                foreach (Microsoft.Office.Interop.Word.Section wordSection in doc.Sections)
                {
                    foreach (Microsoft.Office.Interop.Word.HeaderFooter wordHeader in wordSection.Headers)
                    {
                        Microsoft.Office.Interop.Word.Range docRange = wordHeader.Range;

                        ContentControls headerControls = docRange.ContentControls;

                        if (headerControls != null)
                        {
                            foreach (Microsoft.Office.Interop.Word.ContentControl control in headerControls)
                            {
                                if (control.Tag.ToUpper().Equals(name.ToUpper()))
                                {
                                    contentControl = control;
                                    break;
                                }
                                if (Marshal.IsComObject(control))
                                    Marshal.ReleaseComObject(control);
                            }
                        }
                        if (Marshal.IsComObject(docRange))
                            Marshal.ReleaseComObject(docRange);
                        if (Marshal.IsComObject(headerControls))
                            Marshal.ReleaseComObject(headerControls);

                        if (Marshal.IsComObject(wordHeader))
                            Marshal.ReleaseComObject(wordHeader);
                    }

                    if (Marshal.IsComObject(wordSection))
                        Marshal.ReleaseComObject(wordSection);
                }
            }
            catch (Exception ex)
            {
                logger.LogException(LogLevel.Error, "GetHeaderContentControl", ex);
            }
            return contentControl;
        }

        private static Word.ContentControl GetFooterContentControl (Word.Document doc, string name)
        {
            Word.ContentControl contentControl = null;
            try
            {
                foreach (Microsoft.Office.Interop.Word.Section wordSection in doc.Sections)
                {
                    foreach (Microsoft.Office.Interop.Word.HeaderFooter wordFooter in wordSection.Footers)
                    {
                        Microsoft.Office.Interop.Word.Range docRange = wordFooter.Range;

                        ContentControls footerControls = docRange.ContentControls;

                        if (footerControls != null)
                        {
                            foreach (Microsoft.Office.Interop.Word.ContentControl control in footerControls)
                            {
                                if (control.Tag.ToUpper().Equals(name.ToUpper()))
                                {
                                    contentControl = control;
                                    break;
                                }
                                if (Marshal.IsComObject(control))
                                    Marshal.ReleaseComObject(control);
                            }
                        }
                        if (Marshal.IsComObject(docRange))
                            Marshal.ReleaseComObject(docRange);
                        if (Marshal.IsComObject(footerControls))
                            Marshal.ReleaseComObject(footerControls);

                        if (Marshal.IsComObject(wordFooter))
                            Marshal.ReleaseComObject(wordFooter);
                    }

                    if (Marshal.IsComObject(wordSection))
                        Marshal.ReleaseComObject(wordSection);
                }
            }
            catch (Exception ex)
            {
                logger.LogException(LogLevel.Error, "GetFooterContentControl", ex);
            }
            return contentControl;
        }

        private static Word.ContentControl GetBodyContentControl (Word.Document doc, string name)
        {
            Word.ContentControl contentControl = null;

            try
            {
                foreach (Microsoft.Office.Interop.Word.ContentControl control in doc.ContentControls)
                {
                    if (control.Tag.ToUpper().Equals(name.ToUpper()))
                    {
                        contentControl = control;
                        break;
                    }

                    if (Marshal.IsComObject(control))
                        Marshal.ReleaseComObject(control);
                }
            }
            catch (Exception ex)
            {
                logger.LogException(LogLevel.Error, "GetBodyContentControl", ex);
            }

            return contentControl;
        }

        private static List<Word.ContentControl> GetBodyCtrls (Word.Document doc, string name, bool contains = true)
        {
            List<Microsoft.Office.Interop.Word.ContentControl> contentControls = new List<Word.ContentControl>();

            try
            {
                foreach (Microsoft.Office.Interop.Word.ContentControl control in doc.ContentControls)
                {
                    if (contains)
                    {
                        if (control.Tag.ToUpper().Contains(name.ToUpper()))
                        {
                            contentControls.Add(control);

                        }
                    }
                    else
                    {
                        if (control.Tag.ToUpper().Equals(name.ToUpper()))
                        {
                            contentControls.Add(control);

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logger.LogException(LogLevel.Error, "GetBodyCtrls", ex);
            }

            return contentControls;
        }

        public static void SaveDocumentAsPDF (Word.Document document, string targetPath)
        {

            if (Path.GetExtension(targetPath).ToUpper() != "PDF")
                targetPath = Path.Combine(Path.GetDirectoryName(targetPath), string.Concat(Path.GetFileNameWithoutExtension(targetPath), ".pdf"));
            try
            {
                if (File.Exists(targetPath))
                    File.Delete(targetPath);

                Word.WdExportFormat exportFormat = Word.WdExportFormat.wdExportFormatPDF;
                bool openAfterExport = false;
                Word.WdExportOptimizeFor exportOptimizeFor = Word.WdExportOptimizeFor.wdExportOptimizeForPrint;
                Word.WdExportRange exportRange = Word.WdExportRange.wdExportAllDocument;
                int startPage = 0;
                int endPage = 0;
                Word.WdExportItem exportItem = Word.WdExportItem.wdExportDocumentContent;
                bool includeDocProps = false;
                bool keepIRM = true;
                Word.WdExportCreateBookmarks createBookmarks = Word.WdExportCreateBookmarks.wdExportCreateWordBookmarks;
                bool docStructureTags = true;
                bool bitmapMissingFonts = true;
                //11/05/2017
                bool useISO19005_1 = true;
                object missing = Missing.Value;

                // Export it in the specified format.  
                if (document != null)
                    document.ExportAsFixedFormat(
                        targetPath,
                        exportFormat,
                        openAfterExport,
                        exportOptimizeFor,
                        exportRange,
                        startPage,
                        endPage,
                        exportItem,
                        includeDocProps,
                        keepIRM,
                        createBookmarks,
                        docStructureTags,
                        bitmapMissingFonts,
                        useISO19005_1,
                        ref missing);
            }
            catch (System.IO.IOException ex)
            {
                logger.LogException(LogLevel.Error, "SaveDocumentAsPDF", ex);
                throw ex;
            }
            catch (Exception ex)
            {
                logger.LogException(LogLevel.Error, "SaveDocumentAsPDF", ex);
            }
        }

        private static void ByteArrayToImageFilebyMemoryStream (byte[] imageByte, string jpegFile)
        {
            if (imageByte == null)
                return;

            try
            {
                if (File.Exists(jpegFile))
                    File.Delete(jpegFile);

                using (MemoryStream ms = new MemoryStream(imageByte))
                {
                    using (Bitmap gif = new Bitmap(ms))
                    {
                        gif.Save(jpegFile);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.LogException(LogLevel.Error, "ByteArrayToImageFilebyMemoryStream", ex);
            }
        }

        internal static void ByteArrayToImageFilebyMemoryStream (SignatureInfo signatureInfo)
        {
            ByteArrayToImageFilebyMemoryStream(signatureInfo.PictureByteArray, signatureInfo.PicturePath);
        }
        internal static void SaveDocumentAsWordWithoutSignatures (Word.Document document, string targetPath)
        {
            try
            {
                object missing = Missing.Value;
                //Save as word document as well - this need to be the text without the signatures
                string wordVersion = Path.Combine(Path.GetDirectoryName(targetPath), (Path.GetFileNameWithoutExtension(targetPath)));
                if (File.Exists(wordVersion))
                    File.Delete(wordVersion);
                //08/05/2017 - add password
                try
                {
                    object noReset = false;
                    object password = "NEOPHYTE";
                    object useIRM = false;
                    object enforceStyleLock = false;

                    document.Protect(Word.WdProtectionType.wdAllowOnlyReading, ref noReset,
                        ref password, ref useIRM, ref enforceStyleLock);
                }
                catch (Exception ex)
                {
                    logger.LogException(LogLevel.Error, "SaveDocumentAsWordWithoutSignatures", ex);
                }
                //24/05/2017
                document.SaveAs(
                  wordVersion,
                  ref missing, //WdSaveFormat.wdFormatDocument,
                  ref missing,
                  ref missing,
                  ref missing,
                  ref missing,
                  ref missing,
                  ref missing,
                  ref missing,
                  ref missing,
                  ref missing,
                  ref missing,
                  ref missing,
                  ref missing,
                  ref missing,
                  ref missing
                  );
                //24/05/2017
                //remove password as this document will be used to insert pictures
                document.Unprotect("NEOPHYTE");
            }
            catch (System.IO.IOException ex)
            {
                logger.LogException(LogLevel.Error, "SaveDocumentAsWordWithoutSignatures", ex);
                throw ex;
            }
            catch (Exception ex)
            {
                logger.LogException(LogLevel.Error, "SaveDocumentAsWordWithoutSignatures", ex);
            }
        }
    }

}

