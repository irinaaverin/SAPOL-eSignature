using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace SAPOLEStatement
{
    public class TemplateDetails
    {
        public TemplateDetails()
        {
            SignatureInfoList = new List<SignatureInfo>();
        }
        public TemplateDetails(bool isSAPOL)
        {
            SignatureInfoList = new List<SignatureInfo>();
            IsSAPOL = isSAPOL;
        }
        //public Word.Document Document { get; set; }
        //public bool IsFinalised { get; set; }

        public List<SignatureInfo> SignatureInfoList { get; set; }
        public string Password { get; set; }
        public string DocumentPath { get; set; }
        public string PDFFullName { get; set; }
        public bool IsSAPOL { get; set; }
    }

    public class SignatureInfo
    {
        public int Number { get; set; }
        public string Title { get; set; }
        public string Tag { get; set; }
        public string PicturePath { get; set; }
        public byte[] PictureByteArray { get; set; }
        public Word.ContentControl SignatureControl { get; set; }
        public string SignatureControlPosition { get; set; }
    }
}
