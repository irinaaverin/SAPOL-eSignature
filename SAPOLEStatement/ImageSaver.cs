using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using Microsoft.Ink;
using System.Drawing.Imaging;
using System.Windows.Forms;
using NLog;
using System.Configuration;

namespace SAPOLEStatement
{
    internal class ImageSaver
    {        
        private Logger logger = LogManager.GetCurrentClassLogger();
        private string FILE_NAME;

        internal ImageSaver()
        {
            logger = LogHelper.InitialiseNLog();
            try
            {
                FILE_NAME = ConfigurationManager.AppSettings["TemporarySignature"].ToString().ToUpper();

            }
            catch (Exception ex)
            {
                logger.LogException(LogLevel.Error, "ImageSaver", ex);
                FILE_NAME = AppDomain.CurrentDomain.BaseDirectory;
            }
        }
        public void ClearInkPicture(InkPicture inkPictureCtrl)
        {
            inkPictureCtrl.DefaultDrawingAttributes.Color = Color.Black;
            inkPictureCtrl.Ink.DeleteStrokes();

            Ink ink = new Ink();
            inkPictureCtrl.InkEnabled = false;
            inkPictureCtrl.Ink = ink;
            inkPictureCtrl.InkEnabled = true;

            inkPictureCtrl.Invalidate();
            inkPictureCtrl.Refresh();
        }

        public bool SaveSignatureToArray(InkPicture inkPictureCtrl, ref SignatureInfo item)
        {           
            try
            {
                string fileToSaveGif = string.Empty;
                if (inkPictureCtrl.Ink.Strokes.Count == 0)
                    return true;

                if (!Directory.Exists(FILE_NAME))
                {
                    MessageBox.Show(string.Format("Electronic Signatures were not transferred to the PDF file. {0}Please  contact your system Administrator.", 
                        Environment.NewLine), "Error on saving document as PDF");
                    logger.Log(LogLevel.Error, string.Format("Directory {0} does not exist.",FILE_NAME));
                    return false;
                }
                fileToSaveGif = Path.Combine(FILE_NAME, string.Concat(DateTime.Now.Ticks, ".Gif"));

                if (!string.IsNullOrEmpty(fileToSaveGif))
                {
                    if (File.Exists(fileToSaveGif))
                        File.Delete(fileToSaveGif);
                }

                byte[] bytes = (byte[])inkPictureCtrl.Ink.Save(PersistenceFormat.Gif, CompressionMode.NoCompression);
                item.PictureByteArray = bytes;
                item.PicturePath = fileToSaveGif;
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Electronic Signatures were not transferred to the PDF file. {0}. Pleae contact system Administrator.", Environment.NewLine), "Saving documnet as PDF");

                logger.LogException(LogLevel.Error, "SaveSignatureToArray", ex);
                return false;
            }
            return true;
        }

    }

}
