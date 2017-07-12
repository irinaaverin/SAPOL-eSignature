using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace SAPOLEStatement
{
    public class EventLog
    {
        private readonly StreamWriter streamWriter;

        public EventLog(string filePrefix, string filePath)
        {
            

            try
            {
                streamWriter = new StreamWriter(Path.Combine(filePath, LogFileName(filePrefix)));
                streamWriter.AutoFlush = true;
            }
            catch (Exception ex)
            {
                throw new EventLogException("Error in EventLog Constructor", ex);
            }
        }
        private string LogFileName(string filePrefix)
        {
            return filePrefix + " " + DateTime.Now.ToString("yyyyMMdd HHmm") + ".log";
        }

        public void Write( Exception exception)
        {
            Write(exception.GetType().Name + ": " + exception.Message);
            if (exception.InnerException != null)
                Write( exception.InnerException.GetType().Name + ": " + exception.InnerException.Message);
        }
        public void Write( string description)
        {
            DateTime now = DateTime.Now;

            if (streamWriter != null)
            {
                streamWriter.WriteLine(now + ": " + description);
            }            
        }

        public void Close()
        {
            if (streamWriter != null)
            {
                streamWriter.Close();
            }
        }

    }

}
