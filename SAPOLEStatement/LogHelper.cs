using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NLog;
using NLog.Config;
using NLog.Targets;
using System.IO;
using System.Configuration;

namespace SAPOLEStatement
{
    internal static class LogHelper
    {
        static Logger logger = null;

        internal static Logger InitialiseNLog()
        {
            if (logger == null)
            {
                logger = LogManager.GetCurrentClassLogger();
                //1.Create configuration object
                LoggingConfiguration config = new LoggingConfiguration();

                //2. Create target and set it to the config
                FileTarget fileTarget = new FileTarget();
                config.AddTarget("file", fileTarget);

                //3. Set target properties
                //fileTarget.FileName = "${basedir}/Logs/statuslog.log";
                fileTarget.FileName = Path.Combine(ConfigurationManager.AppSettings["LogFilePath"].ToString(), "status.log");                
                fileTarget.Layout = "${longdate} | ${level} | ${message} | ${exception:format=tostring,message,method:maxInnerExceptionLevel=5:innerFormat=shortType,message,method}";
                //fileTarget.ArchiveFileName = "${basedir}/LogArchives/statuslog.{##}.log";
                //fileTarget.ArchiveEvery = FileArchivePeriod.Day;
                //fileTarget.ArchiveNumbering = ArchiveNumberingMode.Rolling;
                //fileTarget.MaxArchiveFiles = 31;
                fileTarget.ConcurrentWrites = true;
                fileTarget.KeepFileOpen = false;

                //4. Define rules
                LoggingRule rule = new LoggingRule("*", LogLevel.Info, fileTarget);
                config.LoggingRules.Add(rule);

                //5. Activate configuration
                LogManager.Configuration = config;
                logger = LogManager.GetLogger("SAPOLEStatement");
            }
            return logger;
        }
    }
}
