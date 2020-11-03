using KABService.Object;
using Microsoft.Extensions.Configuration;
using System;
using System.IO;
using System.Linq;

namespace UtilityLibrary.Log
{
    class LogHelper
    {
        private readonly IConfiguration _configuration;
        private readonly string _identifier;
        public LogHelper(IConfiguration configuration, string identifier)
        {
            _configuration = configuration;
            _identifier = identifier;
        }
        public void InsertLog(LogObject _logObject)
        {
            try
            {
                var workingDirectory = _configuration.GetValue<string>("WorkingDirectory:" + _configuration.GetValue<string>("Environment"));
                // get configuration
                DirectoryInfo logDirectoryInfo = new DirectoryInfo(Path.Combine(workingDirectory, ConfigVariables.LogDirectoryName));
                // create a new sub directory if not created yet
                string today = DateTime.Today.ToString(ConfigVariables.SubDirectoryNameDateFormat);
                DirectoryInfo todayDirectoryInfo = logDirectoryInfo.GetDirectories().FirstOrDefault(x => x.Name.IndexOf(today) > -1);
                if(todayDirectoryInfo == null)
                {
                    todayDirectoryInfo = logDirectoryInfo.CreateSubdirectory(today);
                }
                string fileName = Path.Combine(todayDirectoryInfo.FullName, ConfigVariables.LogFileName + DateTime.Now.ToString(ConfigVariables.LogFileNameDateFormat) + ".txt");
                FileInfo newFile = new FileInfo(fileName);
                using StreamWriter file = new StreamWriter(newFile.FullName, newFile.Exists);
                file.WriteLine(_logObject.ToString(_identifier));
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
