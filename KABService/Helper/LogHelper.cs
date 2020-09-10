using KABService.Object;
using System;
using System.IO;
using System.Linq;

namespace UtilityLibrary.Log
{
    class LogHelper
    {
        public static void InsertLog(LogObject _logObject)
        {
            try
            {
                // get configuration
                DirectoryInfo logDirectoryInfo = new DirectoryInfo(ConfigVariables.LogDirectory);
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
                file.WriteLine(_logObject.ToString());
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
