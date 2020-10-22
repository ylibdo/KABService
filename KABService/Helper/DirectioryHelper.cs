using KABService.Object;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace KABService.Helper
{
    class DirectioryHelper
    {
        private readonly ILogger<Worker> _logger;
        private readonly IConfiguration _configuration;

        public DirectioryHelper(ILogger<Worker> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;
        }
        /// <summary>
        /// Method to check if files exist in the target directory
        /// </summary>
        /// <param name="_path">Target directory</param>
        /// <param name="_suffix">Read a perticular type files, if read all file leave it blank. *.xlsx will read all excel files</param>
        /// <returns></returns>
        public bool HasFile(string _path, string _suffix = "")
        {
            bool success;
            if (!Directory.Exists(_path))
            {
                return false;
            }
            else
            {
                // read file from path
                try
                {
                    var files = String.IsNullOrEmpty(_suffix) ? Directory.EnumerateFiles(_path) : Directory.EnumerateFiles(_path, _suffix);
                    success = files.Count() > 0;
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex.Message);
                    return false;
                }
            }
            return success;
        }
        /// <summary>
        /// Move a file to a target directory
        /// </summary>
        /// <param name="_workingDirectory">Current working directory</param>
        /// <param name="_fileName">File name</param>
        /// <param name="_fileMoveOption">Move option to specify which directory it will go</param>
        public void MoveFile(string _workingDirectory, string _fileName, BDOEnum.FileMoveOption _fileMoveOption)
        {
            try
            {
                var targetPath = Path.Combine(_workingDirectory, _fileMoveOption.ToString());
                if (!Directory.Exists(targetPath))
                {
                    Directory.CreateDirectory(targetPath);
                }
                FileInfo targetFileInfo = new FileInfo(Path.Combine(targetPath, _fileName.Substring(_workingDirectory.Length + 1)));
                if(targetFileInfo.Exists)
                {
                    targetFileInfo = new FileInfo(Path.Combine(targetPath, "copy_" + _fileName.Substring(_workingDirectory.Length + 1)));
                }
                Directory.Move(_fileName, targetFileInfo.FullName);

                _logger.LogInformation("File: " + _fileName + " is moved with action " + _fileMoveOption.ToString());
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message);
            }
        }

        /// <summary>
        /// Get list of working directories for each company/vendor
        /// </summary>
        public IEnumerable<string> GetAllWorkingDirectoryFullPath()
        {
            List<string> companyList = new List<string>();
            var directory = _configuration.GetValue<string>("WorkingDirectory:" + _configuration.GetValue<string>("Environment"));
            var companyString = _configuration.GetValue<string>("Company");
            try
            {
                var companyArray = companyString.Split(";");
                foreach(string c in companyArray)
                {
                    companyList.Add(Path.Combine(directory, c));
                }
            }
            catch(Exception ex)
            {
                _logger.LogError(ex.Message);
            }
            return companyList;
        }

    }
}
