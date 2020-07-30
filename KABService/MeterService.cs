using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using KABService.Helper;
using KABService.Object;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;

namespace KABService
{
    class MeterService
    {
        private readonly ILogger<Worker> _logger;
        private readonly IConfiguration _configuration;

        public MeterService(ILogger<Worker> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;
        }
        // Every action starts from here.
        public void Run()
        {
            try
            {
                // step 1. Get all the files from directories
                DirectioryHelper directioryHelper = new DirectioryHelper(_logger, _configuration);
                IEnumerable<string> workingDirectories = directioryHelper.GetAllWorkingDirectoryFullPath();
                // step 2. Read data from files
                foreach(string directory in workingDirectories)
                {
                    // step 2.1 Check if the directory is empty
                    if(directioryHelper.HasFile(directory))
                    {
                        var files = Directory.EnumerateFiles(directory);
                        foreach(var file in files)
                        {
                            try
                            {
                                // step 2.2 Prepare work sheet to read File
                                ExcelHelper excelHelper = new ExcelHelper(_logger, _configuration);
                                ExcelWorksheet excelWorksheet = excelHelper.Prepare(file);

                                // step 3. Data transformation and get output
                                var newFileName = excelHelper.Process(excelWorksheet, directory);
                                if(String.IsNullOrEmpty(newFileName))
                                {
                                    throw new NullReferenceException();
                                }
                                else
                                {
                                    directioryHelper.MoveFile(directory, newFileName, BDOEnum.FileMoveOption.Processed);
                                }
                                // step 4. Move processed file to archive
                                directioryHelper.MoveFile(directory, file, BDOEnum.FileMoveOption.Archive);
                            }
                            catch (Exception ex)
                            {
                                _logger.LogError("Processing file: " + file + " is failed.");
                                _logger.LogError(ex.Message);
                                // step 4. Move processed file to error
                                directioryHelper.MoveFile(directory, file, BDOEnum.FileMoveOption.Error);
                            }
                        }
                    }
                    else
                    {
                        _logger.LogInformation(directory + " has no file to work on.");
                    }
                }

            }
            catch(Exception ex)
            {
                _logger.LogError(ex.Message);
            }
        }
        
    }
}
