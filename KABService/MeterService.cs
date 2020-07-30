using KABService.Helper;
using KABService.Object;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

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
                                FileInfo fileInfo = new FileInfo(file);
                                string newFileName = string.Empty;

                                if (fileInfo.Extension == ".csv")
                                {
                                    // handling CSV files
                                    CSVHelper csvHelper = new CSVHelper(_logger, _configuration);
                                    newFileName = csvHelper.Process(directory, file);
                                }
                                else if (fileInfo.Extension == ".xlsx")
                                {
                                    // handling Excel files

                                    // step 2.2 Prepare work sheet to read File
                                    ExcelHelper excelHelper = new ExcelHelper(_logger, _configuration);
                                    ExcelWorksheet excelWorksheet = excelHelper.Prepare(file);

                                    // step 3. Data transformation and get output
                                    newFileName = excelHelper.Process(excelWorksheet, directory);
                                }
                                else
                                {
                                    throw new FileLoadException("File format is not supported");
                                }
                                
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
