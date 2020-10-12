﻿using KABService.Helper;
using KABService.Models;
using KABService.Object;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using UtilityLibrary.Log;
using static UtilityLibrary.Log.LogObject;

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
                    LogHelper.InsertLog(new LogObject(LogType.Information, ("Start to process" + directory)));
                    // step 2.1 Check if the directory is empty
                    if(directioryHelper.HasFile(directory))
                    {
                        var files = Directory.EnumerateFiles(directory);
                        string excelVersion = string.Empty;
                        CSVHelper csvHelper = new CSVHelper(_logger, _configuration);
                        ExcelHelper excelHelper = new ExcelHelper(_logger, _configuration);
                        DataTable unikDataTable = new DataTable();

                        // 2.2 Read the Unik data
                        //unikDataTable = excelHelper.ReadDataAsDataTable("E:\\KAB Services\\MeterService\\UnikData\\UnikData.xlsx", "12.0");
                        try
                        {
                            unikDataTable = excelHelper.ReadDataAsDataTable(_configuration.GetValue<string>("Unik:Vendor"), "12.0");
                        }
                        catch(Exception)
                        {
                            LogHelper.InsertLog(new LogObject(LogType.Error, "Reading Unik configuration file failed."));
                        }
                        foreach (var file in files)
                        {
                            try
                            {
                                FileInfo fileInfo = new FileInfo(file);
                                string newFileName = string.Empty;
                                string newErrorFileName = string.Empty;
                                DataTable outputDataTable = new DataTable();


                                // 3. Get company name based on directory
                                var company = getCompanyByDirectoryName(directory);

                                LogHelper.InsertLog(new LogObject(LogType.Information, ("Ready to process file: " + fileInfo.FullName)));
                                // 4. Reading from files.
                                if (fileInfo.Extension == ".csv")
                                {
                                    // handling CSV files
                                    
                                    outputDataTable = csvHelper.ReadDataAsDataTable(file);

                                    var stringArray = string.Join(", ", outputDataTable.Rows[0].ItemArray);

                                    var isNumeric = Regex.Match(stringArray, "[0-9]", RegexOptions.IgnoreCase).Success;

                                    if (!isNumeric)
                                    {
                                        outputDataTable.Rows[0].Delete();
                                        outputDataTable.AcceptChanges();
                                    }
                                    
                                }
                                else if (fileInfo.Extension == ".xlsx" || fileInfo.Extension == ".xlsb")
                                {
                                    // handling Excel files
                                    excelVersion = "12.0";
                                    outputDataTable = excelHelper.ReadDataAsDataTable(file, excelVersion);
                                }
                                else if(fileInfo.Extension == ".xls")
                                {
                                    // handling Excel files
                                    excelVersion = "8.0";
                                    outputDataTable = excelHelper.ReadDataAsDataTable(file, excelVersion);
                                }
                                else
                                {
                                    LogHelper.InsertLog(new LogObject(LogType.Warning, (fileInfo.Extension+ " file is not supported.")));
                                    throw new FileLoadException("File format is not supported");
                                }

                                // 4.1 Create factormodel based on company, outputdata and data from unik

                                FactorModel factorModel = BusinessLogic.CreateFactorModelByCompany(company, unikDataTable, outputDataTable);

                                if (factorModel == null)
                                {
                                    throw new Exception("Vendor configuration error. Cannot create factor model based based on vendor information.");
                                }


                                // 5. Filter source data.
                                IEnumerable<DataRow> filteredData = BusinessLogic.FilterDataByCompany(outputDataTable, factorModel, company);

                                IEnumerable<DataRow> errorData = BusinessLogic.ErrorDataByCompany(outputDataTable, factorModel, company);

                                // 6. Save data to file.
                                newFileName = csvHelper.SaveDataToFile(filteredData, factorModel, company, directory, ConfigVariables.OutputFileNameSuffix);

                                newErrorFileName = csvHelper.SaveDataToFile(errorData, factorModel, company, directory, ConfigVariables.ErrorFileNameSuffix);


                                if (String.IsNullOrEmpty(newFileName))
                                {
                                    throw new NullReferenceException();
                                }
                                else
                                {
                                    directioryHelper.MoveFile(directory, newFileName, BDOEnum.FileMoveOption.Processed);
                                    LogHelper.InsertLog(new LogObject(LogType.Information, (newFileName + " has been moved to Processed dictory")));
                                }

                                if (String.IsNullOrEmpty(newErrorFileName))
                                {
                                    throw new NullReferenceException();
                                }
                                else
                                {
                                    LogHelper.InsertLog(new LogObject(LogType.Information, (newErrorFileName + " has been moved to Manual dictory")));
                                    directioryHelper.MoveFile(directory, newErrorFileName, BDOEnum.FileMoveOption.Manual);
                                }
                                // 7. Move processed file to archive
                                directioryHelper.MoveFile(directory, file, BDOEnum.FileMoveOption.Archive);
                                LogHelper.InsertLog(new LogObject(LogType.Information, (file + " has been moved to Archive dictory")));
                            }
                            catch (Exception ex)
                            {
                                _logger.LogError("Processing file: " + file + " is failed.");
                                _logger.LogError(ex.Message);
                                // step 4. Move processed file to error
                                directioryHelper.MoveFile(directory, file, BDOEnum.FileMoveOption.Error);
                                LogHelper.InsertLog(new LogObject(LogType.Information, (file + " has been moved to Error dictory")));
                            }

                            // send notification
                            SMTPHelper smtpHelper = new SMTPHelper(_configuration);
                            smtpHelper.SendEmailAsync("email body");
                        }
                    }
                    else
                    {
                        _logger.LogInformation(directory + " has no file to work on.");
                        LogHelper.InsertLog(new LogObject(LogType.Information, (directory + " has no file to work on.")));
                    }
                }

            }
            catch(Exception ex)
            {
                _logger.LogError(ex.Message);
            }
        }

        private string getCompanyByDirectoryName(string _workingDirectory)
        {
            string company = string.Empty;
            var companyString = _configuration.GetValue<string>("Company");
            try
            {
                var companyArray = companyString.Split(";");
                foreach (string c in companyArray)
                {
                    if (_workingDirectory.IndexOf(c) > 0)
                    {
                        company = c;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message);
            }

            return company;
        }

                
    }
}
