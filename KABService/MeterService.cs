using KABService.Helper;
using KABService.Models;
using KABService.Object;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Data;
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
                        string excelVersion = string.Empty;
                        foreach(var file in files)
                        {
                            try
                            {
                                FileInfo fileInfo = new FileInfo(file);
                                string newFileName = string.Empty;
                                DataTable outputDataTable = new DataTable();

                                // 3. create factor object based on company
                                var company = getCompanyByDirectoryName(directory);

                                FactorModel factorModel = createFactorModelByCompany(company);
                               
                                if(factorModel == null)
                                {
                                    throw new Exception("Vendor configuration error. Cannot create factor model based based on vendor information.");
                                }

                                // 4. Reading from files.
                                CSVHelper csvHelper = new CSVHelper(_logger, _configuration);
                                ExcelHelper excelHelper = new ExcelHelper(_logger, _configuration);
                                if (fileInfo.Extension == ".csv")
                                {
                                    // handling CSV files
                                    
                                    outputDataTable = csvHelper.ReadDataAsDataTable(file);
                                }
                                else if (fileInfo.Extension == ".xlsx")
                                {
                                    // handling Excel files
                                    excelVersion = "12.0";
                                    outputDataTable = excelHelper.ReadDataAsDataTable(file, directory, excelVersion);
                                }
                                else if(fileInfo.Extension == ".xls")
                                {
                                    // handling Excel files
                                    excelVersion = "8.0";
                                    outputDataTable = excelHelper.ReadDataAsDataTable(file, directory, excelVersion);
                                }
                                else
                                {
                                    throw new FileLoadException("File format is not supported");
                                }

                                // 5. Filter source data.
                                IEnumerable<DataRow> filteredData = filterDataByCompany(outputDataTable, factorModel, company);

                                // 6. Save data to file.
                                newFileName = csvHelper.SaveDataToFile(filteredData, company, directory);

                                if (String.IsNullOrEmpty(newFileName))
                                {
                                    throw new NullReferenceException();
                                }
                                else
                                {
                                    directioryHelper.MoveFile(directory, newFileName, BDOEnum.FileMoveOption.Processed);
                                }
                                // 7. Move processed file to archive
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

        private FactorModel createFactorModelByCompany(string _company)
        {
            switch (_company)
            {
                case "1902 Ista":
                    return new FactorModel()
                    {
                        CompanyColumn = 0,
                        DepartmentColumn = 1,
                        ApartmentColumn = 2,
                        MaalerColumn = 4,
                        SerieIDColumn = 5,
                        ReadDateColumn = 7,
                        ReadColumn = 8,
                        FaktorColumn = 9,
                        ReductionColumn = 10,
                        RoomColumn = 11,
                        InstallationDateColumn = 12,
                        MaalerType = "Doprimo 3 SoC"
                    };
                default:
                    _logger.LogWarning("Unknow company/vendor name is found.");
                    return null;
            }
        }

        private IEnumerable<DataRow> filterDataByCompany(DataTable _input, FactorModel _factorModel, string _company)
        {
            switch (_company)
            {
                case "1902 Ista":
                    return _input.AsEnumerable().Where(x => x[_factorModel.ReadDateColumn].Equals(_factorModel.SearchCriteria) && x[_factorModel.MaalerColumn].ToString().Contains(_factorModel.MaalerType));
                default:
                    _logger.LogWarning("Unknow company/vendor name is found.");
                    return null;
            }
        }
    }
}
