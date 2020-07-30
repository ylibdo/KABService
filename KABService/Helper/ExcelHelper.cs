using KABService.Business_Logic;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace KABService.Helper
{
    class ExcelHelper
    {
        private readonly ILogger<Worker> _logger;
        private readonly IConfiguration _configuration;

        public ExcelHelper(ILogger<Worker> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;
        }

        //Prepare to read data from Excel.
        public ExcelWorksheet Prepare(string _fileName)
        {
            FileInfo file = new FileInfo(_fileName);
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                    return worksheet;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message);
                return null;
            }
        }

        // Data tranformation for new output
        public string Process(ExcelWorksheet _worksheet, string _workingDirectory)
        {
            var fileName = string.Empty;
            var company = this.getCompanyByDirectoryName(_workingDirectory);
            switch (company)
            {
                case "ISTA":
                    // Insert business logic
                    ISTA ista = new ISTA(_logger);
                    fileName = ista.Process(_worksheet, company, _workingDirectory);
                        break;
                case "Company2":
                    // Insert business logic
                    break;
                case "Company3":
                    // Insert business logic
                    break;
                default:
                    _logger.LogWarning("Unknow company/vendor name is found.");
                    break;
            }

            return fileName;
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
                    // Find the company by working directory
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
