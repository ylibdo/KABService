using KABService.Business_Logic;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Text;

namespace KABService.Helper
{
    class CSVHelper
    {
        private readonly ILogger<Worker> _logger;
        private readonly IConfiguration _configuration;

        public CSVHelper(ILogger<Worker> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;
        }

        // Data tranformation for new output
        public string Process(string _workingDirectory, string _fileName)
        {
            var fileName = string.Empty;
            var company = this.getCompanyByFileName(_fileName);
            switch (company)
            {
                case "ISTA":
                    // Insert business logic
                    ISTA ista = new ISTA(_logger);
                    fileName = ista.ProcessCSV(company, _workingDirectory, _fileName);
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

        private string getCompanyByFileName(string _fileName)
        {
            string company = string.Empty;
            var companyString = _configuration.GetValue<string>("Company");
            try
            {
                var companyArray = companyString.Split(";");
                foreach (string c in companyArray)
                {
                    if (_fileName.IndexOf(c) > 0)
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
