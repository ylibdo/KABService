using KABService.Helper;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using KABService.Object;
using Microsoft.Extensions.Configuration;
using System.Data;

namespace KABService.Business_Logic
{
    class ISTA
    {
        private readonly ILogger<Worker> _logger;
        private readonly IConfiguration _configuration;
        public ISTA(ILogger<Worker> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;
        }

        // Process data and return newly created file name.
        public string ProcessExcel( string strDepartment, string _workingDirectory)
        {
                        
            try
            {
                int companyColumn = 0;
                int departmentColumn = 0;
                int apartmentColumn = 0;
                int maalerColumn = 0;
                int serieIDColumn = 0;
                int readDateColumn = 0;
                int readColumn = 0;
                int faktorColumn = 0;
                int reductionColumn = 0;
                int roomColumn = 0;
                int installationDateColumn = 0;
                DateTime searchDate = new DateTime(2019, 12, 31);
                string searchCriteria = searchDate.ToString("yyyy-MM-dd");
                double factor = 0;
                string maalerType = "";
                var notContainingValues = new List<DataRow>();

                
                    
                serieIDColumn = 2;
                readDateColumn = 4;
                readColumn = 5;
                factor = 1;
                maalerType = "WHE37 / CASI";
                notContainingValues = dt.AsEnumerable().Where(x => !x[readColumn].ToString().Contains("s")).ToList();
                      
                
                        return null;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message);
                return string.Empty;
            }
        }

        // Process data and return newly created file name.
        public string ProcessCSV(string _company, string _workingDirectory, string _fileName)
        {
            // make a new excel to hold export data
            string newFileName = string.Concat(_company, "_", DateTime.Now.ToString("yyyyMMddHHmmss"), "_unik.csv");
            FileInfo newFile = new FileInfo(Path.Combine(_workingDirectory, newFileName));
            try
            {
                // process csv
                
                return newFile.FullName;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message);
                return string.Empty;
            }
        }
    }
}
