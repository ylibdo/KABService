using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace KABService.Business_Logic
{
    class Casi
    {
        private readonly ILogger<Worker> _logger;

        public Casi(ILogger<Worker> logger)
        {
            _logger = logger;
        }

        // Process data and return newly created file name.
        public string ProcessExcel(ExcelWorksheet _Worksheet, string _company, string _workingDirectory)
        {
            // make a new excel to hold export data
            string newFileName = string.Concat(_company, "_", DateTime.Now.ToString("yyyyMMddHHmmss"), "_unik.xlsx");
            FileInfo newFile = new FileInfo(Path.Combine(_workingDirectory, newFileName));
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage package = new ExcelPackage(newFile))
                {
                    ExcelWorksheet output = package.Workbook.Worksheets.Add("Sheet1");
                    int rowCount = _Worksheet.Dimension.Rows;
                    int ColCount = _Worksheet.Dimension.Columns;
                    for (int row = 1; row <= rowCount; row++)
                    {
                        for (int col = 1; col <= ColCount; col++)
                        {
                            output.Cells[row, col].Value = _Worksheet.Cells[row, col].Value;
                        }
                    }
                    package.Save();
                }
                return newFile.FullName;
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
