using KABWeb.Controllers;
using KABWeb.Models;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Reflection.Metadata;
using System.Text;

namespace KABWeb.Helper
{
    class ExcelHelper
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IConfiguration _configuration;

        public ExcelHelper(ILogger<HomeController> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;
        }

        //Local data
        public IEnumerable<CommonViewModel> Load(string _fileName)
        {
            FileInfo file = new FileInfo(_fileName);
            if(file.Extension != ".xlsx")
            {
                throw new Exception("File has to be new excel file end in .xlsx");
            }
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension.Rows;
                    int ColCount = worksheet.Dimension.Columns;
                    List<CommonViewModel> list = new List<CommonViewModel>();
                    for (int row = 1; row <= rowCount; row++)
                    {
                        CommonViewModel item = new CommonViewModel
                        {
                            Field1 = getValue(worksheet, row, 1),
                            Field2 = getValue(worksheet, row, 2),
                            Field3 = getValue(worksheet, row, 3),
                            Field4 = getValue(worksheet, row, 4),
                            Field5 = getValue(worksheet, row, 5),
                            Field6 = getValue(worksheet, row, 6),
                            Field7 = getValue(worksheet, row, 7),
                            Field8 = getValue(worksheet, row, 8),
                            Field9 = getValue(worksheet, row, 9),
                            Field10 = getValue(worksheet, row, 10),
                            Field11 = getValue(worksheet, row, 11),
                            Field12 = getValue(worksheet, row, 12),
                            Field13 = getValue(worksheet, row, 13),
                            Field14 = getValue(worksheet, row, 14),
                            Field15 = getValue(worksheet, row, 15)
                        };
                        list.Add(item);
                    }
                    return list;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message);
                return null;
            }
        }

        private string getValue(ExcelWorksheet _excelWorksheet, int _row, int _col)
        {
            string returnValue = string.Empty;
            try
            {
                returnValue = _excelWorksheet.Cells[_row, _col].Text ?? string.Empty;
            }
            catch(Exception)
            {

            }
            return returnValue;
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
                    if(_workingDirectory.IndexOf(c) > 0)
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
