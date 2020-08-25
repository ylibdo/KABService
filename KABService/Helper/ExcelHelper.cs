using KABService.Models;
using KABService.Object;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

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

        // Data tranformation for new output
        public DataTable ReadDataAsDataTable(string _fileName, string _sheetName, string _excelVersion)
        {
            string sqlQuery = "Select * From " + _sheetName;
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + _fileName + ";Extended Properties=\"Excel " + _excelVersion + ";HDR=YES;\"";

            OleDbConnection con = new OleDbConnection(connectionString + "");
            OleDbDataAdapter da = new OleDbDataAdapter(sqlQuery, con);
            da.Fill(ds);

            dt = ds.Tables[0];

            return dt;
        }

        public string SaveDataToFile(IEnumerable<DataRow> _input, FactorModel _factorModel, string _company, string _workingDirectory)
        {
            try
            {
                string newExcelFileName = string.Concat(_company, ConfigVariables.OutputFileSeparator, DateTime.Now.ToString(ConfigVariables.OutputFileDateFormatString), ConfigVariables.OutputFileSuffixExcel);
                FileInfo newExcelFile = new FileInfo(Path.Combine(_workingDirectory, newExcelFileName));
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                ExcelPackage package = new ExcelPackage(newExcelFile);

                string regexPattern = @"\d{4}";

                Match department = Regex.Match(_company, regexPattern);

                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(department + " AC");

                //Insert headers
                worksheet.Cells[1, 1].Value = ConfigVariables.Company;
                worksheet.Cells[1, 2].Value = "Afdeling";
                worksheet.Cells[1, 3].Value = "Lejlighed";
                worksheet.Cells[1, 4].Value = "Målertype";
                worksheet.Cells[1, 5].Value = "Serienr";
                worksheet.Cells[1, 6].Value = "Aflæsningsdato";
                worksheet.Cells[1, 7].Value = "Aflæsning";
                worksheet.Cells[1, 8].Value = "Faktor";
                worksheet.Cells[1, 9].Value = "Reduktion";
                worksheet.Cells[1, 10].Value = "Lokale";
                worksheet.Cells[1, 11].Value = "Installation";
                worksheet.Cells[1, 12].Value = "Deaktiveringsdato";
                worksheet.Cells[1, 13].Value = "Bemærkning";
                worksheet.Cells[1, 14].Value = "Nulstillingsmåler";

                //Insert apartment data
                worksheet.Cells["C2"].LoadFromCollection(_input.AsEnumerable().Select(x => x[_factorModel.ApartmentColumn].ToString()).ToList());
                //Insert maaler data
                worksheet.Cells["D2"].LoadFromCollection(_input.AsEnumerable().Select(x => x[_factorModel.MaalerColumn].ToString()).ToList());
                //Insert SerieID's
                worksheet.Cells["E2"].LoadFromCollection(_input.AsEnumerable().Select(x => x[_factorModel.SerieIDColumn].ToString()).ToList());
                //Insert read-date data
                worksheet.Cells["F2"].LoadFromCollection(_input.AsEnumerable().Select(x => x[_factorModel.ReadDateColumn].ToString()).ToList());
                //Insert read data
                worksheet.Cells["G2"].LoadFromCollection(_input.AsEnumerable().Select(x => x[_factorModel.ReadColumn].ToString()).ToList());
                //Insert faktor data
                worksheet.Cells["H2"].LoadFromCollection(_input.AsEnumerable().Select(x => x[_factorModel.FaktorColumn].ToString()).ToList());
                //Insert reduction data
                worksheet.Cells["I2"].LoadFromCollection(_input.AsEnumerable().Select(x => x[_factorModel.ReductionColumn].ToString()).ToList());
                //Insert room data
                worksheet.Cells["J2"].LoadFromCollection(_input.AsEnumerable().Select(x => x[_factorModel.RoomColumn].ToString()).ToList());
                //Insert installation data
                worksheet.Cells["K2"].LoadFromCollection(_input.AsEnumerable().Select(x => x[_factorModel.InstallationDateColumn].ToString()).ToList());

                //Insert Selskab
                worksheet.Cells[2, 1, _input.ToList().Count + 1, 1].Value = department.ToString().Substring(0, 2);
                //Insert Afdeling
                worksheet.Cells[2, 2, _input.ToList().Count + 1, 2].Value = department.ToString().Substring(2, 2);
                package.SaveAs(newExcelFile);

                package.Dispose();

                return newExcelFile.FullName;
            }
            
            catch(Exception)
            {
                return string.Empty;
            }
        }
    }
}
