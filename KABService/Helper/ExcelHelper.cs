using KABService.Models;
using KABService.Object;
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

        public ExcelHelper()
        {
        }

        // Data tranformation for new output
        
        public DataTable ReadDataAsDataTable(string _fileName, string _excelVersion)
        {
            string sheetName;

            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + _fileName + ";Extended Properties=\"Excel " + _excelVersion + ";HDR=YES;\"";
            OleDbConnection con = new OleDbConnection(connectionString + "");

            //Open the connection and get tables in the Excel sheet
            con.Open();

            dt = con.GetSchema("Tables");

            con.Close();

            // The first sheetname is always index 0 and index 2 in the itemarray
            sheetName = dt.Rows[0].ItemArray[2].ToString();

            string sqlQuery = "Select * From [" + sheetName + "]";
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
                worksheet.Cells[1, 2].Value = ConfigVariables.Department;
                worksheet.Cells[1, 3].Value = ConfigVariables.Apartment;
                worksheet.Cells[1, 4].Value = ConfigVariables.MeterType;
                worksheet.Cells[1, 5].Value = ConfigVariables.SerieID;
                worksheet.Cells[1, 6].Value = ConfigVariables.ReadingDate;
                worksheet.Cells[1, 7].Value = ConfigVariables.Reading;
                worksheet.Cells[1, 8].Value = ConfigVariables.Factor;
                worksheet.Cells[1, 9].Value = ConfigVariables.Reduction;
                worksheet.Cells[1, 10].Value = ConfigVariables.Room;
                worksheet.Cells[1, 11].Value = ConfigVariables.Installation;
                worksheet.Cells[1, 12].Value = ConfigVariables.DeactivationDate;
                worksheet.Cells[1, 13].Value = ConfigVariables.Comment;
                worksheet.Cells[1, 14].Value = ConfigVariables.ResetMeter;

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
