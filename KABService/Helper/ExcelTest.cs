using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using OfficeOpenXml;
using System.Linq;
using System.Data;
using System.Security.Cryptography.X509Certificates;
using System.Data.OleDb;
using System.Globalization;
using System.Threading;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Configuration;
using CsvHelper;
using System.ComponentModel;
using System.Text.Json.Serialization;
using KABService.Models;
using System.Text.RegularExpressions;
using KABService.Object;

namespace KABService.Helper
{
    public class ExcelTest
    {
        private readonly ILogger<Worker> _logger;
        private readonly IConfiguration _configuration;
        public ExcelTest(ILogger<Worker> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;
        }
    

        public void ProcessExcelFile(string path, string _workingDirectory)
        {
            
            string _company = getCompanyByDirectoryName(_workingDirectory);

            string newFileName = string.Concat(_company, "_", DateTime.Now.ToString("yyyyMMddHHmmss"), "_unik.xlsx");
            FileInfo newFile = new FileInfo(Path.Combine(_workingDirectory, newFileName));

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            ExcelPackage package = new ExcelPackage(newFile);

            string regexPattern = @"\d{4}";

            Match department = Regex.Match(_company, regexPattern);

            Console.WriteLine(department.Value);

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

            
            DataTable dt = ImportExceltoDatatable(path);

            Console.WriteLine("dt rows count: " + dt.Rows.Count);

            //BusinessLogic

            int serieIDColumn = 0;
            int readDateColumn = 0;
            int readColumn = 0;
            int faktorColumn = 0;
            int maalerColumn = 0;
            int apartmentColumn = 0;
            int reductionColumn = 0;
            int roomColumn = 0;
            int installationDateColumn = 0;
            DateTime searchDate = new DateTime(2019, 12, 31);
            string searchCriteria = searchDate.ToString("yyyy-MM-dd");
            /*string criteria = new DateTime(2019, 12, 31).ToString("yyyy-MM-dd");
            DateTime dateCriteria = Convert.ToDateTime(criteria);
            var dateTimeCriteria = DateTime.ParseExact(criteria, "yyyy-MM-dd", CultureInfo.GetCultureInfoByIetfLanguageTag("en-US"));*/
            double factor = 0;
            string maalerType = "";
            var notContainingValues = new List<DataRow>();

            switch (department.ToString())
            {
                case "1008":
                    serieIDColumn = 2;
                    readDateColumn = 4;
                    readColumn = 5;
                    factor = 1;
                    maalerType = "WHE37 / CASI";
                    notContainingValues = dt.AsEnumerable().Where(x => !x[readColumn].ToString().Contains("s")).ToList();
                    break;

                case "1902":
                    serieIDColumn = 5;
                    readDateColumn = 7;
                    readColumn = 8;
                    factor = 0.123;
                    faktorColumn = 9;
                    maalerColumn = 4;
                    apartmentColumn = 2;
                    reductionColumn = 10;
                    roomColumn = 11;
                    installationDateColumn = 12;
                    maalerType = "Doprimo 3 SoC";
                    notContainingValues = dt.AsEnumerable().Where(x => x[readDateColumn].Equals(searchCriteria) && x[maalerColumn].ToString().Contains(maalerType)).ToList();
                    break;

                case "3020":
                    serieIDColumn = 2;
                    readDateColumn = 4;
                    readColumn = 5;
                    factor = 0.123;
                    maalerType = "M7R";
                    break;

                case "3920":
                    serieIDColumn = 2;
                    readDateColumn = 4;
                    readColumn = 5;
                    factor = 1;
                    maalerType = "";
                    break;

                case "4201":
                    serieIDColumn = 2;
                    readDateColumn = 4;
                    readColumn = 5;
                    factor = 1;
                    maalerType = "Doprimo 3 SoC";
                    break;
                default:
                    break;
            }


            //notContainingValues = dt.AsEnumerable().Where(x => !x[readColumn].ToString().Contains("s")).ToList();
            //Insert apartment data
            worksheet.Cells["C2"].LoadFromCollection(notContainingValues.AsEnumerable().Select(x => x[apartmentColumn].ToString()).ToList());
            //Insert maaler data
            worksheet.Cells["D2"].LoadFromCollection(notContainingValues.AsEnumerable().Select(x => x[maalerColumn].ToString()).ToList());
            //Insert SerieID's
            worksheet.Cells["E2"].LoadFromCollection(notContainingValues.AsEnumerable().Select(x => x[serieIDColumn].ToString()).ToList());
            //Insert read-date data
            worksheet.Cells["F2"].LoadFromCollection(notContainingValues.AsEnumerable().Select(x => x[readDateColumn].ToString()).ToList());
            //Insert read data
            worksheet.Cells["G2"].LoadFromCollection(notContainingValues.AsEnumerable().Select(x => x[readColumn].ToString()).ToList());
            //Insert faktor data
            worksheet.Cells["H2"].LoadFromCollection(notContainingValues.AsEnumerable().Select(x => x[faktorColumn].ToString()).ToList());
            //Insert reduction data
            worksheet.Cells["I2"].LoadFromCollection(notContainingValues.AsEnumerable().Select(x => x[reductionColumn].ToString()).ToList());
            //Insert room data
            worksheet.Cells["J2"].LoadFromCollection(notContainingValues.AsEnumerable().Select(x => x[roomColumn].ToString()).ToList());
            //Insert installation data
            worksheet.Cells["K2"].LoadFromCollection(notContainingValues.AsEnumerable().Select(x => x[installationDateColumn].ToString()).ToList());

            //Insert Selskab
            worksheet.Cells[2, 1, notContainingValues.Count + 1, 1].Value = department.ToString().Substring(0, 2);
            //Insert Afdeling
            worksheet.Cells[2, 2, notContainingValues.Count + 1, 2].Value = department.ToString().Substring(2, 2);
            //Insert Faktor
            //worksheet.Cells[2, 8, notContainingValues.Count + 1, 8].Value = factor;
            //Insert Målertype
            //worksheet.Cells[2, 4, notContainingValues.Count + 1, 4].Value = maalerType;

            package.SaveAs(newFile);



            package.Dispose();

        }


              

        public string getCompanyByDirectoryName(string _workingDirectory)
        {
            string company = string.Empty;
            //var companyString = _configuration.GetValue<string>("Company");

            var companyString = "1008 Casi;1902 Ista;3020 Minol;3920 Techem;4201 Brunata";

                var companyArray = companyString.Split(";");
                foreach (string c in companyArray)
                {
                    if (_workingDirectory.IndexOf(c) > 0)
                    {
                        company = c;
                        break;
                    }
                }
       
            return company;
        }

        public DataTable ImportExceltoDatatable(string path)
        {
            DataTable dt = new DataTable();

            string extension = Path.GetExtension(path);

            string excelVersion;

            switch (extension)
            {
                case ".xlsx":
                    excelVersion = "12.0";
                    dt = ReadExcelToDataTable(path, excelVersion);
                    break;
                case ".xls":
                    excelVersion = "8.0";
                    dt = ReadExcelToDataTable(path, excelVersion);
                    break;
                case ".csv":
                    List<ViewModelCSV> records = Load(path);
                    dt = ConvertToDatatable(records.ToList());
                    break;
                default:
                    Console.WriteLine("Fejl!");
                    break;
            }

            Console.WriteLine("Færdig, fik kørt en fil type af " + extension + " og der er " + dt.Rows.Count + " rækker i tabellen.");
            
            return dt;
            
        }

        public DataTable ReadExcelToDataTable(string path, string excelVersion)
        {
            string sqlQuery = "Select * From [Demo$]";
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel " + excelVersion + ";HDR=YES;\"";

            OleDbConnection con = new OleDbConnection(connectionString + "");
            OleDbDataAdapter da = new OleDbDataAdapter(sqlQuery, con);
            da.Fill(ds);

            dt = ds.Tables[0];

            //var meterIDs = dt.AsEnumerable().Select(r => r.Field<string>("Måler nr#")).ToList();

            return dt;
        }
                
        public List<Models.ViewModelCSV> Load(string _fileName)
        {
            FileInfo fileInfo = new FileInfo(_fileName);
            
            using var reader = new StreamReader(fileInfo.FullName, Encoding.GetEncoding("iso-8859-1"));
                using var csv = new CsvReader(reader, CultureInfo.CurrentCulture);
                csv.Configuration.HasHeaderRecord = false;
                csv.Configuration.MissingFieldFound = null;
                csv.Configuration.Delimiter = ";";

            List<ViewModelCSV> records = new List<ViewModelCSV>();

            records = csv.GetRecords<Models.ViewModelCSV>().ToList();
                    
            return records.ToList();

        }

        public DataTable ConvertToDatatable<T>(IList<T> data)
        {

            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
            {
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            }

            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                {
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                }

                table.Rows.Add(row);
            }

            return table;
        }

    }
}
