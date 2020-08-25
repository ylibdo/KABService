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
            string newCSVFileName = string.Concat(_company, "_", DateTime.Now.ToString("yyyyMMddHHmmss"), "_unik.csv");
            FileInfo newFile = new FileInfo(Path.Combine(_workingDirectory, newFileName));
            FileInfo newCSVFile = new FileInfo(Path.Combine(_workingDirectory, newCSVFileName));

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            ExcelPackage package = new ExcelPackage(newFile);

            string regexPattern = @"\d{4}";

            Match department = Regex.Match(_company, regexPattern);

            string strDepartment = department.Value.ToString();

            Console.WriteLine(department.Value);

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

            
            DataTable dt = ImportExceltoDatatable(path);

            Console.WriteLine("dt rows count: " + dt.Rows.Count);



            //BusinessLogic

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
                    companyColumn = 0;
                    departmentColumn = 1;
                    apartmentColumn = 2;
                    maalerColumn = 4;
                    serieIDColumn = 5;
                    readDateColumn = 7;
                    readColumn = 8;
                    faktorColumn = 9;
                    reductionColumn = 10;
                    roomColumn = 11;
                    installationDateColumn = 12;
                    maalerType = "Doprimo 3 SoC";
                    notContainingValues = dt.AsEnumerable().Where(x => x[readDateColumn].Equals(searchCriteria) && x[maalerColumn].ToString().Contains(maalerType)).ToList();
                    break;

                case "3020":
                    //THERE IS 2 SHEETS WITH DATA WHICH COMBINES TO 1 OUTPUT FILE!
                    apartmentColumn = 1;
                    serieIDColumn = 4;
                    readDateColumn = 0;
                    readColumn = 3;
                    factor = 0.123;
                    maalerType = "M7R";
                    notContainingValues = dt.AsEnumerable().Where(x => !x[readColumn].ToString().Contains("FF")).ToList();
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

            List<OutputModelCSV> output = new List<OutputModelCSV>();
            foreach(DataRow row in notContainingValues)
            {
                object[] rowArray = row.ItemArray;
                OutputModelCSV model = new OutputModelCSV();
                model.Selskab = (string)rowArray[companyColumn];
                model.Afdeling = (string)rowArray[departmentColumn];
                model.Lejlighed = (string)rowArray[apartmentColumn];
                model.Målertype = (string)rowArray[maalerColumn];
                model.Serienr = (string)rowArray[serieIDColumn];
                model.Aflæsningsdato = (string)rowArray[readDateColumn];
                model.Aflæsning = (string)rowArray[readColumn];
                model.Faktor = (string)rowArray[faktorColumn];
                model.Reduktion = (string)rowArray[reductionColumn];
                model.Lokale = (string)rowArray[roomColumn];
                model.Installationsdato = (string)rowArray[installationDateColumn];
                model.Deaktiveringsdato = "";
                model.Bemærkninger = "";
                model.Nulstillingsmåler = "";

                output.Add(model);
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


            var fStream = new FileStream(newCSVFile.ToString(), FileMode.Create);
            using (var writer = new StreamWriter(fStream, Encoding.GetEncoding("ISO-8859-1")))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.Configuration.Delimiter = ";";
                csv.WriteRecords(output);
                csv.Flush();
            }

            package.SaveAs(newFile);

            package.Dispose();

            Console.WriteLine("Done!!");
            Console.ReadLine();
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
                    List<InputModelCSV> records = Load(path);
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
                
        public List<Models.InputModelCSV> Load(string _fileName)
        {
            FileInfo fileInfo = new FileInfo(_fileName);
            
            using var reader = new StreamReader(fileInfo.FullName, Encoding.GetEncoding("iso-8859-1"));
                using var csv = new CsvReader(reader, CultureInfo.CurrentCulture);
                csv.Configuration.HasHeaderRecord = false;
                csv.Configuration.MissingFieldFound = null;
                csv.Configuration.Delimiter = ";";

            List<InputModelCSV> records = new List<InputModelCSV>();

            records = csv.GetRecords<Models.InputModelCSV>().ToList();
                    
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
