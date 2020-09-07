using CsvHelper;
using KABService.Models;
using KABService.Object;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;


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

            FileInfo fileInfo = new FileInfo(path);
            string newFileName = string.Empty;
            string newErrorFileName = string.Empty;
            DataTable outputDataTable = new DataTable();
            DataTable unikDataTable = new DataTable();
            string excelVersion = string.Empty;
            DirectioryHelper directioryHelper = new DirectioryHelper(_logger, _configuration);


            // 3. create factor object based on company
            var company = getCompanyByDirectoryName(_workingDirectory);

            
            // 4. Reading from files.
            CSVHelper csvHelper = new CSVHelper(_logger, _configuration);
            ExcelHelper excelHelper = new ExcelHelper(_logger, _configuration);

            // 4.1 Read Unik data

            unikDataTable = excelHelper.ReadDataAsDataTable("E:\\KAB Services\\MeterService\\UnikData\\UnikData.xlsx", "12.0");

            if (fileInfo.Extension == ".csv")

            {
                // handling CSV files

                outputDataTable = csvHelper.ReadDataAsDataTable(path);
                outputDataTable.Rows[0].Delete();
                outputDataTable.AcceptChanges();
            }
            else if (fileInfo.Extension == ".xlsx" || fileInfo.Extension == ".xlsb")
            {
                // handling Excel files
                excelVersion = "12.0";
                outputDataTable = excelHelper.ReadDataAsDataTable(path, excelVersion);
            }
            else if (fileInfo.Extension == ".xls")
            {
                // handling Excel files
                excelVersion = "8.0";
                outputDataTable = excelHelper.ReadDataAsDataTable(path, excelVersion);
            }
            else
            {
                throw new FileLoadException("File format is not supported");
            }

            FactorModel factorModel = BusinessLogic.CreateFactorModelByCompany(company, unikDataTable, outputDataTable);

            if (factorModel == null)
            {
                throw new Exception("Vendor configuration error. Cannot create factor model based based on vendor information.");
            }

            var tester = Convert.ToDateTime(outputDataTable.Rows[0].ItemArray[7]).ToString("dd-MM-yyyy");


            try
            {
                // 5. Filter source data.
                //IEnumerable<DataRow> filteredData = BusinessLogic.FilterDataByCompany(outputDataTable, factorModel, company);
                IEnumerable<DataRow> filteredData = BusinessLogic.FilterDataByCompany(outputDataTable, factorModel, company);

                IEnumerable<DataRow> errorData = BusinessLogic.ErrorDataByCompany(outputDataTable, factorModel, company);

                // 6. Save data to file.
                newFileName = csvHelper.SaveDataToFile(filteredData, factorModel, company, _workingDirectory, ConfigVariables.OutputFileNameSuffix);

                newErrorFileName = csvHelper.SaveDataToFile(errorData, factorModel, company, _workingDirectory, ConfigVariables.ErrorFileNameSuffix);
            }
            catch(Exception ex)
            {
                _logger.LogError(ex.Message);
            }


             if (String.IsNullOrEmpty(newFileName)) 
            { 

                throw new NullReferenceException();
            }
            else
            {
                directioryHelper.MoveFile(_workingDirectory, newFileName, BDOEnum.FileMoveOption.Processed);
            }

            if (String.IsNullOrEmpty(newErrorFileName))
            {
                throw new NullReferenceException();
            }
            else
            {
                directioryHelper.MoveFile(_workingDirectory, newErrorFileName, BDOEnum.FileMoveOption.Error);
            }

            // 7. Move processed file to archive
            directioryHelper.MoveFile(_workingDirectory, path, BDOEnum.FileMoveOption.Archive);


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
