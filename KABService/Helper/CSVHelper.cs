using CsvHelper;
using KABService.Models;
using KABService.Object;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
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
        public DataTable ReadDataAsDataTable(string _fileName)
        {
            FileInfo fileInfo = new FileInfo(_fileName);

            using var reader = new StreamReader(fileInfo.FullName, Encoding.GetEncoding("iso-8859-1"));
            using var csv = new CsvReader(reader, CultureInfo.CurrentCulture);
            csv.Configuration.HasHeaderRecord = false;
            csv.Configuration.MissingFieldFound = null;
            csv.Configuration.Delimiter = ";";

            List<InputModelCSV> records = new List<InputModelCSV>();

            records = csv.GetRecords<InputModelCSV>().ToList();

            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(InputModelCSV));
            DataTable table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
            {
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            }

            foreach (InputModelCSV item in records)
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

        public string SaveDataToFile(IEnumerable<DataRow> _input, string _company, string _workingDirectory)
        {
            try
            {
                string newCSVFileName = string.Concat(_company, ConfigVariables.OutputFileSeparator, DateTime.Now.ToString(ConfigVariables.OutputFileDateFormatString), ConfigVariables.OutputFileSuffixCSV);
                FileInfo newCSVFile = new FileInfo(Path.Combine(_workingDirectory, newCSVFileName));
                List<OutputModelCSV> output = new List<OutputModelCSV>();
                foreach (DataRow row in _input)
                {
                    object[] rowArray = row.ItemArray;
                    OutputModelCSV model = new OutputModelCSV
                    {
                        Selskab = (string)rowArray[0],
                        Afdeling = (string)rowArray[1],
                        Lejlighed = (string)rowArray[2],
                        Målertype = (string)rowArray[4],
                        Serienr = (string)rowArray[5],
                        Aflæsningsdato = (string)rowArray[7],
                        Aflæsning = (string)rowArray[8],
                        Faktor = (string)rowArray[9],
                        Reduktion = (string)rowArray[10],
                        Lokale = (string)rowArray[11],
                        Installationsdato = (string)rowArray[12],
                        Deaktiveringsdato = "",
                        Bemærkninger = "",
                        Nulstillingsmåler = ""
                    };

                    output.Add(model);
                }

                var fStream = new FileStream(newCSVFile.FullName, FileMode.Create);
                using var writer = new StreamWriter(fStream, Encoding.GetEncoding("ISO-8859-1"));
                using var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);
                csv.Configuration.Delimiter = ";";
                csv.WriteRecords(output);
                csv.Flush();
                return newCSVFile.FullName;
            }
            catch(Exception)
            {
                return string.Empty;
            }
        }
    }
}
