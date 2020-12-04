using CsvHelper;
using KABService.Models;
using KABService.Object;
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

        public CSVHelper()
        {
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

        public string SaveDataToFile(IEnumerable<DataRow> _input, FactorModel _factorModel, string _company, string _workingDirectory, string _outputSuffix)
        {
            try
            {
                string newCSVFileName = string.Concat(_factorModel.CompanyID.PadLeft(3, Convert.ToChar("0")), _factorModel.DepartmentID.PadLeft(3, Convert.ToChar("0"))," ",_company, ConfigVariables.OutputFileSeparator, DateTime.Now.ToString(ConfigVariables.OutputFileDateFormatString), _outputSuffix, ConfigVariables.OutputFileSuffixCSV);
                FileInfo newCSVFile = new FileInfo(Path.Combine(_workingDirectory, newCSVFileName));
                List<OutputModelCSV> output = new List<OutputModelCSV>();
                foreach (DataRow row in _input)
                {
                    object[] rowArray = row.ItemArray;
                    OutputModelCSV model = new OutputModelCSV
                    {
                        Selskab = _factorModel.CompanyColumn < 0 ? _factorModel.CompanyID : Convert.ToString(rowArray[_factorModel.CompanyColumn]),
                        Afdeling = _factorModel.DepartmentColumn < 0 ? _factorModel.DepartmentID : Convert.ToString(rowArray[_factorModel.DepartmentColumn]),
                        Lejlighed = _factorModel.ApartmentColumn < 0 ? string.Empty : Convert.ToString(rowArray[_factorModel.ApartmentColumn]),
                        Målertype = _factorModel.MaalerColumn < 0 ? _factorModel.MaalerType : Convert.ToString(rowArray[_factorModel.MaalerColumn]),
                        Serienr = _factorModel.SerieIDColumn < 0 ? string.Empty : Convert.ToString(rowArray[_factorModel.SerieIDColumn]),
                        Aflæsningsdato = _factorModel.ReadDateColumn < 0 ? _factorModel.ReadDate : Convert.ToString(rowArray[_factorModel.ReadDateColumn]),
                        Aflæsning = _factorModel.ReadColumn < 0 ? string.Empty : Convert.ToString(rowArray[_factorModel.ReadColumn]),
                        Faktor = _factorModel.FaktorColumn < 0 ? _factorModel.Factor.ToString() : Convert.ToString(rowArray[_factorModel.FaktorColumn]),
                        Reduktion = _factorModel.ReductionColumn < 0 ? string.Empty : Convert.ToString(rowArray[_factorModel.ReductionColumn]),
                        Lokale = _factorModel.RoomColumn < 0 ? string.Empty : Convert.ToString(rowArray[_factorModel.RoomColumn]),
                        Installationsdato = _factorModel.InstallationDateColumn < 0 ? string.Empty : Convert.ToString(rowArray[_factorModel.InstallationDateColumn]),
                        Deaktiveringsdato = "",
                        Bemærkninger = "",
                        Nulstillingsmåler = _factorModel.Nustillingsmaaler
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
            catch(Exception ex)
            {
                return ex.Message;
            }
        }
    }
}
