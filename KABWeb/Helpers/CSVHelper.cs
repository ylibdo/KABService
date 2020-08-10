using CsvHelper;
using KABWeb.Controllers;
using KABWeb.Models;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KABWeb.Helpers
{
    public class CSVHelper
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IConfiguration _configuration;

        public CSVHelper(ILogger<HomeController> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;
        }
        public List<CommonViewModel> Load(string _fileName)
        {
            FileInfo fileInfo = new FileInfo(_fileName);
            try
            {
                using var reader = new StreamReader(fileInfo.FullName, Encoding.GetEncoding("iso-8859-1"));
                using var csv = new CsvReader(reader, CultureInfo.CurrentCulture);
                csv.Configuration.HasHeaderRecord = false;
                csv.Configuration.MissingFieldFound = null;
                csv.Configuration.Delimiter = ";";
                var records = csv.GetRecords<CommonViewModel>();
                return records.ToList();

            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message);
                return null;
            }
        }
    }
}
