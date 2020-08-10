using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using KABWeb.Models;
using KABWeb.Helpers;
using Microsoft.Extensions.Configuration;
using System.IO;
using KABWeb.Helper;

namespace KABWeb.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IConfiguration _configuration;

        public HomeController(ILogger<HomeController> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;
        }

        public IActionResult Index()
        {
            List<KABDirectory> models = new List<KABDirectory>();
            int searchDepth = _configuration.GetValue<Int32>("SearchDepth");
            DirectioryHelper directioryHelper = new DirectioryHelper(_logger, _configuration);
            var diretoryList = directioryHelper.GetAllWorkingDirectoryFullPath();
            foreach(var directory in diretoryList)
            {
                KABDirectory model = new KABDirectory();
                model.SearchDepth = searchDepth;
                DirectoryInfo directoryInfo = new DirectoryInfo(directory);

                model.Name = directoryInfo.Name;
                model.SubDirectoryList = directoryInfo.EnumerateDirectories();
                model.FileList = directoryInfo.EnumerateFiles();
                models.Add(model);
            }
            return View(models);
        }

        public IActionResult Detail(string id)
        {
            try
            {

                FileInfo fileInfo = new FileInfo(id);
                if (fileInfo.Exists)
                {
                    if (fileInfo.Extension == ".csv")
                    {
                        // handling CSV files
                        CSVHelper csvHelper = new CSVHelper(_logger, _configuration);
                        IEnumerable<CommonViewModel> model = csvHelper.Load(fileInfo.FullName);
                        return View(model);
                    }
                    else if (fileInfo.Extension == ".xlsx")
                    {
                        // handling Excel files

                        // step 2.2 Prepare work sheet to read File
                        ExcelHelper excelHelper = new ExcelHelper(_logger, _configuration);
                        IEnumerable<CommonViewModel> model = excelHelper.Load(fileInfo.FullName);
                        return View(model);
                    }
                    else
                    {
                        return View("Error", new ErrorViewModel() { Message = "File format is not supported." });
                    }
                }
                else
                {
                    var message = String.Concat("File: ", fileInfo.FullName, " does no longer exist");
                    return View("Error", new ErrorViewModel() { Message = message });
                }
            }
            catch(Exception ex)
            {
                return View("Error", new ErrorViewModel() { Message = ex.Message });
            }
            
        }
        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
