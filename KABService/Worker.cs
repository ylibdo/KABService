using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using KABService.Helper;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace KABService
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;
        private readonly IConfiguration _configuration;

        public Worker(ILogger<Worker> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                _logger.LogInformation("Meter service is running at: {time}", DateTimeOffset.Now);
                MeterService meterService = new MeterService(_logger, _configuration);
                meterService.Run();
                //ExcelHelper excelhelper = new ExcelHelper(_logger, _configuration);

                //excelhelper.ReadDataAsDataTable2("E:\\KAB Services\\MeterService\\1008 Casi\\KAB-AKB_Afd-_10-08_Karr_8 311219 visninger - Copy.xlsx", "12.0");


                //ExcelTest excelTest = new ExcelTest(_logger, _configuration);

                //excelTest.ImportExceltoDatatable("E:\\KAB Services\\MeterService\\1008 Casi\\KAB-AKB_Afd-_10-08_Karr_8 311219 visninger.xlsx");

                //excelTest.ImportExceltoDatatable("E:\\KAB Services\\MeterService\\4201 Brunata\\Payroll.csv");

                //excelTest.ProcessExcelFile("E:\\KAB Services\\MeterService\\1902 Ista\\740582_2020-02-14 133808_KAB_2019-12-31_Version_200.csv", "E:\\KAB Services\\MeterService\\1902 Ista");
                
                //ExcelHelper excelHelper = new ExcelHelper(_logger,_configuration);
                //excelHelper.Process(excelHelper.Prepare(@"E:\KAB Services\MeterService\1008 Casi\KAB - AKB_Afd - _10 - 08_Karr_8 311219 visninger.xlsx"), @"E:\KAB Services\MeterService\1008 Casi");
                await Task.Delay(30000, stoppingToken);
            }
        }

    }
}
