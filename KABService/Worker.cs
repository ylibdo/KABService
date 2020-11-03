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
                int workerRuntimeIntervalByMillesecond = _configuration.GetValue<int>("WorkerRuntimeIntervalByMinute") * 60 * 1000;
                if(workerRuntimeIntervalByMillesecond == 0)
                {
                    throw new NullReferenceException("Configuration is missing for WorkerRuntimeIntervalByMinute.");
                }
                await Task.Delay(30000, stoppingToken);
            }
        }

    }
}
