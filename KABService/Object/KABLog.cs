using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Text;

namespace KABService.Object
{
    class KABLog
    {
        public LogLevel LogLevel { get; set; }
        public string Log { get; set; }
        public DateTime CreatedDateTime { get; set; }
    }
}
