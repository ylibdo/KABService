using System;
using System.Collections.Generic;
using System.Text;

namespace KABService.Object
{
    static class BDOEnum
    {
        public enum FileMoveOption { Processed, Archive, Error, Manual}

        public enum LogLevel { Info, Warning, Error } 
    }
}