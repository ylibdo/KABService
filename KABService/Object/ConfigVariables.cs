using System;
using System.Collections.Generic;
using System.Text;

namespace KABService.Object
{
    public static class ConfigVariables
    {
        public const string OutputFileSeparator = "_";
        public const string OutputFileSuffixCSV = ".csv";
        public const string OutputFileSuffixExcel = ".xlsx";
        public const string OutputFileDateFormatString = "yyyyMMddHHmmssFFF";
        public const string OutputFileNameSuffix = "_indlæsningsfil";
        public const string ErrorFileNameSuffix = "_manuelbehandling";


        //Insert headers
        public const string Company = "Selskab";
        public const string Department = "Afdeling";
        public const string Apartment = "Lejlighed";
        public const string MeterType = "Målertype";
        public const string SerieID = "Serienr";
        public const string ReadingDate = "Aflæsningsdato";
        public const string Reading = "Aflæsning";
        public const string Factor = "Faktor";
        public const string Reduction = "Reduktion";
        public const string Room = "Lokale";
        public const string Installation = "Installation";
        public const string DeactivationDate = "Deaktiveringsdato";
        public const string Comment = "Bemærkning";
        public const string ResetMeter = "Nulstillingsmåler";


        // log
        public const string LogDirectoryName = "Log";
        public const string SubDirectoryNameDateFormat = "yyyyMMdd";
        public const string LogFileNameDateFormat = "yyyyMMddHHmm";
        public const string LogFileName = "KABlogfile_";

    }
}
