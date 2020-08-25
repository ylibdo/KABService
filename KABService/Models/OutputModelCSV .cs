using CsvHelper.Configuration.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace KABService.Models
{
    
    public class OutputModelCSV
    {
        [Index(0)]
        public string Selskab { get; set; }
        [Index(1)]
        public string Afdeling { get; set; }
        [Index(2)]
        public string Lejlighed { get; set; }
        [Index(3)]
        public string Målertype { get; set; }
        [Index(4)]
        public string Serienr { get; set; }
        [Index(5)]
        public string Aflæsningsdato { get; set; }
        [Index(6)]
        public string Aflæsning { get; set; }
        [Index(7)]
        public string Faktor { get; set; }
        [Index(8)]
        public string Reduktion { get; set; }
        [Index(9)]
        public string Lokale { get; set; }
        [Index(10)]
        public string Installationsdato { get; set; }
        [Index(11)]
        public string Deaktiveringsdato { get; set; }
        [Index(12)]
        public string Bemærkninger { get; set; }
        [Index(13)]
        public string Nulstillingsmåler { get; set; }
        /*[Index(14)]
        public string Field15 { get; set; }*/
    }
}
