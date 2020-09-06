using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace KABService.Models
{
    class FactorModel
    {
        public int CompanyColumn { get; set; }
        public string CompanyID { get; set; }
        public int DepartmentColumn { get; set; }
        public string DepartmentID { get; set; }
        public int ApartmentColumn { get; set; }
        public int MaalerColumn { get; set; }
        public int MaalerTypeColumn { get; set; }
        public int SerieIDColumn { get; set; }
        public int ReadDateColumn { get; set; }
        public int ReadColumn { get; set; }
        public int FaktorColumn { get; set; }
        public int ReductionColumn { get; set; }
        public int RoomColumn { get; set; }
        public int InstallationDateColumn { get; set; }
        public string SearchCriteria { get; set; }
        public int SearchCriteriaColumn { get; set; }
        public int Factor { get; set; }
        public string MaalerType { get; set; }
        public string MaalerColumnName { get; set; }
        public string MaalerControlText { get; set; }
        public string Nustillingsmaaler { get; set; }
        public string ReadDate { get; set; }
        public string ReadDateFormatted { get; set; }
        public DataRow MaalerRow { get; set; }

        public FactorModel()
        {

        }
            
    }
}
