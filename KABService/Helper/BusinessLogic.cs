using KABService.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System.Linq;


namespace KABService.Helper
{
    class BusinessLogic
    {

               
        public static FactorModel CreateFactorModelByCompany(string _company, DataTable _unikDatatable, DataTable _outputDatatable)
        
        {
            string regexPattern = @"\d{4}";
            
            switch (_company)
            {
                case "1008 Casi":

                    FactorModel casi = new FactorModel();
                    casi.CompanyColumn = -1;
                    casi.DepartmentColumn = -1;
                    casi.ApartmentColumn = -1;
                    casi.MaalerColumn = -1;
                    casi.SerieIDColumn = 2;
                    casi.ReadDateColumn = 4;
                    casi.ReadColumn = 5;
                    casi.FaktorColumn = -1;
                    casi.Factor = 1;
                    casi.ReductionColumn = -1;
                    casi.RoomColumn = -1;
                    casi.InstallationDateColumn = -1;
                    //casi.CompanyID = Regex.Match(_company, regexPattern).Value.ToString().Substring(0, 2);
                    //casi.DepartmentID = Regex.Match(_company, regexPattern).Value.ToString().Substring(2, 2);
                    casi.SearchCriteria = "s";
                    casi.CompanyID = "10";
                    casi.DepartmentID = "8";

                    casi.MaalerColumnName = _outputDatatable.Columns[3].ToString();
                    casi.MaalerControlText = _outputDatatable.Select(casi.MaalerColumnName + " IS NOT NULL AND " + casi.MaalerColumnName + " <> ''").First().ItemArray[3].ToString();
                    casi.MaalerType = _unikDatatable.AsEnumerable().Where(x => x[4].ToString().Equals(casi.MaalerControlText) && x[0].ToString().Equals(casi.CompanyID) && x[1].ToString().Equals(casi.DepartmentID.Replace("0",""))).First().ItemArray[2].ToString();
                    casi.Nustillingsmaaler = _unikDatatable.AsEnumerable().Where(x => x[4].ToString().Equals(casi.MaalerControlText) && x[0].ToString().Equals(casi.CompanyID) && x[1].ToString().Equals(casi.DepartmentID.Replace("0", ""))).First().ItemArray[3].ToString();
                    casi.ReadDate = _unikDatatable.AsEnumerable().Where(x => x[4].ToString().Equals(casi.MaalerControlText) && x[0].ToString().Equals(casi.CompanyID) && x[1].ToString().Equals(casi.DepartmentID.Replace("0", ""))).First().ItemArray[5].ToString();

                    return casi;
                    
                case "1902 Ista":
                    FactorModel ista = new FactorModel();
                    ista.CompanyColumn = 0;
                    ista.DepartmentColumn = 1;
                    ista.ApartmentColumn = 2;
                    ista.MaalerColumn = 4;
                    ista.SerieIDColumn = 5;
                    ista.ReadDateColumn = 7;
                    ista.ReadColumn = 8;
                    ista.FaktorColumn = 9;
                    ista.ReductionColumn = 10;
                    ista.RoomColumn = 11;
                    ista.InstallationDateColumn = 12;
                    ista.CompanyID = _outputDatatable.Rows[0].ItemArray[0].ToString();
                    ista.CompanyID = ista.CompanyID.Length >= 4 ? ista.CompanyID.Substring(0, 2) : ista.CompanyID;
                    ista.DepartmentID = ista.CompanyID.Length >= 4 ? ista.CompanyID.Substring(2, 2) : _outputDatatable.Rows[0].ItemArray[1].ToString();

                    ista.MaalerType = _unikDatatable.AsEnumerable().Where(x => x[0].ToString().Equals(ista.CompanyID) && x[1].ToString().Equals(ista.DepartmentID.Replace("0", ""))).First().ItemArray[2].ToString();
                    ista.Nustillingsmaaler = _unikDatatable.AsEnumerable().Where(x => x[0].ToString().Equals(ista.CompanyID) && x[1].ToString().Equals(ista.DepartmentID.Replace("0", ""))).First().ItemArray[3].ToString();
                    ista.ReadDate = _unikDatatable.AsEnumerable().Where(x => x[0].ToString().Equals(ista.CompanyID) && x[1].ToString().Equals(ista.DepartmentID.Replace("0", ""))).First().ItemArray[5].ToString();
                    ista.ReadDate = "2019-12-31";

                    return ista;

                case "3020 Minol":
                    FactorModel minol = new FactorModel();

                    minol.CompanyColumn = -1;
                    minol.DepartmentColumn = -1;
                    minol.ApartmentColumn = 0;
                    minol.MaalerColumn = -1;
                    minol.SerieIDColumn = 6;
                    minol.ReadDateColumn = -1;
                    minol.ReadColumn = 5;
                    minol.FaktorColumn = 4;
                    minol.ReductionColumn = -1;
                    minol.RoomColumn = -1;
                    minol.InstallationDateColumn = -1;
                    minol.SearchCriteriaColumn = 3;
                    minol.CompanyID = Regex.Match(_company, regexPattern).Value.ToString().Substring(0, 2);
                    minol.DepartmentID = Regex.Match(_company, regexPattern).Value.ToString().Substring(2, 2);

                    minol.MaalerType = _unikDatatable.AsEnumerable().Where(x => x[0].ToString().Equals(minol.CompanyID) && x[1].ToString().Equals(minol.DepartmentID)).First().ItemArray[2].ToString();
                    minol.Nustillingsmaaler = _unikDatatable.AsEnumerable().Where(x => x[0].ToString().Equals(minol.CompanyID) && x[1].ToString().Equals(minol.DepartmentID)).First().ItemArray[3].ToString();
                    minol.ReadDate = _unikDatatable.AsEnumerable().Where(x => x[0].ToString().Equals(minol.CompanyID) && x[1].ToString().Equals(minol.DepartmentID)).First().ItemArray[5].ToString();

                    return minol;

                case "3920 Techem":
                    FactorModel techem = new FactorModel();

                    techem.CompanyColumn = -1;
                    techem.DepartmentColumn = -1;
                    techem.ApartmentColumn = 0;
                    techem.MaalerColumn = 4;
                    techem.SerieIDColumn = 6;
                    techem.ReadDateColumn = -1;
                    techem.ReadColumn = 8;
                    techem.FaktorColumn = -1;
                    techem.ReductionColumn = -1;
                    techem.RoomColumn = -1;
                    techem.InstallationDateColumn = -1;
                    techem.CompanyID = Regex.Match(_company, regexPattern).Value.ToString().Substring(0, 2);
                    techem.DepartmentID = Regex.Match(_company, regexPattern).Value.ToString().Substring(2, 2);
                    techem.SearchCriteria = "Energimåler";

                    

                    techem.MaalerType = _unikDatatable.AsEnumerable().Where(x => x[0].ToString().Equals(techem.CompanyID) && x[1].ToString().Equals(techem.DepartmentID)).First().ItemArray[2].ToString();
                    techem.Nustillingsmaaler = _unikDatatable.AsEnumerable().Where(x => x[0].ToString().Equals(techem.CompanyID) && x[1].ToString().Equals(techem.DepartmentID)).First().ItemArray[3].ToString();
                    techem.ReadDate = _unikDatatable.AsEnumerable().Where(x => x[0].ToString().Equals(techem.CompanyID) && x[1].ToString().Equals(techem.DepartmentID)).First().ItemArray[5].ToString();

                    //Temp value
                    techem.ReadDate = "31-12-2019";

                    techem.ReadColumn = _outputDatatable.Columns.IndexOf(techem.ReadDate);

                    return techem;

                case "4201 Brunata":
                    FactorModel brunata = new FactorModel();

                    brunata.CompanyColumn = -1;
                    brunata.DepartmentColumn = -1;
                    brunata.ApartmentColumn = 2;
                    brunata.MaalerColumn = -1;
                    brunata.SerieIDColumn = 5;
                    brunata.ReadDateColumn = 8;
                    brunata.ReadColumn = 7;
                    brunata.SearchCriteriaColumn = 10;
                    brunata.FaktorColumn = -1;
                    brunata.Factor = 1;
                    brunata.ReductionColumn = -1;
                    brunata.RoomColumn = -1;
                    brunata.InstallationDateColumn = -1;
                    brunata.SearchCriteria = "[ABN]";
                    brunata.CompanyID = _outputDatatable.Rows[0].ItemArray[0].ToString();
                    brunata.DepartmentID = brunata.CompanyID.Length >= 4 ? brunata.CompanyID.Substring(2, 2) : _outputDatatable.Rows[0].ItemArray[1].ToString();
                    brunata.CompanyID = brunata.CompanyID.Length >= 4 ? brunata.CompanyID.Substring(0, 2) : brunata.CompanyID;
                    
                    brunata.MaalerType = _unikDatatable.AsEnumerable().Where(x => x[0].ToString().Equals(brunata.CompanyID) && x[1].ToString().Equals(brunata.DepartmentID.Replace("0", ""))).First().ItemArray[2].ToString();
                    brunata.Nustillingsmaaler = _unikDatatable.AsEnumerable().Where(x => x[0].ToString().Equals(brunata.CompanyID) && x[1].ToString().Equals(brunata.DepartmentID.Replace("0", ""))).First().ItemArray[3].ToString();
                    brunata.ReadDate = _unikDatatable.AsEnumerable().Where(x => x[0].ToString().Equals(brunata.CompanyID) && x[1].ToString().Equals(brunata.DepartmentID.Replace("0", ""))).First().ItemArray[5].ToString();
                    //Temp condition - Unik data is 2020 year and csv file is 2019
                    brunata.ReadDate = "31-12-2019";
                    brunata.ReadDateFormatted = DateTime.Parse(brunata.ReadDate).ToString("yyyyMMdd");
                    
                    
                    return brunata;

                default:
                    
                    throw new Exception("Unknow company/vendor name is found.");

            }
        }

        public static IEnumerable<DataRow> FilterDataByCompany(DataTable _input, FactorModel _factorModel, string _company)
        {
            switch (_company)
            {
                case "1008 Casi":
                    return _input.AsEnumerable().Where(x => !x[_factorModel.ReadColumn].ToString().Contains(_factorModel.SearchCriteria));
                case "1902 Ista":
                    return _input.AsEnumerable().Where(x => x[_factorModel.ReadDateColumn].Equals(_factorModel.ReadDate) && x[_factorModel.MaalerColumn].ToString().Contains(_factorModel.MaalerType));
                case "3020 Minol":
                    return _input.AsEnumerable().Where(x => !string.IsNullOrEmpty(x[_factorModel.ReadColumn].ToString()) && string.IsNullOrWhiteSpace(x[_factorModel.SearchCriteriaColumn].ToString()));
                case "3920 Techem":
                    IEnumerable<DataRow> filteredValueTechem = _input.AsEnumerable().Where(x => !x[_factorModel.MaalerColumn].ToString().Contains(_factorModel.SearchCriteria));
                    List<DataRow> outputTechem = new List<DataRow>();
                    foreach (DataRow item in filteredValueTechem.ToList())
                    {
                        item[_factorModel.MaalerColumn] = item[_factorModel.MaalerColumn].ToString().Replace("vandsmåler", "t vand");
                        item[_factorModel.ApartmentColumn] = item[_factorModel.ApartmentColumn].ToString().Substring(0, item[_factorModel.ApartmentColumn].ToString().Length - 1);
                        outputTechem.Add(item);
                    }
                    return outputTechem;
                case "4201 Brunata":
                    IEnumerable<DataRow> filteredValueBrunata = _input.AsEnumerable().Where(x => Regex.Match(x[_factorModel.SearchCriteriaColumn].ToString().ToUpper(), _factorModel.SearchCriteria).Success && x[_factorModel.ReadDateColumn].ToString().Equals(_factorModel.ReadDateFormatted));
                    List<DataRow> outputBrunata = new List<DataRow>();
                    foreach (DataRow item in filteredValueBrunata.ToList())
                    {
                        item[_factorModel.ReadDateColumn] = _factorModel.ReadDate;
                        outputBrunata.Add(item);
                    }
                    return outputBrunata;
                default:
                    throw new Exception("Unknow company/vendor name is found.");
            }
        }

        public static IEnumerable<DataRow> ErrorDataByCompany(DataTable _input, FactorModel _factorModel, string _company)
        {
            switch (_company)
            {
                case "1008 Casi":
                    return _input.AsEnumerable().Where(x => x[_factorModel.ReadColumn].ToString().Contains(_factorModel.SearchCriteria));
                case "1902 Ista":
                    return _input.AsEnumerable().Where(x => !x[_factorModel.ReadDateColumn].Equals(_factorModel.SearchCriteria) && !x[_factorModel.MaalerColumn].ToString().Contains(_factorModel.MaalerType));
                case "3020 Minol":
                    return _input.AsEnumerable().Where(x => string.IsNullOrEmpty(x[_factorModel.ReadColumn].ToString()) || !string.IsNullOrWhiteSpace(x[_factorModel.SearchCriteriaColumn].ToString()));
                case "3920 Techem":
                    return _input.AsEnumerable().Where(x => x[_factorModel.MaalerColumn].ToString().Contains(_factorModel.SearchCriteria));
                case "4201 Brunata":
                    IEnumerable<DataRow> filteredValueBrunata = _input.AsEnumerable().Where(x => !Regex.Match(x[_factorModel.SearchCriteriaColumn].ToString().ToUpper(), _factorModel.SearchCriteria).Success);
                    List<DataRow> outputBrunata = new List<DataRow>();
                    foreach (DataRow item in filteredValueBrunata.ToList())
                    {
                        item[_factorModel.ReadDateColumn] = _factorModel.ReadDate;
                        outputBrunata.Add(item);
                    }
                    return outputBrunata;
                default:
                    throw new Exception("Unknow company/vendor name is found.");
            }
        }
    }
}
