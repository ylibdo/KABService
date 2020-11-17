using KABService.Models;
using KABService.Object;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace KABService.Helper
{
    class BusinessLogic
    {       
        public static FactorModel CreateFactorModelByCompany(string _company, DataTable _unikDatatable, DataTable _outputDatatable)
        
        {
            string regexPattern = @"\d{4}";
            
            //int unikMaalerColumn = 2;
            int unikMaalerTypeColumn = 4;
            int unikCompanyColumn = 7;
            int unikDepartmentColumn = 8;
            //int unikMaalerCountColumn = 13;
            //int unikDateColumn = 14;

            switch (_company)
            {
                case "Casi":

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
                    //casi.DepartmentID = Convert.ToString(Convert.ToInt32(casi.DepartmentID));
                    casi.SearchCriteria = "s";
                    casi.CompanyID = "10";
                    casi.DepartmentID = "08";

                    casi.MaalerColumnName = _outputDatatable.Columns[3].ToString();
                    casi.MaalerControlText = _outputDatatable.Select(casi.MaalerColumnName + " IS NOT NULL AND " + casi.MaalerColumnName + " <> ''").First().ItemArray[3].ToString();
                    
                    casi.MaalerRow = _unikDatatable.AsEnumerable().FirstOrDefault(x => x[unikMaalerTypeColumn].ToString().Equals(casi.MaalerControlText) && x[unikCompanyColumn].ToString().Equals(casi.CompanyID) && x[unikDepartmentColumn].ToString().Equals(casi.DepartmentID));

                    if (casi.MaalerRow != null)
                    {
                        casi.MaalerType = casi.MaalerRow.ItemArray[2].ToString();
                        casi.Nustillingsmaaler = casi.MaalerRow.ItemArray[3].ToString();
                        casi.ReadDate = casi.MaalerRow.ItemArray[5].ToString();
                        casi.ReadDate = casi.MaalerRow.ItemArray[14].ToString();
                        //casi.ReadDate = "2019-12-31";
                    }

                    return casi;
                    
                case "Ista":
                    FactorModel ista = new FactorModel();
                    ista.CompanyColumn = 0;
                    ista.DepartmentColumn = 1;
                    ista.ApartmentColumn = 2;
                    ista.MaalerTypeColumn = 3;
                    ista.MaalerColumn = 4;
                    ista.SerieIDColumn = 5;
                    ista.ReadDateColumn = 7;
                    ista.ReadColumn = 8;
                    ista.FaktorColumn = 9;
                    ista.ReductionColumn = -1;
                    ista.RoomColumn = 11;
                    ista.InstallationDateColumn = 12;
                    ista.CompanyID = _outputDatatable.Rows[0].ItemArray[0].ToString();
                    ista.DepartmentID = ista.CompanyID.Length >= 4 ? ista.CompanyID.Substring(2, 2) : _outputDatatable.Rows[0].ItemArray[1].ToString();
                    ista.CompanyID = ista.CompanyID.Length >= 4 ? ista.CompanyID.Substring(0, 2) : ista.CompanyID;
                    ista.DepartmentID = Convert.ToString(Convert.ToInt32(ista.DepartmentID));

                    var listTest = BDOEnum.GetMaalerDescription(_outputDatatable.Rows[1].ItemArray[ista.MaalerTypeColumn].ToString());



                    ista.MaalerRow = _unikDatatable.AsEnumerable().FirstOrDefault(x => x[unikCompanyColumn].ToString().Equals(ista.CompanyID) && x[unikDepartmentColumn].ToString().Equals(ista.DepartmentID));

                    var tester123 = _unikDatatable.AsEnumerable().Where(x => x[unikCompanyColumn].ToString().Equals(ista.CompanyID) && x[unikDepartmentColumn].ToString().Equals(ista.DepartmentID));

                    if (ista.MaalerRow != null)
                    {
                        ista.MaalerType = ista.MaalerRow.ItemArray[2].ToString().Split(" ")[0];
                        ista.Nustillingsmaaler = ista.MaalerRow.ItemArray[3].ToString();
                        ista.ReadDate = ista.MaalerRow.ItemArray[5].ToString();
                        ista.ReadDate = ista.MaalerRow.ItemArray[14].ToString();
                        //ista.ReadDate = "2019-12-31";
                    }
                                        

                    return ista;

                case "Minol":
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
                    //minol.CompanyID = Regex.Match(_company, regexPattern).Value.ToString().Substring(0, 2);
                    //minol.DepartmentID = Regex.Match(_company, regexPattern).Value.ToString().Substring(2, 2);
                    minol.CompanyID = "30";
                    minol.DepartmentID = "20";
                    //minol.CompanyID = _outputDatatable.Rows[0].ItemArray[0].ToString();
                    //minol.DepartmentID = minol.CompanyID.Length >= 4 ? minol.CompanyID.Substring(2, 2) : _outputDatatable.Rows[0].ItemArray[1].ToString();
                    //minol.CompanyID = minol.CompanyID.Length >= 4 ? minol.CompanyID.Substring(0, 2) : minol.CompanyID;
                    //minol.DepartmentID = Convert.ToString(Convert.ToInt32(minol.DepartmentID));

                    minol.MaalerRow = _unikDatatable.AsEnumerable().FirstOrDefault(x => x[unikCompanyColumn].ToString().Equals(minol.CompanyID) && x[unikDepartmentColumn].ToString().Equals(minol.DepartmentID));

                    if (minol.MaalerRow != null)
                    {
                        minol.MaalerType = minol.MaalerRow.ItemArray[2].ToString();
                        minol.Nustillingsmaaler = minol.MaalerRow.ItemArray[3].ToString();
                        minol.ReadDate = minol.MaalerRow.ItemArray[5].ToString();
                        minol.ReadDate = Convert.ToDateTime(minol.MaalerRow.ItemArray[14].ToString()).ToString("dd-MM-yyyy");
                    }

                    return minol;

                case "Techem":
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
                    //techem.CompanyID = Regex.Match(_company, regexPattern).Value.ToString().Substring(0, 2);
                    //techem.DepartmentID = Regex.Match(_company, regexPattern).Value.ToString().Substring(2, 2);
                    //techem.DepartmentID = Convert.ToString(Convert.ToInt32(techem.DepartmentID));
                    techem.SearchCriteria = "Energimåler";
                    techem.CompanyID = "39";
                    techem.DepartmentID = "20";

                    techem.MaalerRow = _unikDatatable.AsEnumerable().FirstOrDefault(x => x[unikCompanyColumn].ToString().Equals(techem.CompanyID) && x[unikDepartmentColumn].ToString().Equals(techem.DepartmentID));

                    if (techem.MaalerRow != null)
                    {
                        techem.MaalerType = techem.MaalerRow.ItemArray[2].ToString();
                        techem.Nustillingsmaaler = techem.MaalerRow.ItemArray[3].ToString();
                        techem.ReadDate = techem.MaalerRow.ItemArray[5].ToString();
                        techem.ReadDate = Convert.ToDateTime(techem.MaalerRow.ItemArray[14].ToString()).ToString("dd-MM-yyyy");
                    }

                    //Temp value
                    //techem.ReadDate = "31-12-2019";

                    techem.ReadColumn = _outputDatatable.Columns.IndexOf(techem.ReadDate);

                    return techem;

                case "Brunata":
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
                    brunata.DepartmentID = Convert.ToString(Convert.ToInt32(brunata.DepartmentID));

                    brunata.MaalerRow = _unikDatatable.AsEnumerable().FirstOrDefault(x => x[unikCompanyColumn].ToString().Equals(brunata.CompanyID) && x[unikDepartmentColumn].ToString().Equals(brunata.DepartmentID));

                    if (brunata.MaalerRow != null)
                    {
                        brunata.MaalerType = brunata.MaalerRow.ItemArray[2].ToString();
                        brunata.Nustillingsmaaler = brunata.MaalerRow.ItemArray[3].ToString();
                        brunata.ReadDate = brunata.MaalerRow.ItemArray[5].ToString();
                        brunata.ReadDate = brunata.MaalerRow.ItemArray[14].ToString();
                        brunata.ReadDateFormatted = DateTime.Parse(brunata.ReadDate).ToString("yyyyMMdd");
                    }
                    else
                    {
                        throw new Exception("Cannot find måler type + date in Unik");
                    }
                    
                    return brunata;

                default:
                    return null;
            }
        }

        public static IEnumerable<DataRow> FilterDataByCompany(DataTable _input, FactorModel _factorModel, string _company)
        {

            DataTable dtCloned = _input.Clone();


            if (_factorModel.ReadDateColumn != -1)
            {
                dtCloned.Columns[_factorModel.ReadDateColumn].DataType = typeof(string);
            }

            if (_factorModel.InstallationDateColumn != -1)
            {
                dtCloned.Columns[_factorModel.InstallationDateColumn].DataType = typeof(string);
            }

            foreach (DataRow row in _input.Rows)
            {

                dtCloned.ImportRow(row);
            }

            switch (_company)
            {
                case "Casi":
                    return dtCloned.AsEnumerable().Where(x => !x[_factorModel.ReadColumn].ToString().Contains(_factorModel.SearchCriteria));
                case "Ista":
                    IEnumerable<DataRow> filteredValueIsta = dtCloned.AsEnumerable().Where(x => Convert.ToDateTime(x[_factorModel.ReadDateColumn].ToString()).ToString("dd-MM-yyyy").Contains(Convert.ToDateTime(_factorModel.ReadDate.ToString()).ToString("dd-MM-yyyy")) && x[_factorModel.MaalerColumn].ToString().ToLower().Contains(_factorModel.MaalerType.ToLower()) && !string.IsNullOrWhiteSpace(x[_factorModel.ReadColumn].ToString()));
                    List<DataRow> outputIsta = new List<DataRow>();
                    foreach (DataRow item in filteredValueIsta.ToList())
                    {
                        item[_factorModel.ReadDateColumn] = Convert.ToDateTime(item[_factorModel.ReadDateColumn]).ToString("dd-MM-yyyy");

                        string[] formats = { "yyyy-MM-dd", "dd-MM-yyyy" };

                        var tester = DateTime.TryParseExact(item[_factorModel.InstallationDateColumn].ToString(), formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime testtemp);

                        if (_factorModel.InstallationDateColumn != -1 && tester)
                        {
                            item[_factorModel.InstallationDateColumn] = Convert.ToDateTime(item[_factorModel.InstallationDateColumn]).ToString("dd-MM-yyyy");
                        }
                        outputIsta.Add(item);
                    }
                    return outputIsta;
                case "Minol":
                    return dtCloned.AsEnumerable().Where(x => !string.IsNullOrEmpty(x[_factorModel.ReadColumn].ToString()) && string.IsNullOrWhiteSpace(x[_factorModel.SearchCriteriaColumn].ToString()));
                case "Techem":
                    IEnumerable<DataRow> filteredValueTechem = dtCloned.AsEnumerable().Where(x => !x[_factorModel.MaalerColumn].ToString().Contains(_factorModel.SearchCriteria));
                    List<DataRow> outputTechem = new List<DataRow>();
                    foreach (DataRow item in filteredValueTechem.ToList())
                    {
                        item[_factorModel.MaalerColumn] = item[_factorModel.MaalerColumn].ToString().Replace("vandsmåler", "t vand");
                        item[_factorModel.ApartmentColumn] = item[_factorModel.ApartmentColumn].ToString().Substring(0, item[_factorModel.ApartmentColumn].ToString().Length - 1);
                        //item[_factorModel.ReadDateColumn] = Convert.ToDateTime(_factorModel.ReadDate.ToString()).ToString("dd-MM-yyyy");
                        outputTechem.Add(item);
                    }
                    return outputTechem;
                case "Brunata":
                    IEnumerable<DataRow> filteredValueBrunata = dtCloned.AsEnumerable().Where(x => Regex.Match(x[_factorModel.SearchCriteriaColumn].ToString().ToUpper(), _factorModel.SearchCriteria).Success && x[_factorModel.ReadDateColumn].ToString().Equals(_factorModel.ReadDateFormatted));
                    List<DataRow> outputBrunata = new List<DataRow>();
                    foreach (DataRow item in filteredValueBrunata.ToList())
                    {
                        item[_factorModel.ReadDateColumn] = Convert.ToDateTime(_factorModel.ReadDate.ToString()).ToString("dd-MM-yyyy");
                        outputBrunata.Add(item);
                    }
                    return outputBrunata;
                default:
                    throw new Exception("Unknow company/vendor name is found.");
            }
        }

        public static IEnumerable<DataRow> ErrorDataByCompany(DataTable _input, FactorModel _factorModel, string _company)
        {

            DataTable dtCloned = _input.Clone();

            if (_factorModel.ReadDateColumn != -1)
            {
                dtCloned.Columns[_factorModel.ReadDateColumn].DataType = typeof(string);
            }

            if (_factorModel.InstallationDateColumn != -1)
            {
                dtCloned.Columns[_factorModel.InstallationDateColumn].DataType = typeof(string);
            }

            foreach (DataRow row in _input.Rows)
            {
                dtCloned.ImportRow(row);
            }

            switch (_company)
            {
                case "Casi":
                    return dtCloned.AsEnumerable().Where(x => x[_factorModel.ReadColumn].ToString().Contains(_factorModel.SearchCriteria));
                case "Ista":
                    IEnumerable<DataRow> filteredValueIsta = dtCloned.AsEnumerable().Where(x => Convert.ToDateTime(x[_factorModel.ReadDateColumn].ToString()).ToString("dd-MM-yyyy").Contains(Convert.ToDateTime(_factorModel.ReadDate.ToString()).ToString("dd-MM-yyyy")) && (!x[_factorModel.MaalerColumn].ToString().ToLower().Contains(_factorModel.MaalerType.ToLower())) || string.IsNullOrWhiteSpace(x[_factorModel.ReadColumn].ToString()));
                    List<DataRow> outputIsta = new List<DataRow>();
                    foreach (DataRow item in filteredValueIsta.ToList())
                    {
                        //var tester = Convert.ToDateTime(item[_factorModel.ReadDateColumn]).ToString("dd-MM-yyyy");
                        item[_factorModel.ReadDateColumn] = Convert.ToDateTime(item[_factorModel.ReadDateColumn]).ToString("dd-MM-yyyy");
                        if (_factorModel.InstallationDateColumn != -1)
                        {
                            item[_factorModel.InstallationDateColumn] = Convert.ToDateTime(item[_factorModel.InstallationDateColumn]).ToString("dd-MM-yyyy");
                        }
                        outputIsta.Add(item);
                    }
                    return outputIsta;
                //return _input.AsEnumerable().Where(x => !x[_factorModel.ReadDateColumn].Equals(_factorModel.SearchCriteria) && !x[_factorModel.MaalerColumn].ToString().Contains(_factorModel.MaalerType));
                case "Minol":
                    return dtCloned.AsEnumerable().Where(x => string.IsNullOrEmpty(x[_factorModel.ReadColumn].ToString()) || !string.IsNullOrWhiteSpace(x[_factorModel.SearchCriteriaColumn].ToString()));
                case "Techem":
                    return dtCloned.AsEnumerable().Where(x => x[_factorModel.MaalerColumn].ToString().Contains(_factorModel.SearchCriteria));
                case "Brunata":
                    IEnumerable<DataRow> filteredValueBrunata = dtCloned.AsEnumerable().Where(x => !Regex.Match(x[_factorModel.SearchCriteriaColumn].ToString().ToUpper(), _factorModel.SearchCriteria).Success);
                    List<DataRow> outputBrunata = new List<DataRow>();
                    foreach (DataRow item in filteredValueBrunata.ToList())
                    {
                        item[_factorModel.ReadDateColumn] = Convert.ToDateTime(_factorModel.ReadDate.ToString()).ToString("dd-MM-yyyy");
                        outputBrunata.Add(item);
                    }
                    return outputBrunata;
                default:
                    throw new Exception("Unknow company/vendor name is found.");
            }
        }
    }
}
