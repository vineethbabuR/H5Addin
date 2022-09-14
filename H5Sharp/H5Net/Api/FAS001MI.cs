using ExcelDna.Integration;
using H5Net.Utils;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace H5Net.Api
{
    public class FAS001MI
    {
        [FieldAttribute("ASID", true)]
        public string Fixed_asset { get; set; }

        [FieldAttribute("SBNO", true)]
        public string Subnumber { get; set; }

        [FieldAttribute("FADS", true)]
        public string Name { get; set; }

        [FieldAttribute("FATP", true)]
        public string Fixed_asset_type { get; set; }

        [FieldAttribute("CUCD", true)]
        public string Currency { get; set; }

        [FieldAttribute("ARAT", true)]
        public string Exchange_rate { get; set; }

        [FieldAttribute("PPER", true)]
        public string Acquisition_date { get; set; }

        [FieldAttribute("FAQT", true)]
        public string Fixed_asset_quantity { get; set; }

        [FieldAttribute("DIVI", false)]
        public string Division { get; set; }

        [FieldAttribute("TXT1", false)]
        public string Text_line_1 { get; set; }

        [FieldAttribute("TXT2", false)]
        public string Text_line_2 { get; set; }

        [FieldAttribute("ACAT", false)]
        public string Fixed_asset_group { get; set; }

        [FieldAttribute("LOC1", false)]
        public string Location_type_1 { get; set; }

        [FieldAttribute("LOC2", false)]
        public string Location_type_2 { get; set; }

        [FieldAttribute("LOC3", false)]
        public string Location_type_3 { get; set; }

        [FieldAttribute("SRNO", false)]
        public string Serial_number { get; set; }

        [FieldAttribute("PINO", false)]
        public string GUI_picture { get; set; }

        [FieldAttribute("WADT", false)]
        public string Warranty_date { get; set; }

        [FieldAttribute("SECN", false)]
        public string Service_agreement { get; set; }

        [FieldAttribute("SECS", false)]
        public string Service_company { get; set; }

        [FieldAttribute("LCNO", false)]
        public string Leasing_agreement { get; set; }

        [FieldAttribute("LCCO", false)]
        public string Leasing_company { get; set; }

        [FieldAttribute("MPER", false)]
        public string Manufacturing_date { get; set; }

        [FieldAttribute("APER", false)]
        public string Activation_date { get; set; }

        [FieldAttribute("BPER", false)]
        public string Building_permit_date { get; set; }

        [FieldAttribute("SPYN", false)]
        public string Payee { get; set; }

        [FieldAttribute("VONO", false)]
        public string Voucher_number { get; set; }

        [FieldAttribute("CCCO", false)]
        public string Cost_of_capital_method { get; set; }

        [FieldAttribute("AIT2", false)]
        public string Accounting_dimension_2 { get; set; }

        [FieldAttribute("AIT3", false)]
        public string Accounting_dimension_3 { get; set; }

        [FieldAttribute("AIT4", false)]
        public string Accounting_dimension_4 { get; set; }

        [FieldAttribute("AIT5", false)]
        public string Accounting_dimension_5 { get; set; }

        [FieldAttribute("AIT6", false)]
        public string Accounting_dimension_6 { get; set; }

        [FieldAttribute("AIT7", false)]
        public string Accounting_dimension_7 { get; set; }

        [FieldAttribute("REAR", false)]
        public string Planning_area { get; set; }

        [FieldAttribute("PCDA", false)]
        public string Last_physical_inventory_date { get; set; }

        [FieldAttribute("PHCN", false)]
        public string Physical_inventory_number { get; set; }

        [FieldAttribute("PHSN", false)]
        public string Physical_inventory_run_number { get; set; }

        [FieldAttribute("PHCT", false)]
        public string Physical_inventory_text { get; set; }

        [FieldAttribute("INNO", false)]
        public string Individual_item_number { get; set; }

        [FieldAttribute("FRF1", false)]
        public string User_defined_field_1 { get; set; }

        [FieldAttribute("FRF2", false)]
        public string User_defined_field_2 { get; set; }

        [FieldAttribute("FRF3", false)]
        public string User_defined_field_3 { get; set; }

        [FieldAttribute("FRF4", false)]
        public string User_defined_field_4 { get; set; }

        [FieldAttribute("FRF5", false)]
        public string User_defined_field_5 { get; set; }

        [FieldAttribute("LRVD", false)]
        public string Last_revaluation_date { get; set; }

        [FieldAttribute("ITNO", false)]
        public string Item_number { get; set; }

        [FieldAttribute("BANO", false)]
        public string Lot_number { get; set; }

        [FieldAttribute("GEOX", false)]
        public string Geographic_code_X { get; set; }

        [FieldAttribute("GEOY", false)]
        public string Geographic_code_Y { get; set; }

        [FieldAttribute("TAGP", false)]
        public string Tax_asset_group { get; set; }

        [FieldAttribute("LKST", false)]
        public string Like_kind_status { get; set; }

        [FieldAttribute("RESP", false)]
        public string Responsible { get; set; }

        [FieldAttribute("BIRT", false)]
        public string Origin_identity { get; set; }

        [FieldAttribute("UNIT", false)]
        public string Unit_of_measure { get; set; }

        [FieldAttribute("CSNO", false)]
        public string Customs_statistical_number { get; set; }

        [FieldAttribute("PRNR", false)]
        public string Property_number { get; set; }

        //public static int count = 0;

        public FAS001MI()
        {
            // Interlocked.Increment(ref count);
        }

        //public static void Add()
        //{
        //    string[] cols = { "Division","Currency","Fixed_asset", "Subnumber", "Name", "Fixed_asset_type",
        //                      "Exchange_rate", "Activation_date", "Acquisition_date","Manufacturing_date",
        //                      "Fixed_asset_quantity","Individual_item_number" ,
        //                      "Accounting_dimension_2","Accounting_dimension_4" };

        //    var xlApp = (Excel.Application)ExcelDnaUtil.Application;
        //    var wks = xlApp.ActiveSheet as Excel.Worksheet;
        //    wks.Name = "Transaction";

        //    /*Reimplement mandatory keys/header logic here*/
        //    for (var i = 0; i < cols.Length; i++)
        //    {
        //        wks.Range[wks.Cells[1, i + 1], wks.Cells[1, i + 1]] = cols[i];
        //    }
        //}

        //public static void UpdateAsset()
        //{
        //    string[] cols = { "Division", "Fixed_asset", "Subnumber", "Name", "Text_line_1", "Text_line_2", "Fixed_asset_type",
        //                      "Activation_date","Acquisition_date", "Manufacturing_date","Individual_item_number",
        //                       "Accounting_dimension_2" };

        //    var xlApp = (Excel.Application)ExcelDnaUtil.Application;
        //    var wks = xlApp.ActiveSheet as Excel.Worksheet;
        //    wks.Name = "Transaction";

        //    /*Reimplement mandatory keys/header logic here*/
        //    for (var i = 0; i < cols.Length; i++)
        //    {
        //        wks.Range[wks.Cells[1, i + 1], wks.Cells[1, i + 1]] = cols[i];
        //    }
        //}

        public static Dictionary<string, string> HeaderToFieldMap()
        {
            Dictionary<string, string> pairs = new Dictionary<string, string>();

            PropertyInfo[] props = typeof(FAS001MI).GetProperties();
            foreach (PropertyInfo p in props)
            {
                object[] attrs = p.GetCustomAttributes(true);
                foreach (object attr in attrs)
                {
                    FieldAttribute fAttr = attr as FieldAttribute;
                    string propName = p.Name;
                    string fieldName = fAttr.FieldName;

                    pairs.Add(propName, fieldName);
                }
            }

            return pairs;
        }

        public static void Add()
        {
            FieldToColumnMap.ProcessHeader(new FAS001MI(), new string[] { "Division","Currency","Fixed_asset", "Subnumber", "Name", "Fixed_asset_type",
                                                                          "Exchange_rate", "Activation_date", "Acquisition_date","Manufacturing_date",
                                                                          "Fixed_asset_quantity","Individual_item_number" ,
                                                                           "Accounting_dimension_2","Accounting_dimension_4" });

            #region Revisit this after implementing IDisposable

            //var xlApp = (Excel.Application)ExcelDnaUtil.Application;
            //var wks = xlApp.ActiveSheet as Excel.Worksheet;

            //Dictionary<string, List<string>> pairs = new Dictionary<string, List<string>>();

            //PropertyInfo[] props = typeof(FAS001MI).GetProperties();
            //foreach (PropertyInfo p in props)
            //{
            //    object[] attrs = p.GetCustomAttributes(true);
            //    foreach (object attr in attrs)
            //    {
            //        FieldAttribute fAttr = attr as FieldAttribute;
            //        string propName = p.Name;
            //        string fieldName = fAttr.FieldName;
            //        string mandatory = fAttr.Mandatory.ToString();
            //        pairs.Add(propName, new List<string> { fieldName, mandatory });
            //    }
            //}

            ///*Add required columns here. Column names should match the properties defined above*/
            //string[] cols = { "Division","Currency","Fixed_asset", "Subnumber", "Name", "Fixed_asset_type",
            //                  "Exchange_rate", "Activation_date", "Acquisition_date","Manufacturing_date",
            //                  "Fixed_asset_quantity","Individual_item_number" ,
            //                 "Accounting_dimension_2","Accounting_dimension_4" };

            ///*Block below might not be the most efficient mechanism to support mandatory key highlights
            // this operation is deleting the keys from the dictionary which holds the key:Property_name
            // which is unique between required column and class properties.
            // */
            //var colList = cols.ToList();
            //var keyList = pairs.Keys.ToList();
            //var ignoreFields = keyList.Except(colList).ToList();
            //foreach (var item in ignoreFields)
            //{
            //    pairs.Remove(item);
            //}

            //var mandatoryKeys = new List<string>();
            //var cellsToFill = pairs.Count;
            //var colNames = pairs.Keys.ToArray();

            //foreach (var item in pairs)
            //{
            //    foreach (var it in item.Value[1])
            //    {
            //        if (it.ToString() == "T")
            //        {
            //            mandatoryKeys.Add(item.Key);
            //        }
            //    }
            //}

            //for (int i = 0; i < cellsToFill; i++)
            //{
            //    wks.Range[wks.Cells[1, i + 1], wks.Cells[1, i + 1]] = colNames[i];
            //    if (mandatoryKeys.Contains(colNames[i]))
            //    {
            //        wks.Range[wks.Cells[1, i + 1], wks.Cells[1, i + 1]].Font.Color = Excel.XlRgbColor.rgbRed;
            //    }
            //    else
            //    {
            //        wks.Range[wks.Cells[1, i + 1], wks.Cells[1, i + 1]].Font.Color = Excel.XlRgbColor.rgbLightGreen;
            //    }
            //}
            //wks.Range[wks.Cells[1, 1], wks.Cells[1, cellsToFill]].Interior.Color = Excel.XlRgbColor.rgbBlack;

            //wks.Name = "Transaction";

            #endregion Revisit this after implementing IDisposable
        }

        public static void UpdateAsset()
        {
            FieldToColumnMap.ProcessHeader(new FAS001MI(), new string[] { "Division", "Fixed_asset", "Subnumber", "Name",
                                                                          "Text_line_1", "Text_line_2", "Fixed_asset_type",
                                                                          "Activation_date","Acquisition_date", "Manufacturing_date","Individual_item_number",
                                                                           "Accounting_dimension_2" });

            #region Revisit this after implementing IDisposable

            //var xlApp = (Excel.Application)ExcelDnaUtil.Application;
            //var wks = xlApp.ActiveSheet as Excel.Worksheet;
            //Dictionary<string, List<string>> pairs = new Dictionary<string, List<string>>();

            //PropertyInfo[] props = typeof(FAS001MI).GetProperties();
            //foreach (PropertyInfo p in props)
            //{
            //    object[] attrs = p.GetCustomAttributes(true);
            //    foreach (object attr in attrs)
            //    {
            //        FieldAttribute fAttr = attr as FieldAttribute;
            //        string propName = p.Name;
            //        string fieldName = fAttr.FieldName;
            //        string mandatory = fAttr.Mandatory.ToString();
            //        pairs.Add(propName, new List<string> { fieldName, mandatory });
            //    }
            //}

            ///*Add required columns here. Column names should match the properties defined above*/
            //string[] cols = { "Division", "Fixed_asset", "Subnumber", "Name", "Text_line_1", "Text_line_2", "Fixed_asset_type",
            //                  "Activation_date","Acquisition_date", "Manufacturing_date","Individual_item_number",
            //                 "Accounting_dimension_2" };

            ///*Block below might not be the most efficient mechanism to support mandatory key highlights
            // this operation is deleting the keys from the dictionary which holds the key:Property_name
            // which is unique between required column and class properties.
            // */
            //var colList = cols.ToList();
            //var keyList = pairs.Keys.ToList();
            //var ignoreFields = keyList.Except(colList).ToList();
            //foreach (var item in ignoreFields)
            //{
            //    pairs.Remove(item);
            //}

            //var mandatoryKeys = new List<string>();
            //var cellsToFill = pairs.Count;
            //var colNames = pairs.Keys.ToArray();

            //foreach (var item in pairs)
            //{
            //    foreach (var it in item.Value[1])
            //    {
            //        if (it.ToString() == "T")
            //        {
            //            mandatoryKeys.Add(item.Key);
            //        }
            //    }
            //}

            //for (int i = 0; i < cellsToFill; i++)
            //{
            //    wks.Range[wks.Cells[1, i + 1], wks.Cells[1, i + 1]] = colNames[i];
            //    if (mandatoryKeys.Contains(colNames[i]))
            //    {
            //        wks.Range[wks.Cells[1, i + 1], wks.Cells[1, i + 1]].Font.Color = Excel.XlRgbColor.rgbRed;
            //    }
            //    else
            //    {
            //        wks.Range[wks.Cells[1, i + 1], wks.Cells[1, i + 1]].Font.Color = Excel.XlRgbColor.rgbLightGreen;
            //    }
            //}
            //wks.Range[wks.Cells[1, 1], wks.Cells[1, cellsToFill]].Interior.Color = Excel.XlRgbColor.rgbBlack;

            //wks.Name = "Transaction";

            #endregion Revisit this after implementing IDisposable
        }

        public static string SelectionToJSON(string programName, string transactionName, int company)
        {
            var xlApp = (Excel.Application)ExcelDnaUtil.Application;
            var wkb = xlApp.ActiveWorkbook;
            var wks = xlApp.ActiveSheet as Excel.Worksheet;

            bool allowTransaction = true;

            //var fieldNames = HeaderToFieldMap()["Currency"];
            //var h1 = fieldNames["Currency"];

            dynamic bulkApi = new JObject();
            bulkApi.program = programName;
            bulkApi.cono = company;
            bulkApi.maxReturnedRecords = 3;

            var transactions = new JArray() as dynamic;
            dynamic transaction = new JObject();

            dynamic record = new JObject();

            var colCount = wks.Range[wks.Cells[1, 1], wks.Cells[1, 1]].CurrentRegion.Columns.Count;  //wkb.Worksheets["Transaction"].UsedRange.Columns.count;
            var rowCount = wks.Range[wks.Cells[1, 1], wks.Cells[1, 1]].CurrentRegion.Rows.Count; //wkb.Worksheets["Transaction"].UsedRange.Rows.count;

            /*Block below will auto select the Current Region, this prevents partial or incomplete selection by the user*/

            var dataRange = wks.Range[wks.Cells[1, 1], wks.Cells[1, 1]].CurrentRegion.Select();

            Excel.Range selectedRange = xlApp.Selection as Excel.Range;

            /************************/

            if (selectedRange == null)
            {
                MessageBox.Show("No Range Selected");
            }
            else
            {
                object values = selectedRange.Value;
                object[,] valuesArray = values as object[,];

                try
                {
                    for (int i = 2; i <= rowCount; i++)
                    {
                        record = new JObject();

                        for (int j = valuesArray.GetLowerBound(1); j <= valuesArray.GetUpperBound(1); j++)
                        {
                            if (valuesArray == null)
                            {
                                return "Empty Array";
                            }
                            else
                            {
                                var header = HeaderToFieldMap()[(valuesArray[1, j]).ToString()];
                                var value = valuesArray[i, j]?.ToString();

                                if (value != null)
                                {
                                    record.Add(header, value);
                                }
                            }
                        }

                        transaction.transaction = transactionName.ToString();
                        transaction.record = record;

                        transactions.Add(transaction);
                        transaction = new JObject();
                        bulkApi.transactions = transactions;
                        allowTransaction = true;
                    }
                }
                catch (System.Collections.Generic.KeyNotFoundException ex)
                {
                    MessageBox.Show("Exception => Incorrect Column", "Header Error", MessageBoxButtons.OK);
                    allowTransaction = false;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Exception => {ex.Message}");
                    allowTransaction = false;
                }
            }

            //Second Message Box
            //MessageBox.Show($"This is from Second Message Box =>  {bulkApi.ToString()}");
            if (allowTransaction)
            {
                //MessageBox.Show(count.ToString());
                return bulkApi.ToString();
            }
            else
            {
                return "Error";
            }
        }

        public static string SelectionToJSON(string programName, string transactionName, int company, string division)
        {
            var xlApp = (Excel.Application)ExcelDnaUtil.Application;
            var wkb = xlApp.ActiveWorkbook;
            var wks = xlApp.ActiveSheet as Excel.Worksheet;

            bool allowTransaction = true;

            //var fieldNames = HeaderToFieldMap()["Currency"];
            //var h1 = fieldNames["Currency"];

            dynamic bulkApi = new JObject();
            bulkApi.program = programName;
            bulkApi.cono = company;
            bulkApi.divi = division;
            bulkApi.maxReturnedRecords = 0;

            var transactions = new JArray() as dynamic;
            var selectedColumns = new JArray() as dynamic;
            dynamic transaction = new JObject();

            dynamic record = new JObject();

            var colCount = wks.Range[wks.Cells[1, 1], wks.Cells[1, 1]].CurrentRegion.Columns.Count;  //wkb.Worksheets["Transaction"].UsedRange.Columns.count;
            var rowCount = wks.Range[wks.Cells[1, 1], wks.Cells[1, 1]].CurrentRegion.Rows.Count; //wkb.Worksheets["Transaction"].UsedRange.Rows.count;

            //TODO: Do we provide an option to the user to select the columns ?

            if (transactionName == "LstRentalLine")
            {
                selectedColumns.Add("AGNB");
                selectedColumns.Add("PONR");
                selectedColumns.Add("POSX");
                selectedColumns.Add("VERS");
                selectedColumns.Add("SAID");
            }
            else
            {
                selectedColumns.Clear();
                // selectedColumns.Add("AGNB");
                // selectedColumns.Add("ASTH");
            }

            /*Block below will auto select the Current Region, this prevents partial or incomplete selection by the user*/

            var dataRange = wks.Range[wks.Cells[1, 1], wks.Cells[1, 1]].CurrentRegion.Select();

            Excel.Range selectedRange = xlApp.Selection as Excel.Range;

            /************************/

            if (selectedRange == null)
            {
                MessageBox.Show("No Range Selected");
            }
            else
            {
                object values = selectedRange.Value;
                object[,] valuesArray = values as object[,];

                try
                {
                    for (int i = 2; i <= rowCount; i++)
                    {
                        record = new JObject();

                        for (int j = valuesArray.GetLowerBound(1); j <= valuesArray.GetUpperBound(1); j++)
                        {
                            if (valuesArray == null)
                            {
                                return "Empty Array";
                            }
                            else
                            {
                                var header = HeaderToFieldMap()[(valuesArray[1, j]).ToString()];
                                var value = valuesArray[i, j]?.ToString();

                                if (value != null)
                                {
                                    record.Add(header, value);
                                }
                            }
                        }

                        transaction.transaction = transactionName.ToString();
                        transaction.selectedColumns = selectedColumns;
                        transaction.record = record;

                        transactions.Add(transaction);
                        transaction = new JObject();
                        bulkApi.transactions = transactions;
                        allowTransaction = true;
                    }
                }
                catch (System.Collections.Generic.KeyNotFoundException ex)
                {
                    MessageBox.Show("Exception => Incorrect Column", "Header Error", MessageBoxButtons.OK);
                    allowTransaction = false;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Exception => {ex.Message}");
                    allowTransaction = false;
                }
            }

            //Second Message Box
            //MessageBox.Show($"This is from Second Message Box =>  {bulkApi.ToString()}");

            if (allowTransaction)
            {
                return bulkApi.ToString();
            }
            else
            {
                return "Error";
            }
        }
    }
}