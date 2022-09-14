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
    public class STS201MI
    {
        [FieldAttribute("AGNB", true)]
        public string Agreement_number { get; set; }

        [FieldAttribute("FWHL", true)]
        public string From_warehouse { get; set; }

        [FieldAttribute("LTYP", true)]
        public string Line_type { get; set; }

        [FieldAttribute("ITNO", true)]
        public string Item_number { get; set; }

        [FieldAttribute("ORQT", true)]
        public string Ordered_quantity_basic_UM { get; set; }

        [FieldAttribute("CCAP", true)]
        public string Rental_rate_type { get; set; }

        [FieldAttribute("ANOS", true)]
        public string Number_of_shifts { get; set; }

        [FieldAttribute("CONO", false)]
        public string Company { get; set; }

        [FieldAttribute("BANO", false)]
        public string Serial_number { get; set; }

        [FieldAttribute("FVDT", false)]
        public string Valid_from_Sales_Date { get; set; }

        [FieldAttribute("LVDT", false)]
        public string Valid_to { get; set; }

        [FieldAttribute("PONR", false)]
        public string Line_number { get; set; }

        [FieldAttribute("POSX", false)]
        public string Line_suffix { get; set; }

        [FieldAttribute("VERS", false)]
        public string Version { get; set; }

        [FieldAttribute("SUNO", false)]
        public string Supplier_number { get; set; }

        [FieldAttribute("CUPL", false)]
        public string Customer_site { get; set; }

        [FieldAttribute("SAID", false)]
        public string Address_number { get; set; }

        [FieldAttribute("NOPR", false)]
        public string Number_of_periods { get; set; }

        [FieldAttribute("IPNO", false)]
        public string Included_in_line_number { get; set; }

        [FieldAttribute("PROJ", false)]
        public string Project_number { get; set; }

        [FieldAttribute("ELNO", false)]
        public string Project_element { get; set; }

        [FieldAttribute("FVTM", false)]
        public string Start_time { get; set; }

        [FieldAttribute("ENTM", false)]
        public string End_time { get; set; }

        [FieldAttribute("PNCA", false)]
        public string Net_rate_rental_rate_type { get; set; }

        [FieldAttribute("DLDT", false)]
        public string Planned_delivery_date { get; set; }

        [FieldAttribute("DLTM", false)]
        public string Planned_delivery_time { get; set; }

        [FieldAttribute("COLD", false)]
        public string Planned_pick_up_date { get; set; }

        [FieldAttribute("COTM", false)]
        public string Planned_pick_up_time { get; set; }

        [FieldAttribute("MRTP", false)]
        public string Minimum_rental_type { get; set; }

        [FieldAttribute("MIHP", false)]
        public string Minimum_rental_period { get; set; }

        [FieldAttribute("MINV", false)]
        public string Minimum_order_value { get; set; }

        [FieldAttribute("COAD", false)]
        public string Collection_address { get; set; }

        [FieldAttribute("TWHL", false)]
        public string To_warehouse { get; set; }

        [FieldAttribute("DMOD", false)]
        public string Delivery_method { get; set; }

        [FieldAttribute("CMOD", false)]
        public string Return_delivery_method { get; set; }

        [FieldAttribute("DTED", false)]
        public string Delivery_terms { get; set; }

        [FieldAttribute("CTED", false)]
        public string Return_delivery_terms { get; set; }

        [FieldAttribute("DECH", false)]
        public string Delivery_charge { get; set; }

        [FieldAttribute("CLCH", false)]
        public string Return_charge { get; set; }

        [FieldAttribute("DECO", false)]
        public string Delivery_cost { get; set; }

        [FieldAttribute("CLCO", false)]
        public string Return_cost { get; set; }

        [FieldAttribute("ARCC", false)]
        public string Reason_code_created_agreement { get; set; }

        [FieldAttribute("STRT", false)]
        public string Product_structure_type { get; set; }

        [FieldAttribute("SUFI", false)]
        public string Service { get; set; }

        [FieldAttribute("NODT", false)]
        public string Next_Invoice_Date { get; set; }

        public STS201MI()
        {
        }

        public static Dictionary<string, string> HeaderToFieldMap()
        {
            Dictionary<string, string> pairs = new Dictionary<string, string>();

            PropertyInfo[] props = typeof(STS201MI).GetProperties();
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

        //public static void AddRate()
        //{
        //    string[] cols = { "Currency", "ExchangeRateType", "RateDate", "ExchangeRate", "Division" };
        //    var xlApp = (Excel.Application)ExcelDnaUtil.Application;
        //    var wks = xlApp.ActiveSheet as Excel.Worksheet;
        //    wks.Name = "Transaction";

        //    /*Reimplement mandatory keys/header logic here*/
        //    for (var i = 0; i < cols.Length; i++)
        //    {
        //        wks.Range[wks.Cells[1, i + 1], wks.Cells[1, i + 1]] = cols[i];
        //    }
        //}

        public static void LstRentalLine()
        {
            FieldToColumnMap.ProcessHeader(new STS201MI(), new string[] { "Agreement_number" });

            #region Revisit this after implementing IDisposable

            //var xlApp = (Excel.Application)ExcelDnaUtil.Application;
            //var wks = xlApp.ActiveSheet as Excel.Worksheet;
            //Dictionary<string, List<string>> pairs = new Dictionary<string, List<string>>();

            //PropertyInfo[] props = typeof(CRS055MI).GetProperties();
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
            //string[] cols = { "Currency", "ExchangeRateType", "RateDate", "ExchangeRate", "Division" };

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

        public static void UpdRentalLine()
        {
            FieldToColumnMap.ProcessHeader(new STS201MI(), new string[] { "Agreement_number", "Line_number", "Line_suffix", "Version", "Address_number" });

            #region Revisit this after implementing IDisposable

            //var xlApp = (Excel.Application)ExcelDnaUtil.Application;
            //var wks = xlApp.ActiveSheet as Excel.Worksheet;
            //Dictionary<string, List<string>> pairs = new Dictionary<string, List<string>>();

            //PropertyInfo[] props = typeof(CRS055MI).GetProperties();
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
            //string[] cols = { "Currency", "ExchangeRateType", "RateDate", "ExchangeRate", "Division" };

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

        public static void AddRentalLine()
        {
            FieldToColumnMap.ProcessHeader(new STS201MI(), new string[] { "Agreement_number", "From_warehouse", "Line_type", "Item_number",
                "Ordered_quantity_basic_UM","Rental_rate_type","Number_of_shifts",  "Address_number", "Valid_from_Sales_Date", "Net_rate_rental_rate_type"  });
        }

        /* This looks like the most fastest route to generating a Bulk Data Structure, as the selection.value is returing a fast array
             Would like to avoid using File IO, due to
             1.File IO will always be slow
             2.If the add-in is published in citrix, determining the folder location is not easy
             3.The ExcelDataReader lib might go out of support
        */

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



        /* Might disable this method and start using SelectionToJSON method instead
         */
    }
}