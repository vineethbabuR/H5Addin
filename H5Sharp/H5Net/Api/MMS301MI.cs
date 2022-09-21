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
    public class MMS301MI
    {
        [FieldAttribute("WHLO", true)]
        public string Warehouse { get; set; }

        [FieldAttribute("STNB", true)]
        public string Physical_inventory_number { get; set; }

        [FieldAttribute("STRN", true)]
        public string Physical_inventory_line { get; set; }

        [FieldAttribute("STQI", true)]
        public string Physical_inventory_quantity { get; set; }

        [FieldAttribute("STBT", true)]
        public string Onhand_balance_to_compare { get; set; }

        [FieldAttribute("STDI", false)]
        public string Physical_inventory_date { get; set; }

        [FieldAttribute("STTM", false)]
        public string Physical_inventory_time { get; set; }

        [FieldAttribute("ITNO", false)]
        public string Item_number { get; set; }

        [FieldAttribute("STAG", false)]
        public string Status_physical_inventory { get; set; }





        //public static int count = 0;

        public MMS301MI()
        {
            //Interlocked.Increment(ref count);
        }

        public static Dictionary<string, string> HeaderToFieldMap()
        {
            Dictionary<string, string> pairs = new Dictionary<string, string>();

            PropertyInfo[] props = typeof(MMS301MI).GetProperties();
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

        public static void UpdStockTake()
        {
            FieldToColumnMap.ProcessHeader(new MMS301MI(), new string[] { "Warehouse",
                "Physical_inventory_number", "Physical_inventory_line", "Item_number", "Physical_inventory_quantity", "Physical_inventory_date",
            "Physical_inventory_time", "Onhand_balance_to_compare", "Status_physical_inventory"});

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