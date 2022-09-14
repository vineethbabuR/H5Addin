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
    public class PPS200MI
    {
        [FieldAttribute("PUNO", true)]
        public string Purchase_order_number { get; set; }

        [FieldAttribute("ITNO", true)]
        public string Item_number { get; set; }

        [FieldAttribute("ORQA", true)]
        public string Ordered_quantity { get; set; } //Ordered_quantity__alternate_U/M

        [FieldAttribute("PUPR", true)]
        public string Purchase_price { get; set; }

        [FieldAttribute("PNLI", true)]
        public string Purchase_order_line { get; set; }

        [FieldAttribute("PNLS", true)]
        public string Purchase_order_line_subnumber { get; set; }

        [FieldAttribute("FACI", false)]
        public string Facility { get; set; }

        [FieldAttribute("WHLO", false)]
        public string Warehouse { get; set; }

        [FieldAttribute("SUNO", false)]
        public string Supplier { get; set; }

        [FieldAttribute("DWDT", false)]
        public string Requested_delivery_date { get; set; }

        [FieldAttribute("SITE", false)]
        public string Supplier_item_number { get; set; }

        [FieldAttribute("PITD", false)]
        public string Purchase_order_item_name { get; set; }

        [FieldAttribute("PITT", false)]
        public string Purchase_order_item_description { get; set; }

        [FieldAttribute("PROD", false)]
        public string Manufacturer { get; set; }

        [FieldAttribute("ECVE", false)]
        public string Revision_number { get; set; }

        [FieldAttribute("REVN", false)]
        public string PO_Revision_number { get; set; }

        [FieldAttribute("ETRF", false)]
        public string External_instruction { get; set; }

        [FieldAttribute("ODI1", false)]
        public string Discount_1 { get; set; }

        [FieldAttribute("ODI2", false)]
        public string Discount_2 { get; set; }

        [FieldAttribute("ODI3", false)]
        public string Discount_3 { get; set; }

        [FieldAttribute("PUUN", false)]
        public string Purchase_order_UM { get; set; }

        [FieldAttribute("PPUN", false)]
        public string Purchase_price_UM { get; set; }

        [FieldAttribute("PUCD", false)]
        public string Purchase_price_quantity { get; set; }

        [FieldAttribute("PTCD", false)]
        public string Purchase_price_text { get; set; }

        [FieldAttribute("RORC", false)]
        public string Reference_order_category { get; set; }

        [FieldAttribute("RORN", false)]
        public string Reference_order_number { get; set; }

        [FieldAttribute("RORL", false)]
        public string Reference_order_line { get; set; }

        [FieldAttribute("RORX", false)]
        public string Line_suffix_1 { get; set; }

        [FieldAttribute("OURR", false)]
        public string Our_reference_number { get; set; }

        [FieldAttribute("OURT", false)]
        public string Reference_type { get; set; }

        [FieldAttribute("PRIP", false)]
        public string Priority { get; set; }

        [FieldAttribute("FUSC", false)]
        public string Monitoring_activity_list { get; set; }

        [FieldAttribute("PURC", false)]
        public string Requisition_by { get; set; }

        [FieldAttribute("BUYE", false)]
        public string Buyer { get; set; }

        [FieldAttribute("TERE", false)]
        public string Technical_supervisor { get; set; }

        [FieldAttribute("GRMT", false)]
        public string Goods_receiving_method { get; set; }

        [FieldAttribute("IRCV", false)]
        public string Recipient { get; set; }

        [FieldAttribute("PACT", false)]
        public string Packaging { get; set; }

        [FieldAttribute("VTCD", false)]
        public string VAT_code { get; set; }

        [FieldAttribute("ACRF", false)]
        public string User_defined_accounting_control_object { get; set; }

        [FieldAttribute("COCE", false)]
        public string Cost_center { get; set; }

        [FieldAttribute("CSNO", false)]
        public string Customs_statistical_number { get; set; }

        [FieldAttribute("ECLC", false)]
        public string Labor_code__trade_statistics_TST { get; set; }

        [FieldAttribute("VRCD", false)]
        public string Business_type__trade_statistics_TST { get; set; }

        [FieldAttribute("PROJ", false)]
        public string Project_number { get; set; }

        [FieldAttribute("ELNO", false)]
        public string Project_element { get; set; }

        [FieldAttribute("CPRI", false)]
        public string Customs_procedure__import { get; set; }

        [FieldAttribute("HAFE", false)]
        public string Harbor_or_airport { get; set; }

        [FieldAttribute("TAXC", false)]
        public string Tax_code_customer_address { get; set; }

        [FieldAttribute("TIHM", false)]
        public string Time_hours_minutes { get; set; }

        [FieldAttribute("MSTN", false)]
        public string Milestone_chain { get; set; }

        [FieldAttribute("UPCK", false)]
        public string Unpack { get; set; }

        [FieldAttribute("ORCO", false)]
        public string Country_of_origin { get; set; }

        [FieldAttribute("GEOC", false)]
        public string Geographical_code { get; set; }

        [FieldAttribute("TRRC", false)]
        public string Order_relation_category { get; set; }

        [FieldAttribute("TRRN", false)]
        public string Order_relation_number { get; set; }

        [FieldAttribute("TRRL", false)]
        public string Order_relation_line { get; set; }

        [FieldAttribute("TRRX", false)]
        public string Line_suffix_2 { get; set; }

        [FieldAttribute("RASN", false)]
        public string Rail_station { get; set; }

        [FieldAttribute("PIAD", false)]
        public string Pickup_address_number { get; set; }

        [FieldAttribute("ORAD", false)]
        public string Origin_address { get; set; }

        [FieldAttribute("UCA1", false)]
        public string User_defined_alpha_field_1 { get; set; }

        [FieldAttribute("UCA2", false)]
        public string User_defined_alpha_field_2 { get; set; }

        [FieldAttribute("UCA3", false)]
        public string User_defined_alpha_field_3 { get; set; }

        [FieldAttribute("UCA4", false)]
        public string User_defined_alpha_field_4 { get; set; }

        [FieldAttribute("UCA5", false)]
        public string User_defined_alpha_field_5 { get; set; }

        [FieldAttribute("UCA6", false)]
        public string User_defined_alpha_field_6 { get; set; }

        [FieldAttribute("UCA7", false)]
        public string User_defined_alpha_field_7 { get; set; }

        [FieldAttribute("UCA8", false)]
        public string User_defined_alpha_field_8 { get; set; }

        [FieldAttribute("UCA9", false)]
        public string User_defined_alpha_field_9 { get; set; }

        [FieldAttribute("UCA0", false)]
        public string User_defined_alpha_field_10 { get; set; }

        [FieldAttribute("UDN1", false)]
        public string User_defined_numeric_1 { get; set; }

        [FieldAttribute("UDN2", false)]
        public string User_defined_numeric_2 { get; set; }

        [FieldAttribute("UDN3", false)]
        public string User_defined_numeric_3 { get; set; }

        [FieldAttribute("UDN4", false)]
        public string User_defined_numeric_4 { get; set; }

        [FieldAttribute("UDN5", false)]
        public string User_defined_numeric_5 { get; set; }

        [FieldAttribute("UDN6", false)]
        public string User_defined_numeric_6 { get; set; }

        [FieldAttribute("UID1", false)]
        public string User_defined_date_1 { get; set; }

        [FieldAttribute("UID2", false)]
        public string User_defined_date_2 { get; set; }

        [FieldAttribute("UID3", false)]
        public string User_defined_date_3 { get; set; }

        [FieldAttribute("UCT1", false)]
        public string User_defined_text_field_1 { get; set; }

        public PPS200MI()
        {
        }

        public static Dictionary<string, string> HeaderToFieldMap()
        {
            Dictionary<string, string> pairs = new Dictionary<string, string>();

            PropertyInfo[] props = typeof(PPS200MI).GetProperties();
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

        public static void AddLine()
        {
            FieldToColumnMap.ProcessHeader(new PPS200MI(), new string[] { "Purchase_order_number", "Item_number", "Ordered_quantity", "Purchase_price", "Requested_delivery_date" });

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

        public static void LstLine()
        {
            FieldToColumnMap.ProcessHeader(new PPS200MI(), new string[] { "Purchase_order_number" });
        }


        public static void UpdLine()
        {
            FieldToColumnMap.ProcessHeader(new PPS200MI(), new string[] { "Purchase_order_number", "Purchase_order_line", "Purchase_order_line_subnumber", "Reference_order_number" });
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

            if (transactionName == "LstLine")
            {
                selectedColumns.Add("PUNO");
                selectedColumns.Add("PNLI");
                selectedColumns.Add("PNLS");
                selectedColumns.Add("RORN");

                
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