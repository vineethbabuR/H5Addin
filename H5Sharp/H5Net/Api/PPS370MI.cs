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
    public class PPS370MI
    {
        [FieldAttribute("BAOR", true)]
        public string BatchOrigin { get; set; }

        [FieldAttribute("MSGN", true)]
        public string MessageNumber { get; set; }

        [FieldAttribute("FACI", true)]
        public string Facility { get; set; }

        [FieldAttribute("WHLO", true)]
        public string Warehouse { get; set; }

        [FieldAttribute("SUNO", true)]
        public string Supplier_number { get; set; }

        [FieldAttribute("DWDT", true)]
        public string Requested_delivery_date { get; set; }

        [FieldAttribute("HREF", false)]
        public string Purchase_order_head_reference { get; set; }

        [FieldAttribute("ORTY", true)]
        public string Order_type { get; set; }

        [FieldAttribute("CMCO", false)]
        public string Communication_code { get; set; }

        [FieldAttribute("PUDT", false)]
        public string Order_date { get; set; }

        [FieldAttribute("LNCD", false)]
        public string Language { get; set; }

        [FieldAttribute("CUCD", false)]
        public string Currency { get; set; }

        [FieldAttribute("TEPY", false)]
        public string Payment_terms { get; set; }

        [FieldAttribute("PYME", false)]
        public string Payment_method_accounts_payable { get; set; }

        [FieldAttribute("MODL", false)]
        public string Delivery_method { get; set; }

        [FieldAttribute("TEDL", false)]
        public string Delivery_terms { get; set; }

        [FieldAttribute("TEAF", false)]
        public string Freight_terms { get; set; }

        [FieldAttribute("TEPA", false)]
        public string Packaging_terms { get; set; }

        [FieldAttribute("YRE1", false)]
        public string Your_reference { get; set; }

        [FieldAttribute("PRSU", false)]
        public string Payee { get; set; }

        [FieldAttribute("OURR", false)]
        public string Our_reference_number { get; set; }

        [FieldAttribute("OURT", false)]
        public string Reference_type { get; set; }

        [FieldAttribute("AGNT", false)]
        public string Recipient_agreement_type_1_commission { get; set; }

        [FieldAttribute("PURC", false)]
        public string Requisition_by { get; set; }

        [FieldAttribute("BUYE", false)]
        public string Buyer { get; set; }

        [FieldAttribute("FUSC", false)]
        public string Monitoring_activity_list { get; set; }

        [FieldAttribute("TFNO", false)]
        public string Facsimile_transmission_number { get; set; }

        [FieldAttribute("LRED", false)]
        public string Last_reply_date { get; set; }

        [FieldAttribute("TEL1", false)]
        public string Terms_text { get; set; }

        [FieldAttribute("DUDT", false)]
        public string Due_date { get; set; }

        [FieldAttribute("CUTE", false)]
        public string Currency_terms { get; set; }

        [FieldAttribute("AGRA", false)]
        public string Agreed_rate { get; set; }

        [FieldAttribute("PROJ", false)]
        public string Project_number { get; set; }

        [FieldAttribute("ELNO", false)]
        public string Project_element { get; set; }

        [FieldAttribute("HAFE", false)]
        public string Harbor_or_airport { get; set; }

        [FieldAttribute("USD1", false)]
        public string User_defined_field1 { get; set; }

        [FieldAttribute("USD2", false)]
        public string User_defined_field2 { get; set; }

        [FieldAttribute("USD3", false)]
        public string User_defined_field3 { get; set; }

        [FieldAttribute("USD4", false)]
        public string User_defined_field4 { get; set; }

        [FieldAttribute("USD5", false)]
        public string User_defined_field5 { get; set; }

        [FieldAttribute("RASN", false)]
        public string Rail_station { get; set; }

        [FieldAttribute("PUNO", false)]
        public string Purchase_order_number { get; set; }

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

        [FieldAttribute("ITNO", true)]
        public string Item_number { get; set; }

        [FieldAttribute("ORQA", true)]
        public string Ordered_quantity { get; set; } //Ordered_quantity_alternate_UM

        [FieldAttribute("PNLI", true)]
        public string Purchase_order_line { get; set; }

        [FieldAttribute("LREF", true)]
        public string Purchase_order_line_reference { get; set; }

        [FieldAttribute("SITE", true)]
        public string Supplier_item_number { get; set; }

        [FieldAttribute("PITD", true)]
        public string Purchase_order_item_name { get; set; }

        [FieldAttribute("PITT", true)]
        public string Purchase_order_item_description { get; set; }

        [FieldAttribute("PUPR", true)]
        public string Purchase_price { get; set; }

        [FieldAttribute("ODI1", true)]
        public string Discount_1 { get; set; }

        [FieldAttribute("ODI2", true)]
        public string Discount_2 { get; set; }

        [FieldAttribute("ODI3", true)]
        public string Discount_3 { get; set; }

        [FieldAttribute("RORC", true)]
        public string Reference_order_category { get; set; }

        [FieldAttribute("RORN", true)]
        public string Reference_order_number { get; set; }

        [FieldAttribute("RORL", true)]
        public string Reference_order_line { get; set; }

        [FieldAttribute("RORX", true)]
        public string Line_suffix { get; set; }

        [FieldAttribute("AIT1", true)]
        public string Dim1 { get; set; }

        [FieldAttribute("AIT2", true)]
        public string Dim2 { get; set; }

        [FieldAttribute("AIT3", true)]
        public string Dim3 { get; set; }

        [FieldAttribute("AIT4", true)]
        public string Dim4 { get; set; }

        [FieldAttribute("AIT5", true)]
        public string Agreement_Number { get; set; }

        [FieldAttribute("AIT6", true)]
        public string Dim6 { get; set; }

        [FieldAttribute("AIT7", true)]
        public string Dim7 { get; set; }

        /*
         [FieldAttribute("LNUM", true)]
         public string Lines { get; set; }
        */

        public PPS370MI()
        {
        }

        public static Dictionary<string, string> HeaderToFieldMap()
        {
            Dictionary<string, string> pairs = new Dictionary<string, string>();

            PropertyInfo[] props = typeof(PPS370MI).GetProperties();
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

        public static void StartEntry()
        {
            FieldToColumnMap.ProcessHeader(new PPS370MI(), new string[] { "BatchOrigin" });
        }

        //public static void FinishEntry()
        //{
        //    string[] cols = { "MessageNumber" };
        //    var xlApp = (Excel.Application)ExcelDnaUtil.Application;
        //    var wks = xlApp.ActiveSheet as Excel.Worksheet;
        //    wks.Name = "Transaction";

        //    /*Reimplement mandatory keys/header logic here*/
        //    for (var i = 0; i < cols.Length; i++)
        //    {
        //        wks.Range[wks.Cells[1, i + 1], wks.Cells[1, i + 1]] = cols[i];
        //    }
        //}

        public static void FinishEntry()
        {
            FieldToColumnMap.ProcessHeader(new PPS370MI(), new string[] { "MessageNumber" });
        }

        public static void AddHead()
        {
            FieldToColumnMap.ProcessHeader(new PPS370MI(), new string[] { "MessageNumber", "Facility", "Warehouse", "Order_type", "Supplier_number", "Requested_delivery_date" });
        }

        public static void AddLine()
        {
            FieldToColumnMap.ProcessHeader(new PPS370MI(), new string[] { "MessageNumber", "Purchase_order_number", "Purchase_order_line", "Item_number",
                                                                          "Ordered_quantity","Purchase_price", "Requested_delivery_date","Reference_order_number" });
        }

        public static void AddAccString()
        {
            FieldToColumnMap.ProcessHeader(new PPS370MI(), new string[] { "MessageNumber", "Purchase_order_number", "Purchase_order_line", "Item_number",
                                                                          "Dim1","Dim2","Dim3","Dim4","Agreement_Number","Dim6","Dim7"  });
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

                                //PPS370MI/AddAccString does not require an ITNO as input. Added as header to facilitate the user to know the Dim1-7 per item
                                // but ignoring ITNO while JSON DS creation

                                if (programName == "PPS370MI" && transactionName == "AddAccString")
                                {
                                    if (value != null && header != "ITNO")
                                    {
                                        record.Add(header, value);
                                    }
                                }
                                else
                                {
                                    if (value != null /* && header != "LNUM" */)
                                    {
                                        record.Add(header, value);
                                    }
                                }

                                /*if (value != null *//* && header != "LNUM" *//*)
                                {
                                    record.Add(header, value);
                                }*/
                            }
                        }

                        transaction.transaction = transactionName.ToString();
                        transaction.record = record;

                        transactions.Add(transaction);
                        transaction = new JObject();
                        bulkApi.transactions = transactions;
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