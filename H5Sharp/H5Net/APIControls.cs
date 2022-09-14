using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using H5AuthData;
using H5Net.Api;
using H5Net.RESTCalls;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace H5Net
{
    [ComVisible(true)]
    public partial class APIControls : UserControl
    {
        private string ionApiFileName;
        private string bearerToken;
        private string programName;
        private string division;
        private List<string> transactionName = new List<string>();
        private string selectedTransaction;
        private const int company = 1;

        private bool isBatchTransaction = false;
        private string batchMessageNum = "";
        private string purchaseNumber = "";
        private string lineNumber = "";
        private const int MIN_ROW_COUNT = 2;
        private const int MAX_ROW_COUNT = 81;

        //PurchaseAgreementModel purchaseAgreementModel = new PurchaseAgreementModel();

        public APIControls()
        {
            InitializeComponent();
        }

        private void APIControls_Load(object sender, EventArgs e)
        {
            lblEnv.Text = AuthData.SelectedEnvironment;
            var programTransList = ApiTransactionInfo.TransPerProgram();
            //var progList = new List<string>();

            foreach (var k in programTransList)
            {
                cmbProgram.Items.Add(k.Key);
            }

            cmbTransaction.Items.Clear();

            cmbDivision.Items.Add("Default");
            cmbDivision.Items.Add("200");
            cmbDivision.Items.Add("210");
            cmbDivision.Items.Add("240");
            cmbDivision.Items.Add("300");
            cmbDivision.Items.Add("310");

            // Always default to "Default"
            cmbDivision.SelectedIndex = 0;

        }

        private void cmbProgram_SelectedIndexChanged(object sender, EventArgs e)
        {
            cmbTransaction.Text = string.Empty;
            cmbTransaction.Items.Clear();

            transactionName.Clear();

            var selectedItem = cmbProgram.Text;
            programName = selectedItem.Substring(0, 8);

            var programTransList = ApiTransactionInfo.TransPerProgram();
            var transList = new List<string>();

            foreach (var k in programTransList)
            {
                if (k.Key == selectedItem)
                {
                    foreach (var v in k.Value)
                    {
                        transactionName.Add(v);
                    }
                }
            }

            foreach (var tr in transactionName)
            {
                cmbTransaction.Items.Add(tr);
            }
        }

        private void btnCreateHeader_Click(object sender, EventArgs e)
        {
            // Handle disbaling of control if user has not selected an API and Transaction

            txtResponse.Clear();

            var xlApp = (Excel.Application)ExcelDnaUtil.Application;
            var wkb = xlApp.ActiveWorkbook;
            var wks = xlApp.ActiveSheet as Excel.Worksheet;

            var headerRange = wks.Range[wks.Cells[1, 1], wks.Cells[1, 1]].CurrentRegion;
            headerRange.Clear();

            if (string.IsNullOrEmpty(programName) || string.IsNullOrWhiteSpace(programName) || string.IsNullOrEmpty(selectedTransaction) || string.IsNullOrWhiteSpace(selectedTransaction))
            {
                txtResponse.Text = "Error: Select a valid API and Transaction";
                return;
            }

            if(programName == "STS201MI" && selectedTransaction == "AddRentalLine" && division == "Default")
            {
                txtResponse.Text = "Error: Select a valid Division";
                return;
            }


            /*
            var xlApp = (Excel.Application)ExcelDnaUtil.Application;
            var wkb = xlApp.ActiveWorkbook;
            var wks = xlApp.ActiveSheet as Excel.Worksheet;

            var headerRange = wks.Range[wks.Cells[1, 1], wks.Cells[1, 1]].CurrentRegion;
            headerRange.Clear();
        */

            Type apiName = Type.GetType($"H5Net.Api.{programName.ToString()}");
            apiName.InvokeMember(selectedTransaction.ToString(), BindingFlags.InvokeMethod | BindingFlags.Static | BindingFlags.Public, null, null, null);

            //TODO: Frankenstein Block starts
            /*if(programName.ToString() == "PPS370MI" && selectedTransaction.ToString() == "AddHead")
            {
                wks.Range["A2:Z1000"].NumberFormat = "@";
                wks.Range["A2"].Value = batchMessageNum;
            }
            */
        }

        [Obsolete("No longer in use, replaced with Create Json From Selection")]
        private void btnSaveAsCsv_Click(object sender, EventArgs e)
        {
            /*switch (programName)
            {
                case "CRS055MI":
                    {
                        CRS055MI.SaveAsCsv();
                        break;
                    }
            }
            */

            //var jsonFromSelection = CRS055MI.SelectionToJSON(programName.ToString(), transactionName.ToString());
            //txtResponse.Text = jsonFromSelection;
        }

        private void cmbTransaction_SelectedIndexChanged(object sender, EventArgs e)
        {
            transactionName.Clear();
            var selectedItem = cmbTransaction.Text;
            selectedTransaction = selectedItem;
        }

        [Obsolete("No longer in use, replaced with Create Json From Selection")]
        private void btnCreateJson_Click(object sender, EventArgs e)
        {
            Type apiName = Type.GetType($"H5Net.Api.{programName.ToString()}");

            apiName.InvokeMember("SaveAsCsv", BindingFlags.InvokeMethod | BindingFlags.Static | BindingFlags.Public, null, null, null);

            //CRS055MI.SaveAsCsv();

            string csvFilePath = @"C:\Users\vineebabu\Desktop\Git_Projects\Project_Assets_Files\payload_Files\Project.csv";
            string jsonFilePath = @"C:\Users\vineebabu\Desktop\Git_Projects\Project_Assets_Files\payload_Files\jsObject.json";

            var inputParams = new Type[] { typeof(string), typeof(string), typeof(string), typeof(string), typeof(int) };
            var jsonFileMethod = apiName.GetMethod("CreateJsonFile", inputParams);
            object[] passParams = new object[] { csvFilePath, jsonFilePath, programName, selectedTransaction, company };

            jsonFileMethod.Invoke(null, passParams);
        }

        // This finally executes the transaction, transfers the json to H5 over REST
        private async void btnExecTrans_Click(object sender, EventArgs e)
        {
            //Handle disbaling of control if user has not generated a JSON bulk structure

            if (string.IsNullOrEmpty(programName) || string.IsNullOrWhiteSpace(programName) || string.IsNullOrEmpty(selectedTransaction) || string.IsNullOrWhiteSpace(selectedTransaction))
            {
                txtResponse.Text = "Error: Select a valid API and Transaction";
                return;
            }

          //  var mainThread = System.Threading.Thread.CurrentThread.ManagedThreadId;


            var xlApp = (Excel.Application)ExcelDnaUtil.Application;
            var wkb = xlApp.ActiveWorkbook;
            var wks = xlApp.ActiveSheet as Excel.Worksheet;

            var colCount = wks.Range[wks.Cells[1, 1], wks.Cells[1, 1]].CurrentRegion.Columns.Count; //wkb.Worksheets["Transaction"].UsedRange.Columns.count;
            var rowCount = wks.Range[wks.Cells[1, 1], wks.Cells[1, 1]].CurrentRegion.Rows.Count; // wkb.Worksheets["Transaction"].UsedRange.Rows.count;

            if (rowCount < MIN_ROW_COUNT || rowCount > MAX_ROW_COUNT)
            {
                txtResponse.Text = "Error:  Data Row Limit Error";
                return;
            }

            if (txtResponse.Text.Substring(0, 5) == "Error")
            {
                return;
            }

            btnExecTrans.Enabled = false;

            /*Don't clear this now, possibly write JSON bulk data here and execute*/
            //txtResponse.Clear();

            //TODO: Confirm if baseUri is same across TST,TRN,PRD.
            var jsonPayload = (string)txtResponse.Text; //@"C:\Users\vineebabu\Desktop\Git_Projects\Project_Assets_Files\payload_Files\jsObject.json";
            var baseUri = @"https://mingle-ionapi.eu1.inforcloudsuite.com";
            var apiEndPoint = $"/{AuthData.SelectedEnvironment}/M3/m3api-rest/v2/execute";
            var accessToken = AuthData.GetAccessToken(); //txtBearerToken.Text;

            //not sure if usedrange is the right api for counting rows and columns
            // user can have a non contiguous cell populated and usedrange will include that in the count
            // come back later to this.

            var result = new string[3];
            List<string> errorMessageToArray = new List<string>();

            if (programName == "STS201MI" && selectedTransaction == "LstRentalLine")
            {
                // These are mandatory fields for Line Updates
                var agr_records_agrNum = new List<string>();
                var agr_records_lnNum = new List<string>();
                var agr_records_lnSfx = new List<string>();
                var agr_records_Ver = new List<string>();

                var agr_records_Next_InvDt = new List<string>();
                // var agr_records_Item = new List<string>();
                // var agr_records_qty = new List<string>();

                //TODO: pass in Lst generics
                var agreement_lines = await  LoadTransaction.ExecuteLstTransaction(jsonPayload, baseUri, apiEndPoint, accessToken, programName, selectedTransaction);

                //TODO: While returing columns to later update make sure to verify with STS201MI UpdRentalLine, columns should match the update list
                foreach (var items in agreement_lines.results)
                {
                    foreach (var item in items.records)
                    {
                        agr_records_agrNum.Add(item.AGNB);
                        agr_records_lnNum.Add(item.PONR);
                        agr_records_lnSfx.Add(item.POSX);
                        agr_records_Ver.Add(item.VERS);
                        agr_records_Next_InvDt.Add(item.SAID); // this is a test column, will remove this after review. matches with STS201MI UpdRentalLine
                        //agr_records_Item.Add(item.ITNO);
                        //agr_records_qty.Add(item.ORQT);
                    }
                }

                string[] agrNum = agr_records_agrNum.ToArray();
                string[] linNum = agr_records_lnNum.ToArray();
                string[] lnSfx = agr_records_lnSfx.ToArray();
                string[] vers = agr_records_Ver.ToArray();
                string[] nextInvDt = agr_records_Next_InvDt.ToArray();

                //string[] itemNum = agr_records_Item.ToArray();
                //string[] ordqty = agr_records_qty.ToArray();
                var recCount = agrNum.Count() + 1;

                object[] fullData = new object[]
                {
                    agrNum,linNum,lnSfx,vers,nextInvDt
                };

                //var rwCount = fullData.GetLength(0);
                //var clCount = fullData.GetLength(1);

                // Ensure writes happens on the main thread when excel is ready
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    wks.Range[wks.Cells[1, 1], wks.Cells[1, 1]].CurrentRegion.Clear();

                    //Make sure the column name math the class property names
                    wks.Range["A1"].Value = "Agreement_number";
                    wks.Range["B1"].Value = "Line_number";
                    wks.Range["C1"].Value = "Line_suffix";
                    wks.Range["D1"].Value = "Version";
                    wks.Range["E1"].Value = "Address_number";

                    //TODO: Refactor to autoset mandatory fields to red
                    wks.Range["A1:D1"].Interior.Color = Excel.XlRgbColor.rgbBlack;
                    wks.Range["A1:D1"].Font.Color = Excel.XlRgbColor.rgbRed;

                    // Transposing a single object seems to be faster than multiple array transpose
                    wks.Range["A2:E" + recCount.ToString()].Value = xlApp.WorksheetFunction.Transpose(fullData);

                    //wks.Range["B2:B" + recCount.ToString()].Value = xlApp.WorksheetFunction.Transpose(linNum);
                    //wks.Range[wks.Cells[2, 1], wks.Cells[recCount, 1]].Value = xlApp.WorksheetFunction.Transpose(agrNum);

                    btnExecTrans.Enabled = true;
                });

                
            }

            else if(programName == "PPS200MI" && selectedTransaction == "LstLine")
            {
                var po_records_poNum = new List<string>();
                var po_records_poLine = new List<string>();
                var po_records_poSubLine = new List<string>();
                var po_records_poRefOrdNum = new List<string>();

                var purchase_lines = await LoadTransaction.ExecutePurchaseLstTransaction(jsonPayload, baseUri, apiEndPoint, accessToken, programName, selectedTransaction);

                foreach (var items in purchase_lines.results)
                {
                    foreach (var item in items.records)
                    {
                        po_records_poNum.Add(item.PUNO);
                        po_records_poLine.Add(item.PNLI);
                        po_records_poSubLine.Add(item.PNLS);
                        po_records_poRefOrdNum.Add(item.RORN);
                     
                    }
                }

                string[] poNum = po_records_poNum.ToArray();
                string[] lineNum = po_records_poLine.ToArray();
                string[] subLine = po_records_poSubLine.ToArray();
                string[] refOrdNum = po_records_poRefOrdNum.ToArray();

                var recCount = poNum.Count() + 1;

                object[] fullData = new object[]
                {
                    poNum,lineNum,subLine,refOrdNum
                };

                // Ensure writes happens on the main thread when excel is ready
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    wks.Range[wks.Cells[1, 1], wks.Cells[1, 1]].CurrentRegion.Clear();

                    wks.Range["A1"].Value = "Purchase_order_number";
                    wks.Range["B1"].Value = "Purchase_order_line";
                    wks.Range["C1"].Value = "Purchase_order_line_subnumber";
                    wks.Range["D1"].Value = "Reference_order_number";

                    wks.Range["A1:C1"].Interior.Color = Excel.XlRgbColor.rgbBlack;
                    wks.Range["A1:C1"].Font.Color = Excel.XlRgbColor.rgbRed;

                    wks.Range["A2:D" + recCount.ToString()].Value = xlApp.WorksheetFunction.Transpose(fullData);

                    btnExecTrans.Enabled = true;
                });

                

            }

            
            else
            {
                // this is where the json is transferred to H5
                //TODO:  Handle error load result with no value in number of transaction since tokek has expired
                
                var loadResult = await  LoadTransaction.LoadBulkJSONTransaction(jsonPayload, baseUri, apiEndPoint, accessToken);

                
               

              //  await Task.Delay(5000);

                foreach (var item in loadResult.results)
                {
                    //TODO: Write better response handling mechanism for batch transactions
                    // handling batch api transactions
                    

                    if (programName == "PPS370MI" && selectedTransaction == "StartEntry" && item.errorMessage == null)
                    {
                        var messageNo = JObject.Parse(item.records[0].ToString());
                        batchMessageNum = messageNo["MSGN"].Value<string>();
                        isBatchTransaction = true;

                        //purchaseAgreementModel.messageNum = batchMessageNum;

                        errorMessageToArray.Add(item.transaction.Trim() + " - " + "OK");
                    }
                    else if (programName == "PPS370MI" && selectedTransaction == "AddHead" && item.errorMessage == null)
                    {
                        var messageNo = JObject.Parse(item.records[0].ToString());
                        purchaseNumber = messageNo["PUNO"].Value<string>();
                        isBatchTransaction = true;

                        errorMessageToArray.Add(item.transaction.Trim() + " - " + "OK " + "PO - " + purchaseNumber);
                    }
                    else if (programName == "PPS370MI" && selectedTransaction == "AddLine" && item.errorMessage == null)
                    {
                        var messageNo = JObject.Parse(item.records[0].ToString());
                        purchaseNumber = messageNo["PUNO"].Value<string>();
                        lineNumber = messageNo["PNLI"].Value<string>();
                        isBatchTransaction = true;

                        errorMessageToArray.Add(item.transaction.Trim() + " - " + "OK " + "PO - " + purchaseNumber + " - " + lineNumber);
                    }
                    else if (programName == "STS201MI" && selectedTransaction == "UpdRentalLine" && item.errorMessage == null)
                    {
                        var messageNo = JObject.Parse(item.records[0].ToString());
                        var agreementNum = messageNo["AGNB"].Value<string>();
                        //var ln_status = messageNo["ASTH"].Value<string>();
                        isBatchTransaction = true;

                        errorMessageToArray.Add(item.transaction.Trim() + " - " + "OK " + agreementNum);
                    }
                    else
                    {
                        // here handle null error results
                        if (item.errorMessage == null )
                        {
                            errorMessageToArray.Add(item.transaction.Trim() + " - " + "OK");
                        }
                        else if(item.errorField == null)
                        {
                            errorMessageToArray.Add("Error Field" + " - " + item.errorMessage.Trim());
                        }
                        

                        else
                        {
                            errorMessageToArray.Add(item.errorField.Trim() + " - " + item.errorMessage.Trim());
                        }
                        //errorMessageToArray.Add(item.errorMessage.Trim());
                    }
                }

                result = errorMessageToArray.ToArray();

                // Ensure writes happens on the main thread when excel is ready
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    // if the entire payload fails due to json header error
                    if (loadResult.wasTerminated)
                    {
                        
                        wks.Cells[1, colCount + 1].Value = "Response";
                        wks.Range[wks.Cells[2, colCount + 1], wks.Cells[rowCount, colCount + 1]].Value = $"Terminated: {loadResult.terminationReason}";
                        btnExecTrans.Enabled = true;
                        
                    }

                    else
                    {
                        txtResponse.Clear();
                        wks.Cells[1, colCount + 1].Value = "Response";

                        wks.Range[wks.Cells[2, colCount + 1], wks.Cells[rowCount, colCount + 1]].Value = xlApp.WorksheetFunction.Transpose(result);
                        

                        if (isBatchTransaction)
                        {
                            
                            txtResponse.Text += $"Batch No: {batchMessageNum}" + Environment.NewLine;
                        }

                        
                        txtResponse.Text += $"Total: {loadResult.nrOfSuccessfullTransactions + loadResult.nrOfFailedTransactions}" + Environment.NewLine;
                        txtResponse.Text += $"Success: {loadResult.nrOfSuccessfullTransactions}" + Environment.NewLine;
                        txtResponse.Text += $"Failed: {loadResult.nrOfFailedTransactions}" + Environment.NewLine;

                        btnExecTrans.Enabled = true;
                    }
                    

                    //   var mainThreadEndInside = System.Threading.Thread.CurrentThread.ManagedThreadId;
                });

             //   var mainThreadEnd = System.Threading.Thread.CurrentThread.ManagedThreadId;




                // handling batch api transactions
                //TODO: Not required anymore. Handling message batch number back to response range in excel

            }
        }

        private void btnCrtJsonFrSel_Click(object sender, EventArgs e)
        {
            // Handle disbaling of control if user has not selected an API and Transaction and a data range

            txtResponse.Clear();

            if (string.IsNullOrEmpty(programName) || string.IsNullOrWhiteSpace(programName) || string.IsNullOrEmpty(selectedTransaction) || string.IsNullOrWhiteSpace(selectedTransaction))
            {
                txtResponse.Text = "Error: Select a valid API and Transaction";
                return;
            }

            Type[] inputParams;
            MethodInfo jsonSelectionMethod;
            object[] passParams;

            var xlApp = (Excel.Application)ExcelDnaUtil.Application;
            var wkb = xlApp.ActiveWorkbook;
            var wks = xlApp.ActiveSheet as Excel.Worksheet;
            var colCount = wks.Range[wks.Cells[1, 1], wks.Cells[1, 1]].CurrentRegion.Columns.Count;
            var rowCount = wks.Range[wks.Cells[1, 1], wks.Cells[1, 1]].CurrentRegion.Rows.Count;
            

            if (rowCount < MIN_ROW_COUNT || rowCount > MAX_ROW_COUNT)
            {
                txtResponse.Text = "Error:  Data Row Limit Error";
                return;
            }

            Type apiName = Type.GetType($"H5Net.Api.{programName.ToString()}");

            if(division == "Default")
            {
                inputParams = new Type[] { typeof(string), typeof(string), typeof(int) };

                jsonSelectionMethod = apiName.GetMethod("SelectionToJSON", inputParams);

                passParams = new object[] { programName.ToString(), selectedTransaction.ToString(), company };
            } 
            else
            {
                inputParams = new Type[] { typeof(string), typeof(string), typeof(int), typeof(string) };

                jsonSelectionMethod = apiName.GetMethod("SelectionToJSON", inputParams);

                passParams = new object[] { programName.ToString(), selectedTransaction.ToString(), company, division };
            }

            

            var jsonFromSelection = jsonSelectionMethod.Invoke(null, passParams);

            txtResponse.Text = (string)jsonFromSelection;
        }

        private void cmbDivision_SelectedIndexChanged(object sender, EventArgs e)
        {
            var divi = cmbDivision.Text;
            division = divi;
        }
    }

    internal static class CTPManager
    {
        private static CustomTaskPane ctp;

        public static void ShowCTP()
        {
            if (ctp == null)
            {
                ctp = CustomTaskPaneFactory.CreateCustomTaskPane(typeof(APIControls), "H5 API");
                ctp.Visible = true;
                ctp.Width = 400;
                ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionLeft;
                /* ctp.DockPositionStateChange += ctp_DockPositionStateChange;
                ctp.VisibleStateChange += ctp_VisibleStateChange; */
            }
            else
            {
                ctp.Visible = true;
            }
        }

        public static void DeleteCTP()
        {
            if (ctp != null)
            {
                ctp.Delete();
                ctp = null;
                //ctp.Visible = false;
            }
        }

        private static void ctp_VisibleStateChange(CustomTaskPane CustomTaskPaneInst)
        {
            MessageBox.Show($"Visibility Changed To {CustomTaskPaneInst.Visible}");
        }

        private static void ctp_DockPositionStateChange(CustomTaskPane CustomTaskPaneInst)
        {
            MessageBox.Show($"Dock Position Changed To {CustomTaskPaneInst.DockPosition}");
        }
    }
}