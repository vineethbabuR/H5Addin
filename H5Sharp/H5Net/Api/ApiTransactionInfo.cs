using System.Collections.Generic;

namespace H5Net.Api
{
    public class ApiTransactionInfo
    {
        public ApiTransactionInfo(string programName)
        {
        }

        public static Dictionary<string, List<string>> TransPerProgram()
        {
            var transactionList = new Dictionary<string, List<string>>();
            transactionList.Add("CRS055MI-Currency Interface", new List<string> { "AddRate" });
            transactionList.Add("FAS001MI-Fixed Assets", new List<string> { "Add", "UpdateAsset" });
            transactionList.Add("FAS002MI-Depreciation Types Interface", new List<string> { "CrtStdDeprType", "Upd" });
            transactionList.Add("FAS003MI-Value Type Interface", new List<string> { "CrtStdValueType", "Upd" });
            transactionList.Add("MMS100MI-DO/RO Interface", new List<string> { "AddDOLine" });
            transactionList.Add("PPS370MI-Purchase Order Batch Entry", new List<string> { "StartEntry", "AddHead", "AddLine", "AddAccString", "FinishEntry" });
            transactionList.Add("PPS200MI-Purchase Order Interface", new List<string> { "AddLine", "LstLine", "UpdLine" });
            transactionList.Add("PPS205MI-Monitor Purchase Order", new List<string> { "AddMonitor", "DltMonitor", "UpdMonitor" });
            transactionList.Add("STS201MI-STR Agreement Line", new List<string> { "LstRentalLine", "UpdRentalLine", "AddRentalLine" });
            transactionList.Add("MMS175MI-Item Change location", new List<string> { "Update" });
            transactionList.Add("MOS100MI-Maint WO", new List<string> { "AddMtrl" });
            transactionList.Add("MOS070MI-Maint Time Rept", new List<string> { "UpdOperation" });
            transactionList.Add("MMS301MI-Report Stock Count", new List<string> { "UpdStockTake" });

            return transactionList;
        }
    }
}