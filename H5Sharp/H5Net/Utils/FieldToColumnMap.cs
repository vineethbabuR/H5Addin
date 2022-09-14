using ExcelDna.Integration;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace H5Net.Utils
{
    public class FieldToColumnMap
    {
       
        public static void ProcessHeader<T>(T type, string[] ColumnHeader)
        {
            var xlApp = (Excel.Application)ExcelDnaUtil.Application;
            var wks = xlApp.ActiveSheet as Excel.Worksheet;
            Dictionary<string, List<string>> pairs = new Dictionary<string, List<string>>();

            PropertyInfo[] props = typeof(T).GetProperties();
            foreach (PropertyInfo p in props)
            {
                object[] attrs = p.GetCustomAttributes(true);
                foreach (object attr in attrs)
                {
                    FieldAttribute fAttr = attr as FieldAttribute;
                    string propName = p.Name;
                    string fieldName = fAttr.FieldName;
                    string mandatory = fAttr.Mandatory.ToString();
                    pairs.Add(propName, new List<string> { fieldName, mandatory });
                }
            }

            /*Add required columns here. Column names should match the properties defined above*/
            //string[] cols = ColumnHeader;//{ "Currency", "ExchangeRateType", "RateDate", "ExchangeRate", "Division" };

            /*Block below might not be the most efficient mechanism to support mandatory key highlights
             this operation is deleting the keys from the dictionary which holds the key:Property_name
             which is unique between required column and class properties.
             */
            var colList = ColumnHeader.ToList();
            var keyList = pairs.Keys.ToList();
            var ignoreFields = keyList.Except(colList).ToList();
            foreach (var item in ignoreFields)
            {
                pairs.Remove(item);
            }

            var mandatoryKeys = new List<string>();
            var cellsToFill = pairs.Count;
            var colNames = pairs.Keys.ToArray();

            foreach (var item in pairs)
            {
                foreach (var it in item.Value[1])
                {
                    if (it.ToString() == "T")
                    {
                        mandatoryKeys.Add(item.Key);
                    }
                }
            }

            for (int i = 0; i < cellsToFill; i++)
            {
                wks.Range[wks.Cells[1, i + 1], wks.Cells[1, i + 1]] = colNames[i];
                if (mandatoryKeys.Contains(colNames[i]))
                {
                    wks.Range[wks.Cells[1, i + 1], wks.Cells[1, i + 1]].Font.Color = Excel.XlRgbColor.rgbRed;
                }
                else
                {
                    wks.Range[wks.Cells[1, i + 1], wks.Cells[1, i + 1]].Font.Color = Excel.XlRgbColor.rgbLightGreen;
                }
            }
            wks.Range[wks.Cells[1, 1], wks.Cells[1, cellsToFill]].Interior.Color = Excel.XlRgbColor.rgbBlack;

            wks.Name = "Transaction";
        }

        public static void ProcessHeader<T>(T type, string[] ColumnHeader, string messageNumber)
        {
            var xlApp = (Excel.Application)ExcelDnaUtil.Application;
            var wks = xlApp.ActiveSheet as Excel.Worksheet;
            Dictionary<string, List<string>> pairs = new Dictionary<string, List<string>>();

            PropertyInfo[] props = typeof(T).GetProperties();
            foreach (PropertyInfo p in props)
            {
                object[] attrs = p.GetCustomAttributes(true);
                foreach (object attr in attrs)
                {
                    FieldAttribute fAttr = attr as FieldAttribute;
                    string propName = p.Name;
                    string fieldName = fAttr.FieldName;
                    string mandatory = fAttr.Mandatory.ToString();
                    pairs.Add(propName, new List<string> { fieldName, mandatory });
                }
            }

            /*Add required columns here. Column names should match the properties defined above*/
            //string[] cols = ColumnHeader;//{ "Currency", "ExchangeRateType", "RateDate", "ExchangeRate", "Division" };

            /*Block below might not be the most efficient mechanism to support mandatory key highlights
             this operation is deleting the keys from the dictionary which holds the key:Property_name
             which is unique between required column and class properties.
             */
            var colList = ColumnHeader.ToList();
            var keyList = pairs.Keys.ToList();
            var ignoreFields = keyList.Except(colList).ToList();
            foreach (var item in ignoreFields)
            {
                pairs.Remove(item);
            }

            var mandatoryKeys = new List<string>();
            var cellsToFill = pairs.Count;
            var colNames = pairs.Keys.ToArray();

            foreach (var item in pairs)
            {
                foreach (var it in item.Value[1])
                {
                    if (it.ToString() == "T")
                    {
                        mandatoryKeys.Add(item.Key);
                    }
                }
            }

            for (int i = 0; i < cellsToFill; i++)
            {
                wks.Range[wks.Cells[1, i + 1], wks.Cells[1, i + 1]] = colNames[i];
                if (mandatoryKeys.Contains(colNames[i]))
                {
                    wks.Range[wks.Cells[1, i + 1], wks.Cells[1, i + 1]].Font.Color = Excel.XlRgbColor.rgbRed;
                }
                else
                {
                    wks.Range[wks.Cells[1, i + 1], wks.Cells[1, i + 1]].Font.Color = Excel.XlRgbColor.rgbLightGreen;
                }
            }
            wks.Range[wks.Cells[1, 1], wks.Cells[1, cellsToFill]].Interior.Color = Excel.XlRgbColor.rgbBlack;

            // wks.Range[wks.Cells[1, 1], wks.Cells[1, 1]].CurrentRegion.Columns.NumberFormat = "@";

            wks.Name = "Transaction";
        }
    }
}