using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Text.RegularExpressions;

namespace Methods_Console
{
    public enum Column { Level = 1, SubClass, BeiNum, BeiRev, BeiDesc, FindNum, Quantity, UnitOfMeasure, RefDes, BomNotes };
    class BOMExplosionParser : FileParser
    {
        public override string FileType { get; set; }
        public override string FileName { get; set; }
        public override string FullFilePath { get; set; }
        public string AssemblyName { get; private set; }
        public string DateOfListing { get; private set; }
        public string AssyDescription { get; private set; }
        public string Rev { get; private set; }
        public bool IsValid { get; private set; }
        public Dictionary<string, List<string>> BomMap { get; private set; } //Key = '<FindNum>', List = [0]Part Number, [1]Description, [2]comma-separated ref des's, [3]bomqty
        public object[,] ValueArray { get; private set; }


        public BOMExplosionParser(string path, string fileExtension)
        {
            AssemblyName = null;
            DateOfListing = null;
            AssyDescription = null;
            Rev = null;
            FullFilePath = path;
            FileName = Path.GetFileName(FullFilePath);
            FileType = Path.GetExtension(FullFilePath).ToLower();
            BomMap = new Dictionary<string, List<string>>();          
            ReadBomExplosion();
            if (IsValidBomExplosion())
            {
                ProcessExcelBOM();
                SetValid();
            }                  
            else
            {
                MessageBox.Show("Excel BOM files must be Agile BOM Explosions(NASHUA METHODS format), exported to .csv or .xls.\n\nPlease reload a valid BOM.", "Agile BOM Format Error", MessageBoxButton.OK, MessageBoxImage.Error);
                ClearData();
                IsValid = false;
            }
        }
        private void ClearData()
        {
            FileType = null;
            FileName = null;
            FullFilePath = null;
            if (ValueArray != null && ValueArray.Length > 0)
                Array.Clear(ValueArray, 1, ValueArray.Length);
            BomMap.Clear();
            AssemblyName = null;
            DateOfListing = null;
            AssyDescription = null;
            Rev = null;
        }
        private void SetValid()
        {
            if (!string.IsNullOrEmpty(AssemblyName) && !string.IsNullOrEmpty(AssyDescription) && !string.IsNullOrEmpty(DateOfListing) && !string.IsNullOrEmpty(Rev) && BomMap.Count > 0)
                IsValid = true;
            else
                IsValid = false;
        }

        private void ReadBomExplosion()
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(FullFilePath,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);
            try
            {
                Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                ValueArray = xlRange.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);

                GC.Collect();
                GC.WaitForPendingFinalizers();

                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }
            catch (Exception e)
            {
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                MessageBox.Show("Problem loading BOM. Please check the file and try again.\n" + e.Message, "ParseBomExplosion()", MessageBoxButton.OK, MessageBoxImage.Error);
                ClearData();
                IsValid = false;
            }    
        }

        private void ProcessExcelBOM()
        {
            int row = ProcessHeader();
            ProcessBomItems(row);
            
        }
        private void ProcessBomItems(int nFirstPartRow)
        {
            int row = nFirstPartRow;
            string strPN = null, strDesc = null, strFindNum = null, strQty = null, strRefDes = null;
            Regex reLevelOne = new Regex(@"\.\s1");
            Regex rePart = new Regex(@"^Part\b");
            while (row <= ValueArray.GetLength(0))
            {
                if(reLevelOne.Match(ValueArray[row, (int)Column.Level].ToString()).Success && rePart.Match(ValueArray[row, (int)Column.SubClass].ToString()).Success)
                {
                    strPN = ValueArray[row, (int)Column.BeiNum].ToString().Trim();
                    strDesc = ValueArray[row, (int)Column.BeiDesc].ToString().Trim();
                    strFindNum = ValueArray[row, (int)Column.FindNum].ToString().Trim();
                    strQty = ValueArray[row, (int)Column.Quantity].ToString().Trim();
                    if (ValueArray[row, (int)Column.RefDes] == null || string.IsNullOrWhiteSpace(ValueArray[row, (int)Column.RefDes].ToString()))
                        strRefDes = "No Ref Des";
                    else
                        strRefDes = ValueArray[row, (int)Column.RefDes].ToString().Trim();

                    BomMap.Add(strFindNum, new List<string> { strPN, strDesc, strRefDes, strQty });
                }
                ++row;
            }

        }
        private int ProcessHeader()
        {
            int row = 1, col = 1;
            string strLevel = null, strCellValue = null;
            string[] arrTemp = null;
            while (string.IsNullOrEmpty(strLevel))
            {
                if (ValueArray[row, col] == null || (string.IsNullOrWhiteSpace(ValueArray[row, col].ToString())))
                {
                    ++row;
                    continue;
                }                   
                arrTemp = ValueArray[row, col].ToString().Split();
                strCellValue = string.Join(" ", arrTemp).Trim();
                if (strCellValue.Equals("Create Time:"))
                {
                    arrTemp = ValueArray[row, col + 1].ToString().Split();
                    strCellValue = arrTemp[0];
                    DateOfListing = strCellValue;
                }
                else if (strCellValue.Equals("Item Number:"))
                {
                    arrTemp = ValueArray[row, col + 1].ToString().Split();
                    strCellValue = string.Join(" ", arrTemp);
                    AssemblyName = strCellValue;
                }
                else if (strCellValue.Equals("Description:"))
                {
                    arrTemp = ValueArray[row, col + 1].ToString().Split();
                    strCellValue = string.Join(" ", arrTemp);
                    AssyDescription = strCellValue;
                }
                else if (strCellValue.Equals("Item Revision:"))
                {
                    arrTemp = ValueArray[row, col + 1].ToString().Split();
                    strCellValue = arrTemp[0];
                    Rev = strCellValue;
                }
                else if (strCellValue.Equals("Level"))
                {
                    strLevel = strCellValue;
                }
                ++row;
            }
            return row;
        }

        private bool IsValidBomExplosion()
        {
            bool isValid = false;
            string strbomexp = null;
            if (ValueArray[1, 1] != null)
            {
                string[] strCsvArray = ValueArray[1, 1].ToString().Split();
                strbomexp = string.Join(" ", strCsvArray);
            }
            if (strbomexp.Equals("BOM Explosion Report"))
                isValid = true;
            else
                IsValid = false;
            return isValid;
        }


    }
}
