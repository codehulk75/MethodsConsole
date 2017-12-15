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
    class BOMExplosionParser : FileParser
    {
        public override string FileType { get; set; }
        public override string FileName { get; set; }
        public override string FullFilePath { get; set; }
        public object[,] ValueArray { get; private set; }
        public bool hasRouting { get; private set; }
        private void ClearData()
        {
            FileType = null;
            FileName = null;
            FullFilePath = null;
            hasRouting = false;
            Array.Clear(ValueArray, 1, ValueArray.Length);
        }
        public BOMExplosionParser(string path, string fileExtension)
        {
            FullFilePath = path;
            FileName = Path.GetFileName(FullFilePath);
            FileType = fileExtension;
            hasRouting = false;
            if (FileType.Equals(".txt"))
            {
                if (IsValidBaanBOM())
                {
                    hasRouting = true;
                    ParseBaanBom();
                }

                else
                {
                    ClearData();
                    MessageBox.Show("Not a valid BAAN BOM!\nPlease check your bom file and try again.", "BAAN BOM Format Error");
                }
                    
            }
            else
            {
                ParseBomExplosion();
            }
        }

        private void ParseBaanBom()
        {
            MessageBox.Show("ParseBaanBom() will be done soon!", "lazy sry not sry");
        }

        private void ParseBomExplosion()
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
                ClearData();
                MessageBox.Show("Problem loading BOM. Please check the file and try again.\n" + e.Message, "ParseBomExplosion()");
            }

            //Determine whether it's a legit csv or excel bom explosion and send it to the correct method
            string strCsv = null;
            string strExcel = null;
            if (ValueArray[1,1] != null)
            {
                string[] strCsvArray = ValueArray[1, 1].ToString().Split();
                strCsv = string.Join(" ", strCsvArray);
            }
            if (ValueArray[1,2] != null)
            {
                string[] strExcelArray = ValueArray[1, 2].ToString().Split();
                strExcel = string.Join(" ", strExcelArray);
            }    
            if (FileType.Equals(".xls") && strExcel.Equals("BOM Explosion Report"))
                ProcessExcelBOM();
            else if (FileType.Equals(".csv") && strCsv.Equals("BOM Explosion Report"))
                 ProcessCsvBOM();
            else
            {
                ClearData();
                MessageBox.Show("Unrecognized file type. Please try again with BOM Explosion in .xls or .csv format, or a BAAN BOM in .txt format", "ParseBomExplosion()");
            }       
        }

        private void ProcessExcelBOM()
        {

            MessageBox.Show("Welcome to Excel BOM Processing!");
        }

        private void ProcessCsvBOM()
        {

            MessageBox.Show("ProcessCsvBOM() coming soon!", "Maybe...");
        }

        private bool IsValidBaanBOM()
        {
            bool isvalid = false;
            bool hasbomline = false;
            bool hasroutingline = false;
            
            string line;

            Regex bomline = new Regex(@"BILLS OF MATERIAL.*SINGLE LEVEL.*WITH BOM QUANTITIES");
            Regex routingline = new Regex(@"Routing Item");
            try
            {
                using (StreamReader sr = new StreamReader(FullFilePath))
                {
                    while ((line = sr.ReadLine()) != null)
                    {
                        if (bomline.Match(line).Success)
                            hasbomline = true;
                        if (routingline.Match(line).Success)
                            hasroutingline = true;
                    }
                }
                if (hasbomline && hasroutingline)
                    isvalid = true;
            }
            catch (Exception e)
            {
                MessageBox.Show("Error reading BAAN BOM. Could not load BOM.\n" + e.Message, "IsValidBaanBOM()");
            }
            return isvalid;
        }
    }
}
