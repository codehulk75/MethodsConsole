using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Methods_Console
{

    class SetupSheetGenerator
    {
        Assembly ThisProgram;
        AssemblyName ThisProgramName;
        Version ThisProgramVersion;
        public int EndOfPage { get; private set; }
        public int HeaderLength { get; private set; }
        public BeiBOM Bom { get; private set; }
        public List<Ci2Parser> ProgramList { get; private set; }
        public List<string> LoadingList { get; private set; }
        public List<string> PassesList { get; set; }
        public int CurrentPageNumber { get; private set; }
        public int CurrentLineNumber { get; private set; }
        public string OutputDir { get; set; }
        public string FileName { get; private set; }
        public string FullPath { get; private set; }
        public SetupSheetGenerator(BeiBOM bom, List<Ci2Parser> ci2List, List<string> loadingInstructionsList)
        {
            CurrentLineNumber = 1;
            CurrentPageNumber = 1;
            EndOfPage = 63;
            HeaderLength = 6;
            ThisProgram = Assembly.GetEntryAssembly();
            ThisProgramName = ThisProgram.GetName();
            ThisProgramVersion = ThisProgramName.Version;
            OutputDir = @"C:\BaaN-DAT";
            PassesList = new List<string>(new string[] { "SMT 1", "SMT 2" });
            Bom = bom;
            ProgramList = ci2List;
            LoadingList = loadingInstructionsList;
            FileName = Bom.AssemblyName + "_" + Bom.Rev + ".rtf";
            FullPath = Path.Combine(OutputDir, FileName);
        }



        public void CreateBaanSetupSheet()
        {
            CreateRtfDoc();
            WriteProgramData();



            RtfDocWriteLastLine();
        }
        public void CreateSheet()
        {
            if (Bom.HasRouting)
                CreateBaanSetupSheet();
            else
                CreateAgileSetupSheet();
        }

        private void WriteNewProgramHeader(Ci2Parser ci2, string strLoadingInstructions)
        {
            string strNewProgramHeader = @"\par Program: " + ci2.ProgramName + @"\tab Date: " + ci2.DateCreated + @"\tab \tab " + ci2.MachineName + @"\tab Side: " + ci2.Pass + "\n"
                                        + @"\par Loading Instructions: " + strLoadingInstructions + "\n"
                                        + "\\par \n"
                                        + "\\par Part Number\\tab Description\n"
                                        + @"\par\tab Feeder\tab Qty\tab Reference Designators" + "\n"
                                        + @"\par _______________________________________________________________________________________________";
            using (StreamWriter writer = new StreamWriter(FullPath, true))
            {
                writer.WriteLine(strNewProgramHeader);
            }
            CurrentLineNumber = 14;
        }

        private void WriteMidProgramFooterHeader()
        {
            string strFooterHeader = "\n\\par\n"
                                    + @"\par " + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString() + @" \tab Page  " + CurrentPageNumber.ToString() + @" \tab Version " + ThisProgramVersion.ToString() + @"\page \tab Assembly:\tab " + Bom.AssemblyName + "\n"
                                    + @"\par \tab BOM Rev:\tab " + Bom.Rev + "\n"
                                    + @"\par \tab Listing Date:\tab " + Bom.DateOfListing + "\n"
                                    + "\\par _______________________________________________________________________________________________\n";
            using (StreamWriter writer = new StreamWriter(FullPath, true))
            {
                while(CurrentLineNumber < EndOfPage)
                {
                    writer.WriteLine(@"\par");
                    CurrentLineNumber++;
                }
                writer.WriteLine(strFooterHeader);
            }
            ++CurrentPageNumber;
            CurrentLineNumber = 7;
        }
        private void RtfDocWriteLastLine()
        {
            string rtfLastLine = @"\par " + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString() + @" \tab Page  " + CurrentPageNumber.ToString() + @" \tab Version " + ThisProgramVersion.ToString() + @"\page }}";
            using (StreamWriter writer = new StreamWriter(FullPath, true))
            {
                while(CurrentLineNumber < EndOfPage)
                {
                    writer.WriteLine("\\par");
                    ++CurrentLineNumber;
                }
                writer.WriteLine(rtfLastLine);
            }
        }
        private void CreateAgileSetupSheet()
        {

        }
        public void CreateRtfDoc()
        {
            string strRtfFirstLine = @"{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fnil\fcharset0 Courier New;}}\viewkind4\uc1\pard\tx2160\tx3600\tx4320\tx7200\margl360\margr360\margt360\margb360 {\f0\fs20";
            string strInitialHeader = @"\tab Assembly:\tab " + Bom.AssemblyName + "\n"
                       + @"\par \tab BOM Rev:\tab " + Bom.Rev + "\n"
                       + @"\par \tab Listing Date:\tab " + Bom.DateOfListing + "\n"
                       + @"\par _______________________________________________________________________________________________" + "\n";
            using (StreamWriter writer = new StreamWriter(FullPath, false))
            {
                writer.WriteLine(strRtfFirstLine);
                writer.WriteLine(strInitialHeader);
            }
            CurrentLineNumber = 6;
        }

        private void WriteProgramData()
        {
            int exportcounter = 1;
            CurrentLineNumber = 6;
            foreach (Ci2Parser export in ProgramList)
            {
                string strInstructions = LoadingList[0];
                
                if (export.Pass.Equals(PassesList[1]))
                    strInstructions = LoadingList[1];
                if(CurrentLineNumber > HeaderLength + 3)
                {
                    using (StreamWriter writer = new StreamWriter(FullPath, true))
                    {
                        ////fill page with \par before starting new program                    
                        while (CurrentLineNumber <= EndOfPage)
                        {
                            writer.WriteLine("\\par");
                            ++CurrentLineNumber;
                        }
                        /////                      
                    }
                    WriteMidProgramFooterHeader();
                }


                WriteNewProgramHeader(export, strInstructions);
                if (exportcounter == 1)
                    CurrentLineNumber = 13;
                List<string> lstPartInfo = PrepPartInfo(export);

                    foreach(string part in lstPartInfo)
                    {
                        if (CurrentLineNumber >= EndOfPage - HeaderLength + 3)
                            WriteMidProgramFooterHeader();
                        using (StreamWriter writer = new StreamWriter(FullPath, true))
                        {
                            writer.WriteLine(part);
                            CurrentLineNumber += 3;
                        }

                    }
                ++exportcounter;
            }
            MessageBox.Show("Done");               
        }

        private List<string> PrepPartInfo(Ci2Parser export)
        {
            List<string> info = new List<string>();
            string partnum = null;
            string desc = null;
            string slottrack = null;
            string feeder = null;
            string refdes = null;
            string strTemp = null;

            //  Using LINQ to sort the feeder list by slot then by track, then working with returned copy
            var items = from pair in export.Feedermap
                        orderby Convert.ToInt32(pair.Value[1]), pair.Value[2] ascending
                        select pair;


            foreach (var part in items)
            {
                partnum = part.Key;
                feeder = part.Value[0];
                
                List<Tuple<string, string, string, string, string>> tBomInfo = FindInBom(part.Key);
                if (tBomInfo.Count == 0)
                    desc = "*****PART NOT IN BOM*****";
                else
                    desc = tBomInfo[0].Item3;
                if (part.Value[0].Equals("PTF RR BTWN RAIL"))
                    slottrack = part.Value[2];
                else
                    slottrack = "SL " + part.Value[1] + " TK " + part.Value[2];
                refdes = export.Refdesmap[part.Key].ElementAt(0).Value[0];
                strTemp = @"\par " + part.Key + @"\tab " + desc + @"\tab " + feeder + "\n" + @"\par \tab " + slottrack + "\n" + @"\par";

                info.Add(strTemp);
            }
            
            return info;
        }

        private List<Tuple<string, string, string, string, string>> FindInBom(string partnum)
        {
            List<Tuple<string, string, string, string, string>> bomPart = new List<Tuple<string, string, string, string, string>>();
            foreach(var bomitem in Bom.Bom)
            {
                if (bomitem.Value.Item1.Equals(partnum))
                    bomPart.Add(bomitem.Value);
            }
            return bomPart;
        }
    }
}
