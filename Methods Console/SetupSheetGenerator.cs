using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
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
        public Dictionary<string, string> OpByRefCopy { get; private set; }
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
            OpByRefCopy = bom.OpByRefDict;
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
        private void WriteHandPlaceHeader(string pass)
        {         
            string strNewProgramHeader = @"\par Program: " + pass + " " + "Hand Place" + @"\tab Date: " + ProgramList.First().DateCreated + @"\tab Hand Place\tab Side: " + pass + "\n"
                            + @"\par Loading Instructions: Hand Place\n"
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

        private void WriteHandPlaceOneSection()
        {
            List<string> lstHpOneLines = PrepHpInfo(PassesList[0]);
            WriteMidProgramFooterHeader();
            WriteHandPlaceHeader(PassesList[0]);
            CurrentLineNumber = 13;
            string[] arr = { "\n" };
            for (int j = 0; j < lstHpOneLines.Count; ++j)
            {
                string[] tempLines = lstHpOneLines[j].Split(arr, StringSplitOptions.RemoveEmptyEntries);
                List<string> lstPartLines = new List<string>(tempLines);
                if (CurrentLineNumber >= EndOfPage - 3)//3 here represents length of data.  will need to change to length of ref des data
                    WriteMidProgramFooterHeader();
                using (StreamWriter writer = new StreamWriter(FullPath, true))
                {
                    writer.WriteLine(lstPartLines[0]);
                    lstPartLines.RemoveAt(0);
                    writer.WriteLine(lstPartLines[0]);
                    lstPartLines.RemoveAt(0);
                    CurrentLineNumber += 2;
                }
                for (int i = 0; i < lstPartLines.Count; ++i)
                {
                    if (CurrentLineNumber >= EndOfPage)
                        WriteMidProgramFooterHeader();
                    using (StreamWriter writer = new StreamWriter(FullPath, true))
                    {
                        writer.WriteLine(lstPartLines[i]);
                    }
                    ++CurrentLineNumber;
                }
            }

        }
        private void WriteHandPlaceTwoSection()
        {
            List<string> lstHpOneLines = PrepHpInfo(PassesList[1]);
            WriteMidProgramFooterHeader();
            WriteHandPlaceHeader(PassesList[1]);
            CurrentLineNumber = 13;
            string[] arr = { "\n" };
            for (int j = 0; j < lstHpOneLines.Count; ++j)
            {
                string[] tempLines = lstHpOneLines[j].Split(arr, StringSplitOptions.RemoveEmptyEntries);
                List<string> lstPartLines = new List<string>(tempLines);
                if (CurrentLineNumber >= EndOfPage - 3)//3 here represents length of data.  will need to change to length of ref des data
                    WriteMidProgramFooterHeader();
                using (StreamWriter writer = new StreamWriter(FullPath, true))
                {
                    writer.WriteLine(lstPartLines[0]);
                    lstPartLines.RemoveAt(0);
                    writer.WriteLine(lstPartLines[0]);
                    lstPartLines.RemoveAt(0);
                    CurrentLineNumber += 2;
                }
                for (int i = 0; i < lstPartLines.Count; ++i)
                {
                    if (CurrentLineNumber >= EndOfPage)
                        WriteMidProgramFooterHeader();
                    using (StreamWriter writer = new StreamWriter(FullPath, true))
                    {
                        writer.WriteLine(lstPartLines[i]);
                    }
                    ++CurrentLineNumber;
                }
            }
        }
        private List<string> PrepHpInfo(string pass)
        {
            List<string> lstHPLines = new List<string>();
            Dictionary<string, string> dictHpParts = new Dictionary<string, string>();
            Dictionary<string, List<string>> dictHpPartInfo = new Dictionary<string, List<string>>();
            string strOpCode = GetOpCode(pass);
            Regex reOp = new Regex(@"\b" + strOpCode + @"\b:([^,]+|\b)");
            foreach(var entry in OpByRefCopy)
            {
                MatchCollection matches = reOp.Matches(entry.Value);
                foreach(Match match in matches)
                {
                    ///add to dictionary with pn as key and rd entry.Key as value, keep adding the refdes's
                    ///
                    if(dictHpParts.ContainsKey(entry.Key))
                        dictHpParts[entry.Key] += "," + match.Groups[1].Value;
                    else
                        dictHpParts[entry.Key] = match.Groups[1].Value;
                }
            }
            ///convert to part based dict and assign ref des's to parts 
            ///
            foreach(var entry in dictHpParts)
            {
                List<string> lstPns = new List<string>(entry.Value.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToArray());
                foreach(string part in lstPns)
                {
                    if(dictHpPartInfo.ContainsKey(part))
                        dictHpPartInfo[part][0] += "," + entry.Key;
                    else
                    {
                        dictHpPartInfo[part] = new List<string>();
                        dictHpPartInfo[part].Add(entry.Key);
                    }                      
                }
            }
            ///add to converted part dict the part description and bom qty
            ///
            foreach(var entry in dictHpPartInfo)
            {
                //Key = '<Findnum>:<Sequence>', Tuple.Item1=Part Number, Item2=Operation, Item3=Description, Item4=comma-separated ref des's, Item5=BOM qty
                //-- Item2 (operation) is null for Agile BOMs
                List<Tuple<string, string, string, string, string>> partInfo = FindInBom(entry.Key);
                if (partInfo.Count == 0)
                {
                    MessageBox.Show(string.Format("Part: {0} not found in BOM.", entry.Key), "SetupSheetGenerator:PrepHpPartInfo()", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                foreach(var tInfo in partInfo)
                {
                    if (tInfo.Item2.Equals(strOpCode))
                    {
                        dictHpPartInfo[entry.Key].Add(tInfo.Item3);
                        dictHpPartInfo[entry.Key].Add(tInfo.Item5);
                    }
                }

            }
            ///Prepare list of HP Lines using 
            ///
            lstHPLines = FormatPartInfo(dictHpPartInfo);
            lstHPLines.Sort();
            return lstHPLines;
        }

        private List<string> FormatPartInfo(Dictionary<string, List<string>> dictPartInfo)
        {
            List<string> lstHpLines = new List<string>();
            string partnum = null;
            string desc = null;
            string strTemp = null;
            string strRefDesLines = null;
            string strQty = null;
            foreach (var entry in dictPartInfo)
            {
                string strTempLine = "";
                partnum = entry.Key;
                desc = entry.Value[1];
                strQty = entry.Value[2];
                List<string> lstRds = new List<string>(entry.Value[0].Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToArray());
                lstRds.Sort();
                ///first ref des line
                for (int i = 0; lstRds.Count > 0 && strTempLine.Length + lstRds[0].Length + 1 < 54; ++i)
                {
                    strTempLine = strTempLine + lstRds[0] + ',';
                    lstRds.RemoveAt(0);
                }
                if (lstRds.Count == 0)
                    strTempLine = strTempLine.TrimEnd(',');
                strTempLine += "\n";
                strRefDesLines = strTempLine;
                //rest of ref des lines
                if (lstRds.Count > 0)
                {
                    for (int i = 0; lstRds.Count > 0; ++i)
                    {
                        strTempLine = "";
                        for (int j = 0; lstRds.Count > 0 && strTempLine.Length + lstRds[0].Length + 1 < 58; ++j)
                        {
                            strTempLine += lstRds[0] + ',';
                            lstRds.RemoveAt(0);
                        }
                        if (lstRds.Count == 0)
                            strTempLine = strTempLine.TrimEnd(',');
                        strTempLine += "\n";
                        strRefDesLines += @"\par \tab \tab \tab " + strTempLine;
                    }
                }
                strTemp = @"\par " + partnum + @"\tab " + desc + "\\tab \n" + @"\par \tab Hand Place\tab " + strQty + @"\tab " + strRefDesLines + "\n" + @"\par";

                lstHpLines.Add(strTemp);
            }

            return lstHpLines;
        }
        private string GetOpCode(string strPass)
        {
            string opcode = null;
            Regex reFirstPass = new Regex(@"(smt 1|smt first)", RegexOptions.IgnoreCase);
            Regex reSecondPass = new Regex(@"(smt 2|smt second)", RegexOptions.IgnoreCase);
            foreach (string op in Bom.RouteList)
            {
                if (strPass.Equals(PassesList[0]))
                {
                    if (reFirstPass.Match(op).Success)
                    {
                        opcode = op.Split(new char[] { ':' }, StringSplitOptions.RemoveEmptyEntries)[0];
                        break;
                    }
                }
                else if (strPass.Equals(PassesList[1]))
                {
                    if (reSecondPass.Match(op).Success)
                    {
                        opcode = op.Split(new char[] { ':' }, StringSplitOptions.RemoveEmptyEntries)[0];
                        break;
                    }
                }

            }
            if (string.IsNullOrEmpty(opcode))
                MessageBox.Show("Could not find opcode for "+strPass, "GetOpCode() Error", MessageBoxButton.OK, MessageBoxImage.Error);
            return opcode;
        }
        private void WriteProgramData()
        {
            int exportcounter = 1;
            CurrentLineNumber = 6;
            bool bFirstPassHandWritten = false;
            foreach (Ci2Parser export in ProgramList)
            {
                string strInstructions = LoadingList[0];
               
                if (export.Pass.Equals(PassesList[1]))
                {
                    strInstructions = LoadingList[1];
                    if(bFirstPassHandWritten == false)
                    {
                        WriteHandPlaceOneSection();
                        bFirstPassHandWritten = true;
                    }
                   
                }
                    
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
                string[] arr = { "\n" };
                for(int j=0; j < lstPartInfo.Count; ++j)
                {
                    string[] tempLines = lstPartInfo[j].Split(arr, StringSplitOptions.RemoveEmptyEntries);
                    List<string> lstPartLines = new List<string>(tempLines);
                    if (CurrentLineNumber >= EndOfPage - 3)//3 here represents length of data.  will need to change to length of ref des data
                        WriteMidProgramFooterHeader();
                    using (StreamWriter writer = new StreamWriter(FullPath, true))
                    {
                        writer.WriteLine(lstPartLines[0]);
                        lstPartLines.RemoveAt(0);
                        writer.WriteLine(lstPartLines[0]);
                        lstPartLines.RemoveAt(0);
                        CurrentLineNumber += 2;
                    }
                    for(int i = 0; i < lstPartLines.Count; ++i)
                    {
                        if (CurrentLineNumber >= EndOfPage)
                            WriteMidProgramFooterHeader();
                        using (StreamWriter writer = new StreamWriter(FullPath, true))
                        {
                            writer.WriteLine(lstPartLines[i]);
                        }
                        ++CurrentLineNumber;
                    }
                }
                ++exportcounter;
            }
            WriteHandPlaceTwoSection();
            MessageBox.Show("Done");               
        }

        private List<string> PrepPartInfo(Ci2Parser export)
        {
            List<string> info = new List<string>();
            string partnum = null;
            string desc = null;
            string slottrack = null;
            string feeder = null;
            List<string> refdesInput = new List<string>();
            string strRefDesLines = null;
            string strTemp = null;
            string strQty = null;        
           
            //  Using LINQ to sort the feeder list by slot then by track, then working with returned copy
            var items = from pair in export.Feedermap
                        orderby Convert.ToInt32(pair.Value[1]), pair.Value[2] ascending
                        select pair;
            foreach (var part in items)
            {
                string strTempLine = "";
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
                refdesInput = export.Refdesmap[part.Key].ElementAt(0).Value;
                refdesInput.Sort();
                strQty = refdesInput.Count.ToString();
                ///first ref des line
                for (int i = 0; refdesInput.Count > 0 && strTempLine.Length + refdesInput[0].Length + 1 < 54; ++i)
                {
                    if (OpByRefCopy.ContainsKey(refdesInput[0]))
                        OpByRefCopy.Remove(refdesInput[0]);
                    strTempLine = strTempLine + refdesInput[0] + ',';
                    refdesInput.RemoveAt(0);
                }
                if (refdesInput.Count == 0)
                    strTempLine = strTempLine.TrimEnd(',');
                strTempLine += "\n";
                strRefDesLines = strTempLine;
                //rest of ref des lines
                if( refdesInput.Count > 0)
                {
                    for (int i = 0; refdesInput.Count > 0; ++i)
                    {
                        strTempLine = "";
                        for (int j = 0; refdesInput.Count > 0 && strTempLine.Length + refdesInput[0].Length + 1 < 58; ++j)
                        {
                            if (OpByRefCopy.ContainsKey(refdesInput[0]))
                                OpByRefCopy.Remove(refdesInput[0]);
                            strTempLine += refdesInput[0] + ',';
                            refdesInput.RemoveAt(0);
                        }
                        if (refdesInput.Count == 0)
                            strTempLine = strTempLine.TrimEnd(',');
                        strTempLine += "\n";
                        strRefDesLines += @"\par \tab \tab \tab " + strTempLine;
                    }
                }

                strTemp = @"\par " + part.Key + @"\tab " + desc + @"\tab " + feeder + "\n" + @"\par \tab " + slottrack  + @"\tab " + strQty + @"\tab " + strRefDesLines + "\n" + @"\par";

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
