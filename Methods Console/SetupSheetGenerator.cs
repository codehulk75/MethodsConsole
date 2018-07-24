using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Shapes;

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
        public Dictionary<string, List<string>> PassOneProgramRdMap { get; private set; } //key = pn, value = list of ref des's
        public Dictionary<string, List<string>> PassTwoProgramRdMap { get; private set; }
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
            EndOfPage = 62;
            HeaderLength = 6;
            ThisProgram = Assembly.GetEntryAssembly();
            ThisProgramName = ThisProgram.GetName();
            ThisProgramVersion = ThisProgramName.Version;
            OutputDir = @"C:\BaaN-DAT";
            PassesList = new List<string>(new string[] { "SMT 1", "SMT 2" });
            PassOneProgramRdMap = new Dictionary<string, List<string>>();
            PassTwoProgramRdMap = new Dictionary<string, List<string>>();
            Bom = bom;
            OpByRefCopy = bom.OpByRefDict;
            ProgramList = ci2List;
            LoadingList = loadingInstructionsList;
            FileName = Bom.AssemblyName + "_" + Bom.Rev + ".rtf";
            FullPath = System.IO.Path.Combine(OutputDir, FileName);
        }



        public void CreateBaanSetupSheet()
        {
            CreateRtfDoc();
            WriteProgramData();
            WriteMidProgramFooterHeader();
            WriteRefDesCountHeader();
            WriteRefDesCountData();
            WriteMidProgramFooterHeader();
            WriteDupPartsHeader();
            WriteMidProgramFooterHeader();
            WriteUnassocRefDesHeader();
            WritePartsInProgramNotInBom();
            WritePartsInBomNotInProgramHeader();
            WritePartsInBomNotInProgramData();
            RtfDocWriteLastLine();
            MessageBox.Show("Done");
        }
        private void WritePartsInBomNotInProgramData()
        {
            List<string> lines = new List<string>();
            ///sort BOM  by pn
            var sortedBom = from pair in Bom.Bom
                            orderby pair.Value.Item1 ascending
                            select pair;
            foreach (var bomItem in sortedBom)
            {
                if (PassOneProgramRdMap.ContainsKey(bomItem.Value.Item1) || PassTwoProgramRdMap.ContainsKey(bomItem.Value.Item1))
                {
                    List<string> progRds = new List<string>();
                    List<string> bomRds = new List<string>(bomItem.Value.Item4.Split(',').ToList());
                    if (PassOneProgramRdMap.ContainsKey(bomItem.Value.Item1))
                        progRds.AddRange(PassOneProgramRdMap[bomItem.Value.Item1]);
                    if(PassTwoProgramRdMap.ContainsKey(bomItem.Value.Item1))
                        progRds.AddRange(PassTwoProgramRdMap[bomItem.Value.Item1]);
                    foreach(string rd in progRds)
                    {
                        if (bomRds.Contains(rd))
                            bomRds.Remove(rd);
                    }
                    if(bomRds.Count > 0)
                    {
                        string pn = bomItem.Value.Item1;
                        string qty = bomRds.Count.ToString();
                        string rds = string.Join(",", bomRds);
                        if (bomItem.Value.Item5.Equals("0"))
                            rds = "Zero Qty";
                        if (rds.Length > 58)
                            rds = ChopRds(rds);
                        lines.Add(@"\par " + pn + @"\tab  " + qty + @" \tab " + rds);
                        lines.Add(@"\par ");
                    }                          
                }                  
                else
                {
                    string pn = bomItem.Value.Item1;
                    string qty = bomItem.Value.Item5;
                    string rds = bomItem.Value.Item4;
                    if (qty.Equals("0"))
                        rds = "Zero Qty";
                    if (rds.Length > 58)
                        rds = ChopRds(rds);
                    lines.Add(@"\par " + pn + @"\tab  " + qty + @" \tab " + rds);
                    lines.Add(@"\par ");
                }
            }
            ///Write results
            ///
            foreach (string line in lines)
            {
                if (CurrentLineNumber >= EndOfPage)
                    WriteMidProgramFooterHeader();
                using (StreamWriter writer = new StreamWriter(FullPath, true))
                {
                    writer.WriteLine(line);
                }
                ++CurrentLineNumber;
            }
        }
        private void WritePartsInBomNotInProgramHeader()
        {
            if (CurrentLineNumber >= EndOfPage)
                WriteMidProgramFooterHeader();
            using (StreamWriter writer = new StreamWriter(FullPath, true))
            {
                writer.WriteLine("\\par Parts in BOM, not in Program\n"
                    + "\\par\n\\par Part Number\\tab Qty\\tab Reference Designators\n"
                    + "\\par _______________________________________________________________________________________________\n");
                CurrentLineNumber += 4;
            }
        }
        private void WriteUnassocRefDesHeader()
        {
            if (CurrentLineNumber >= EndOfPage)
                WriteMidProgramFooterHeader();
            using (StreamWriter writer = new StreamWriter(FullPath, true))
            {
                writer.WriteLine("\\par \\tab UNASSOCIATED REFERENCE DESIGNATORS \n\\par\n"
                    + "\\par Parts in Program, not in BOM\n"
                    +"\\par \n\\par Part Number\\tab Qty\\tab Reference Designators\n"
                    + "\\par _______________________________________________________________________________________________\n");
                CurrentLineNumber += 6;
            }
        }
        private void WritePartsInProgramNotInBom()
        {
            
            List<string> lines = new List<string>();
            ///populate any orphan program info from 1st pass programs
            foreach(var part in PassOneProgramRdMap)
            {
                List<Tuple<string, string, string, string, string>> bomParts = FindInBom(part.Key);
                if(bomParts.Count == 0)
                {
                    string pn = part.Key;
                    string qty = part.Value.Count.ToString();
                    string rds = string.Join(",", part.Value);
                    if (rds.Length > 58)
                        rds = ChopRds(rds);
                    lines.Add(@"\par " + part.Key + @"\tab  " + part.Value.Count.ToString() + @" \tab " + rds);
                    lines.Add(@"\par ");
                }
                else
                {
                    List<string> progRds = new List<string>(part.Value);
                    List<string> bomRds = new List<string>();
                    foreach(var entry in bomParts)
                    {
                        bomRds.AddRange(entry.Item4.Split(',').ToList());
                    }
                    foreach(string rd in bomRds)
                    {
                        if (progRds.Contains(rd))
                            progRds.Remove(rd);
                    }
                    if(progRds.Count > 0)
                    {
                        string rds = string.Join(",", progRds);
                        if (rds.Length > 58)
                            rds = ChopRds(rds);
                        lines.Add(@"\par " + part.Key + @"\tab  " + progRds.Count.ToString() + @" \tab " + rds);
                        lines.Add(@"\par ");
                    }
                }
            }
            ///repeat for 2nd pass programs
            foreach (var part in PassTwoProgramRdMap)
            {
                List<Tuple<string, string, string, string, string>> bomParts = FindInBom(part.Key);
                if (bomParts.Count == 0)
                {
                    string pn = part.Key;
                    string qty = part.Value.Count.ToString();
                    string rds = string.Join(",", part.Value);
                    if (rds.Length > 58)
                        rds = ChopRds(rds);
                    lines.Add(@"\par " + part.Key + @"\tab  " + part.Value.Count.ToString() + @" \tab " + rds);
                    lines.Add(@"\par ");
                }
                else
                {
                    List<string> progRds = new List<string>(part.Value);
                    List<string> bomRds = new List<string>();
                    foreach (var entry in bomParts)
                    {
                        bomRds.AddRange(entry.Item4.Split(',').ToList());
                    }
                    foreach (string rd in bomRds)
                    {
                        if (progRds.Contains(rd))
                            progRds.Remove(rd);
                    }
                    if (progRds.Count > 0)
                    {
                        string rds = string.Join(",", progRds);
                        if (rds.Length > 58)
                            rds = ChopRds(rds);
                        lines.Add(@"\par " + part.Key + @"\tab  " + progRds.Count.ToString() + @" \tab " + rds);
                        lines.Add(@"\par ");
                    }
                }
            }
            ///write results
            foreach (string line in lines)
            {
                if (CurrentLineNumber >= EndOfPage)
                    WriteMidProgramFooterHeader();
                using (StreamWriter writer = new StreamWriter(FullPath, true))
                {
                    writer.WriteLine(line);
                }
                ++CurrentLineNumber;
            }
            using (StreamWriter writer = new StreamWriter(FullPath, true))
            {
                writer.WriteLine("\\par");
            }

        }
        private string ChopRds(string rds)
        {
            string strRds = null;
            List<string> refdesInput = new List<string>(rds.Split(','));
            List<string> rdOutput = new List<string>();
            string strTempLine = "";
            ///first ref des line
            for (int i = 0; refdesInput.Count > 0 && strTempLine.Length + refdesInput[0].Length + 1 < 58; ++i)
            {
                strTempLine = strTempLine + refdesInput[0] + ',';
                refdesInput.RemoveAt(0);
            }
            if (refdesInput.Count == 0)
                strTempLine = strTempLine.TrimEnd(',');
            rdOutput.Add(strTempLine);
            //rest of ref des lines
            if (refdesInput.Count > 0)
            {
                for (int i = 0; refdesInput.Count > 0; ++i)
                {
                    strTempLine = "";
                    for (int j = 0; refdesInput.Count > 0 && strTempLine.Length + refdesInput[0].Length + 1 < 58; ++j)
                    {
                        strTempLine += refdesInput[0] + ',';
                        refdesInput.RemoveAt(0);
                    }
                    if (refdesInput.Count == 0)
                        strTempLine = strTempLine.TrimEnd(',');
                    strTempLine += "\n";
                    rdOutput.Add(@"\par \tab \tab " + strTempLine);
                }
            }
            strRds = string.Join("\n", rdOutput);
            return strRds;
        }
        private void WriteDupPartsHeader()
        {
            ///LOOK INTO WHAT HE WAS TRYING TO DO HERE...(DUP CI2 REF DES'S???
            ///
            if (CurrentLineNumber >= EndOfPage)
                WriteMidProgramFooterHeader();
            using (StreamWriter writer = new StreamWriter(FullPath, true))
            {
                writer.WriteLine("\\par No Duplicate Parts in report! \n\\par _______________________________________________________________________________________________\n");
                CurrentLineNumber += 2;
            }
        }

        private void WriteRefDesCountData()
        {
            Regex reFirstPass = new Regex(@"(smt 1(?!.*inspect)|smt first)", RegexOptions.IgnoreCase);
            Regex reSecondPass = new Regex(@"(smt 2(?!.*inspect)|smt second)", RegexOptions.IgnoreCase);
            Dictionary<string, KeyValuePair<string, List<Tuple<string, string, string>>>> BomItemsByRouteStep = PopulateRouteBom();
            int iBomTotal= 0, iProgTotal = 0;
            ///Capture the parts not assigned to a route step, then remove them from the BomItems, write that section last
            KeyValuePair<string, List<Tuple<string, string, string>>> NoRouteStepParts = BomItemsByRouteStep["0"];
            BomItemsByRouteStep.Remove("0");
            foreach (var routeStep in BomItemsByRouteStep)
            {
                HashSet<string> hMasterPartsList = new HashSet<string>();
                
                string strRouteStepDesc = routeStep.Value.Key;
                if (reFirstPass.Match(strRouteStepDesc).Success)
                    strRouteStepDesc = "SMT 1ST PASS";
                else if (reSecondPass.Match(strRouteStepDesc).Success)
                    strRouteStepDesc = "SMT 2ND PASS";
                List<Tuple<string, string, string>> partInfo = routeStep.Value.Value.OrderBy(x => x.Item1).ToList();
                foreach (var part in partInfo)
                {
                    hMasterPartsList.Add(part.Item1);
                }
                foreach(var export in ProgramList)
                {
                    if (export.OpCode.Equals(routeStep.Key))
                    {
                        foreach(var part in export.Refdesmap)
                        {
                            hMasterPartsList.Add(part.Key);
                        }
                    }
                }
                List<string> lstMasterPartList = new List<string>(hMasterPartsList.ToList());
                lstMasterPartList.Sort();
                if (CurrentLineNumber >= EndOfPage)
                    WriteMidProgramFooterHeader();
                using (StreamWriter writer = new StreamWriter(FullPath, true))
                {
                    writer.WriteLine(@"\par " + strRouteStepDesc + "\n" + @"\par Part Number\tab BOM\tab Program"+"\n" + "\\par ");
                    CurrentLineNumber += 3;
                }
                foreach(string part in lstMasterPartList)
                {
                    string strProgramQty = "0";
                    string strError = "";
                    Tuple<string, string, string> tBomPart = partInfo.First(x => x.Item1.Equals(part));
                    foreach(Ci2Parser export in ProgramList)
                    {
                        if(export.OpCode.Equals(routeStep.Key))
                        {
                            if (export.Refdesmap.ContainsKey(part))
                            {
                                List<string> pRds = new List<string>(export.Refdesmap[part].ElementAt(0).Value.OrderBy(x => x));
                                List<string> bRds = new List<string>(tBomPart.Item2.Split(',').OrderBy(x => x));
                                strProgramQty = pRds.Count.ToString();
                                if(!pRds.SequenceEqual(bRds))
                                {
                                    strError = "<--Ref Des Mismatch!!";
                                }
                            }
                        }
                    }
                    int ibomqty = Convert.ToInt32(Math.Round(Convert.ToDouble(tBomPart.Item3)));
                    int iprogqty = Convert.ToInt32(strProgramQty);
                    if (ibomqty == 0)
                        ++iBomTotal;
                    else
                        iBomTotal += ibomqty;
                    iProgTotal += iprogqty;
                    if (ibomqty == 0 && iprogqty > 0)
                        strError = "<--Part not in BOM!!";
                    else if (ibomqty > 0 && iprogqty == 0 && routeStep.Value.Key.Contains("SMT"))
                        strError = "<--Hand Place";
                    else if (ibomqty != iprogqty && routeStep.Value.Key.Contains("SMT"))
                        strError = "<--Count";
                    if (CurrentLineNumber >= EndOfPage)
                        WriteMidProgramFooterHeader();
                    using (StreamWriter writer = new StreamWriter(FullPath, true))
                    {
                        writer.WriteLine(@"\par " + part + @"\tab  " + ibomqty.ToString() + " \\tab  " + strProgramQty + " \\tab " + strError + " ");
                    }
                    ++CurrentLineNumber;
                }
                using (StreamWriter writer = new StreamWriter(FullPath, true))
                {
                    writer.WriteLine("\\par ");
                }
                ++CurrentLineNumber;
                
            }
            //write the zero qyts.......
            if (CurrentLineNumber >= EndOfPage)
                WriteMidProgramFooterHeader();
            using (StreamWriter writer = new StreamWriter(FullPath, true))
            {
                writer.WriteLine("\\par " + NoRouteStepParts.Key + "\n" + @"\par Part Number\tab BOM\tab Program" + "\n" + "\\par ");
                CurrentLineNumber += 3;
            }
            NoRouteStepParts.Value.Sort();
            foreach (var part in NoRouteStepParts.Value)
            {

                int ibomqty = Convert.ToInt32(Math.Round(Convert.ToDouble(part.Item3)));
                int iprogqty = 0;
                ++iBomTotal;
                if (CurrentLineNumber >= EndOfPage)
                    WriteMidProgramFooterHeader();
                using (StreamWriter writer = new StreamWriter(FullPath, true))
                {
                    writer.WriteLine(@"\par " + part.Item1 + @"\tab  " + ibomqty.ToString() + " \\tab  " + iprogqty.ToString() + " \\tab ");
                }
                ++CurrentLineNumber;
            }
            ///write the total counts
            ///
            if (CurrentLineNumber >= EndOfPage - 5)
                WriteMidProgramFooterHeader();
            using (StreamWriter writer = new StreamWriter(FullPath, true))
            {
                writer.WriteLine("\\par\n\\par TOTAL\\tab  " + iBomTotal + @" \tab  " + iProgTotal + " \n\\par\n\\par Error Key:\n" + @"\par C = refdes count; R = refdes mismatch" + "\n\\par");
            }
            CurrentLineNumber += 5;
        }

        private Dictionary<string, KeyValuePair<string, List<Tuple<string, string, string>>>>  PopulateRouteBom()
        {
            ///Go through all the routing steps and keep the ones that have Ops assigned to them
            ///key = routeOpNumber , value = keyvalpair(key = routeOpDesc, value = List of Tuple(partnum, refDesString, bomqty)
            Dictionary<string, KeyValuePair<string, List<Tuple<string, string, string>>>> RouteMap = new Dictionary<string, KeyValuePair<string, List<Tuple<string, string, string>>>>();
            KeyValuePair<string, List<Tuple<string, string,string>>> zeroqtyStep = new KeyValuePair<string, List<Tuple<string, string,string>>>("Parts not assigned to a routing step.", new List<Tuple<string, string, string>>());
            RouteMap.Add("0", zeroqtyStep);
            foreach (var routeStep in Bom.RouteList)
            {
                string[] routeinfo = routeStep.Split(':');
                KeyValuePair<string, List<Tuple<string, string, string>>> descpns = new KeyValuePair<string, List<Tuple<string, string, string>>>(routeinfo[2], new List<Tuple<string, string, string>>());
                RouteMap.Add(routeinfo[0], descpns);
            }
            foreach(var bomline in Bom.Bom)
            {
                ///add partnum and ref des if bomline operation matches RouteMap step
                RouteMap[bomline.Value.Item2].Value.Add(new Tuple<string, string, string>(bomline.Value.Item1, bomline.Value.Item4, bomline.Value.Item5));
            }
            ///Before returning, remove the route steps with empty Lists in the key value pair
            ///
            List<string> routekeys = RouteMap.Keys.ToList();
            foreach(var key in routekeys)
            {
                if (RouteMap[key].Value.Count < 1)
                    RouteMap.Remove(key);
            }
            return RouteMap;
        }

        private void WriteRefDesCountHeader()
        {
            string strRefDesCountHeader = @"\par \tab REFERENCE DESIGNATOR COUNT\n\par \n";
            using (StreamWriter writer = new StreamWriter(FullPath, true))
            {
                writer.WriteLine(strRefDesCountHeader);
            }
            CurrentLineNumber = 14;
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
            CurrentLineNumber = 10;
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
            Regex reFirstPass = new Regex(@"(smt 1(?!.*inspect)|smt first)", RegexOptions.IgnoreCase);
            Regex reSecondPass = new Regex(@"(smt 2(?!.*inspect)|smt second)", RegexOptions.IgnoreCase);
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
            string strInstructions = null;
            foreach (Ci2Parser export in ProgramList)
            {
                
                if (export.Pass.Equals(PassesList[0]))
                {
                    strInstructions = LoadingList[0];
                    export.OpCode = GetOpCode(PassesList[0]);
                    if (string.IsNullOrEmpty(export.OpCode))
                    {
                        export.OpCode = GetOpCode(PassesList[0]);
                    }
                }
                else if (export.Pass.Equals(PassesList[1]))
                {
                    if (string.IsNullOrEmpty(export.OpCode))
                    {
                        export.OpCode = GetOpCode(PassesList[1]);
                    }
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
        }

        private List<string> PrepPartInfo(Ci2Parser export)
        {
            List<string> info = new List<string>();
            string partnum = null;
            string desc = null;
            int slot = 0;
            string track = null;
            string slottrack = null;
            string feeder = null;
            List<string> refdesInput = new List<string>();
            string strRefDesLines = null;
            string strQty = null;
            List<Tuple<string, string, string, string, string, string, KeyValuePair<int, string>>> tupList = new List<Tuple<string, string, string, string, string, string, KeyValuePair<int, string>>>();
    
           
            foreach(var feederItem in export.Feedermap)
            {
                string strTempLine = "";
                partnum = feederItem.Key;
                feeder = feederItem.Value[0];

                List<Tuple<string, string, string, string, string>> tBomInfo = FindInBom(feederItem.Key);
                if (tBomInfo.Count == 0)
                    desc = "*****PART NOT IN BOM*****";
                else
                    desc = tBomInfo[0].Item3;
                if (feederItem.Value[0].Equals("PTF RR BTWN RAIL"))
                {
                    slottrack = feederItem.Value[2];
                    slot = 75;
                    track = feederItem.Value[2];
                }                  
                else
                {
                    slottrack = "SL " + feederItem.Value[1] + " TK " + feederItem.Value[2];
                    slot = Convert.ToInt32(feederItem.Value[1]);
                    track = feederItem.Value[2];
                }
                    
                refdesInput = new List<string>(export.Refdesmap[feederItem.Key].ElementAt(0).Value);
                refdesInput.Sort();
                strQty = refdesInput.Count.ToString();
                for (int i = 0; refdesInput.Count > 0 && strTempLine.Length + refdesInput[0].Length + 1 < 54; ++i)
                {
                    if (export.Pass.Equals(PassesList[0]))
                    {
                        if (PassOneProgramRdMap.ContainsKey(partnum))
                            PassOneProgramRdMap[partnum].Add(refdesInput[0]);
                        else
                            PassOneProgramRdMap.Add(partnum, new List<string>(new string[] { refdesInput[0] }));
                    }
                    else if (export.Pass.Equals(PassesList[1]))
                    {
                        if (PassTwoProgramRdMap.ContainsKey(partnum))
                            PassTwoProgramRdMap[partnum].Add(refdesInput[0]);
                        else
                            PassTwoProgramRdMap.Add(partnum, new List<string>(new string[] { refdesInput[0] }));
                    }
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
                if (refdesInput.Count > 0)
                {
                    for (int i = 0; refdesInput.Count > 0; ++i)
                    {
                        strTempLine = "";
                        for (int j = 0; refdesInput.Count > 0 && strTempLine.Length + refdesInput[0].Length + 1 < 58; ++j)
                        {
                            if (export.Pass.Equals(PassesList[0]))
                            {
                                if (PassOneProgramRdMap.ContainsKey(partnum))
                                    PassOneProgramRdMap[partnum].Add(refdesInput[0]);
                                else
                                    PassOneProgramRdMap.Add(partnum, new List<string>(new string[] { refdesInput[0] }));
                            }
                            else if (export.Pass.Equals(PassesList[1]))
                            {
                                if (PassTwoProgramRdMap.ContainsKey(partnum))
                                    PassTwoProgramRdMap[partnum].Add(refdesInput[0]);
                                else
                                    PassTwoProgramRdMap.Add(partnum, new List<string>(new string[] { refdesInput[0] }));
                            }
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
                //If part number has a duplicate feeder...break it off and add it to the list
                KeyValuePair<int, string> kv = new KeyValuePair<int, string>(slot, track);
                tupList.Add(new Tuple<string, string, string, string, string, string, KeyValuePair<int,string>>(feederItem.Key, desc, feeder, slottrack, strQty, strRefDesLines, kv));
                if (feederItem.Value.Count > 4)
                {
                    for (int i = 4; i < feederItem.Value.Count; i += 4)
                    {
                        feeder = feederItem.Value[i];
                        if (feederItem.Value[i].Equals("PTF RR BTWN RAIL"))
                        {
                            slottrack = feederItem.Value[i + 2];
                            slot = 75;
                            track = feederItem.Value[i+2];
                        }

                        else
                        {
                            slottrack = "SL " + feederItem.Value[i + 1] + " TK " + feederItem.Value[i + 2];
                            slot = Convert.ToInt32(feederItem.Value[i+1]);
                            track = feederItem.Value[i+2];
                        }
                        KeyValuePair<int, string> kv2 = new KeyValuePair<int, string>(slot, track);    
                        tupList.Add(new Tuple<string, string, string, string, string, string, KeyValuePair<int, string>>(feederItem.Key, desc, feeder, slottrack, strQty, "~See Other Feeder", kv2));
                        
                    }
                }
            }
            ///sort by slot,track and prepare lines for printing
            var test = from tup in tupList
                       orderby tup.Item7.Key, tup.Item7.Value ascending
                       select tup;
            foreach (var sortedItems in test)
            {
                string temp  = @"\par " + sortedItems.Item1 + @"\tab " + sortedItems.Item2 + @"\tab " + sortedItems.Item3 + "\n" + @"\par \tab " + sortedItems.Item4 + @"\tab " + sortedItems.Item5 + @"\tab " + sortedItems.Item6 + @"\par";
                info.Add(temp);
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
