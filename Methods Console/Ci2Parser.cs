using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Text.RegularExpressions;

namespace Methods_Console
{
    class Ci2Parser : FileParser
    {
        //Properties
        public int CircuitCount { get; private set; }
        public string PanelLength { get; private set; }
        public string PanelWidth { get; private set;  }
        private HashSet<string> CircuitList { get; set; }
        public override string FileType { get; set; }
        public override string FileName { get; set; }
        public override string FullFilePath { get; set; }
        public string CustomerDbName { get; private set; }
        public string ProgramName { get; private set; }
        public string MainCircuitName { get; private set; }
        public string MachineName { get; set; }
        public string Pass { get; set; }
        public string OpCode { get; set; }
        public string DateCreated { get; private set; }
        public bool IsValid { get; private set; }
        public List<string> Lines { get; private set; }
        public Dictionary<string, Dictionary<string, List<string>>> Refdesmap { get; private set; }
        public Dictionary<string, List<string>> Feedermap { get; private set; }
        public Dictionary<string, Dictionary<string, List<string>>> BypassedRefDesMap { get; private set; }
        public Dictionary<string, List<string>> PlacementMap { get; private set;  }


        //Constructor
        public Ci2Parser(string path)
        {
            MachineName = "";
            Pass = "";
            DateCreated = "";
            FileType = Path.GetExtension(path).ToLower();
            FullFilePath = path;
            CircuitCount = 1;
            FileName = Path.GetFileName(FullFilePath);
            Refdesmap = new Dictionary<string, Dictionary<string, List<string>>>();
            Feedermap = new Dictionary<string, List<string>>();
            Lines = new List<string>();
            CircuitList = new HashSet<string>();
            BypassedRefDesMap = new Dictionary<string, Dictionary<string, List<string>>>();
            PlacementMap = new Dictionary<string, List<string>>();
            MainCircuitName = null;
            OpCode = null;
            LoadFile();
            ParseFile();
            SetValid();
        }

        public void ClearData()
        {
            CircuitCount = -1;
            PanelLength = null;
            PanelWidth = null;
            CircuitList.Clear();
            FileType = null;
            FileName = null;
            FullFilePath = null;
            CustomerDbName = null;
            ProgramName = null;
            MainCircuitName = null;
            MachineName = null;
            Pass = null;
            DateCreated = null;
            Lines.Clear();
            Refdesmap.Clear();
            Feedermap.Clear();
            BypassedRefDesMap.Clear();
            PlacementMap.Clear();
         }
        public void SetValid()
        {
            if (CircuitCount != -1 && !string.IsNullOrEmpty(PanelLength) && !string.IsNullOrEmpty(PanelWidth)
                && !string.IsNullOrEmpty(FileType) && !string.IsNullOrEmpty(FileName) && !string.IsNullOrEmpty(FullFilePath)
                && !string.IsNullOrEmpty(CustomerDbName) && !string.IsNullOrEmpty(ProgramName) && !string.IsNullOrEmpty(MainCircuitName)
                && !string.IsNullOrEmpty(PanelLength) && Refdesmap.Count > 0 && Feedermap.Count > 0 && PlacementMap.Count > 0)
            {
                IsValid = true;
            }
            else IsValid = false;
        }
        //Methods
        private void LoadFile()
        {
            string line;
            try
            {
                using (StreamReader sr = new StreamReader(FullFilePath))
                {
                    while ((line = sr.ReadLine()) != null)
                    {
                        Lines.Add(line);
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Error creating Ci2 File Parser object.\nMake sure it's a valid ci2 and try again.\nLoadFile()\n" + e.Message, "Ci2Parser.LoadFile()", MessageBoxButton.OK, MessageBoxImage.Error);
                ClearData();
                IsValid = false;
            }
        }

        private bool FindProgramName(string line)
        {
            Regex regProgName = new Regex("Product : \"(?<prog>[^\"]*)\"");
            Match match = regProgName.Match(line);
            if (match.Success)
            {
                ProgramName = match.Groups["prog"].Value;
                return true;
            }
            else
                return false;
        }

        private bool FindCustDb(string line)
        {
            Regex regcustdb = new Regex("Component Database Name");
            Match match = regcustdb.Match(line);
            if (match.Success)
            {
                string[] sep = new string[] { "\"" };
                string[] temp = line.Split(sep, StringSplitOptions.RemoveEmptyEntries);
                if (temp.Length > 2)
                {
                    CustomerDbName = Path.GetFileNameWithoutExtension(temp[3]);
                    return true;
                }
                else
                    return false;
            }
            else
                return false;
        }

        private Tuple<bool, bool> FindCircuitCount(string line)
        {
            //tuple values are: "found circuit?" true/false and "finished reading board section?" true/false
            if (line.Contains("Fiducial"))
            {
                if (CircuitList.Count > 0)
                    --CircuitCount;
                return Tuple.Create(false, true);
            }

            Regex reCircuit = new Regex("\\s+Circuit[\\s:\"]+Circuit");
            Regex reOffset = new Regex("\\s+Circuit[\\s:\"]+(Offset |\\d{ 1,2})");
            Match matchCircuit = reCircuit.Match(line);
            Match matchOffset = reOffset.Match(line);
            try
            {
                string[] sep = new string[] { "\"" };
                if (matchOffset.Success)
                {
                    string[] temp = line.Split(sep, StringSplitOptions.RemoveEmptyEntries);
                    if ( !temp[5].Equals(temp[9]) )
                    {
                        ++CircuitCount;
                        return Tuple.Create(true, false);
                    }
                    else
                        return Tuple.Create(false, false);
                }
                else if (matchCircuit.Success)
                {
                    string[] temp = line.Split(sep, StringSplitOptions.RemoveEmptyEntries);
                    if (CircuitList.Add(temp[1]))
                    {
                        ++CircuitCount;
                        return Tuple.Create(true, false);
                    }
                    else return Tuple.Create(false, false);
                        
                }
                else
                    return Tuple.Create(false, false);
            }
            catch (Exception e)
            {
                MessageBox.Show("The program had a problem parsing the board template section.\nReturning current count of " + CircuitCount.ToString() + " but could be wrong.\n" + e.Message, "Ci2Parser:FindCircuitCount()", MessageBoxButton.OK, MessageBoxImage.Warning);
                return Tuple.Create(false, true);
            }
        }

        private void PopulateFeederMap(string line)
        {
            if (line.Contains("Reject") || line.Contains("DUMP") || line.Contains("HANDPLACE") || line.Contains("NOBOM"))
                return;

            Regex reColonsNquotes = new Regex(@"^Slot\s:\s""([^""]+)""\s(?:""([^""]+)""\s:\s""([^""]*)""\s?)+");
            //
            //this regex parses out a feeder list line in a ci2. Capture Groups = 1- "Slot: " + slot, 2- field names, 3- field values
            // sauce *** new Regex(@"^Slot\s:\s""([^""]+)""\s(?:""([^""]+)""\s:\s""([^""]+)""\s?)+");
            //Number of captures in Groups 2+3 == 4 for a normal feeder and == 3 for a PTF part
            //
            string feeder = null;
            string compid = null;
            string slot = null;
            try
            {
                Match match = reColonsNquotes.Match(line);
                if (match.Success)
                {
                    feeder = match.Groups[3].Captures[0].Value;
                    compid = match.Groups[3].Captures[1].Value;
                    slot = match.Groups[3].Captures[2].Value;

                    if (slot.Equals("0"))
                    {
                        return;
                    }

                    else if (match.Groups[3].Captures[0].Value.Contains("PTF RR BTWN RAIL"))
                    {
                        if (Feedermap.ContainsKey(compid))
                        {
                            Feedermap[compid].Add(feeder);
                            Feedermap[compid].Add(slot);
                            Feedermap[compid].Add("Tray");
                            Feedermap[compid].Add(match.Groups[3].Captures[3].Value);//rotation field for tray parts
                        }
                        else
                        {
                            Feedermap.Add(compid, new List<string> { feeder, slot, "Tray", match.Groups[3].Captures[3].Value });
                        }
                    }
                    else
                    {
                        string track = match.Groups[3].Captures[3].Value;
                        string rotation = match.Groups[3].Captures[4].Value;
                        if (Feedermap.ContainsKey(compid))
                        {
                            Feedermap[compid].Add(feeder);
                            Feedermap[compid].Add(slot);
                            Feedermap[compid].Add(track);
                            Feedermap[compid].Add(rotation);
                        }
                        else
                        {
                            Feedermap.Add(compid, new List<string> { feeder, slot, track, rotation });
                        }
                    }
                    //string s = string.Format("Feeder added for part # {0}.\nFeeder : {1}\nSlot: {2}\nTrack: {3}\nRotation : {4}", compid, Feedermap[compid][0], Feedermap[compid][1], Feedermap[compid][2], Feedermap[compid][3]);
                    //MessageBox.Show(s);

                    //for (int ctr = 1; ctr < match.Groups.Count; ctr++)
                    //{
                    //    MessageBox.Show("\tGroup " +  ctr.ToString() + " : " + match.Groups[ctr].Value.ToString());
                    //    int captureCtr = 0;
                    //    foreach (Capture capture in match.Groups[ctr].Captures)
                    //    {
                    //        MessageBox.Show("Capture " + captureCtr.ToString() + " : " + capture.Value);
                    //        captureCtr++;
                    //    }
                    //}
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Error updating Feeder Map\nFeeder = " + feeder + "\nSlot = " + slot + "\nPN = " + compid, "PopulateFeederMap()", MessageBoxButton.OK, MessageBoxImage.Error);
                ClearData();
                IsValid = false;
            }
        }

        private void UpdateTray(string line)
        {
            Regex reTrayline = new Regex(@"^\s+Tray\s:\s""[^""]+""\s(?:""([^""]+)""\s:\s""([^""]*)""\s?)+");
            Match match = reTrayline.Match(line);
            string compid = null;
            string pallet = null;
            string stack = null;
            try
            {
                if (match.Success)
                {
                    compid = match.Groups[2].Captures[0].Value;
                    pallet = match.Groups[2].Captures[4].Value;
                    stack = match.Groups[2].Captures[5].Value == "1" ? "L" : "R";
                    int i = Feedermap[compid].FindLastIndex(x => x.Equals("Tray"));
                    if (i > -1)
                    {
                        Feedermap[compid][i] = "Tray " + pallet + stack;
                    }
                    else 
                    {
                        Feedermap[compid].Add(Feedermap[compid][0]);
                        Feedermap[compid].Add(Feedermap[compid][1]);
                        Feedermap[compid].Add("Tray " + pallet + stack);
                        Feedermap[compid].Add(Feedermap[compid][3]);
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Error updating Tray\nTray = " + pallet + stack + "\nPN = " + compid, "UpdateTray()", MessageBoxButton.OK, MessageBoxImage.Error);
                ClearData();
                IsValid = false;
            }

        }

        private void AddToRdMap(string pn, string rd, string circuitName)
        {
            try
            {
                if (Refdesmap.ContainsKey(pn))
                {
                    if (Refdesmap[pn].ContainsKey(circuitName))
                    {
                        Refdesmap[pn][circuitName].Add(rd);
                        Refdesmap[pn][circuitName].Sort();
                    }
                    else
                    {
                        Refdesmap[pn].Add(circuitName, new List<string> { rd });
                    }
                }
                else
                {
                    Dictionary<string, List<string>> dict = new Dictionary<string, List<string>>();
                    dict.Add(circuitName, new List<string> { rd });
                    Refdesmap.Add(pn, dict);                   
                }
            }
            catch (Exception e)
            {

                string s = string.Format("Error populating RefDes Map!\nParameters:\nPart Number : {0}\nRef Des : {1}\nCircuit Name : {2}\nError : {3}", pn, rd, circuitName, e.Message, MessageBoxButton.OK, MessageBoxImage.Error);
                MessageBox.Show(s, "AddToRdMap() Exception");
                ClearData();
                IsValid = false;
            }
        }

        private void AddToBypassedRdMap(string pn, string rd, string circuitName)
        {
            try
            {
                if (BypassedRefDesMap.ContainsKey(pn))
                {
                    if (BypassedRefDesMap[pn].ContainsKey(circuitName))
                    {
                        BypassedRefDesMap[pn][circuitName].Add(rd);
                    }
                    else
                    {
                        BypassedRefDesMap[pn].Add(circuitName, new List<string> { rd });
                    }
                }
                else
                {
                    Dictionary<string, List<string>> dict = new Dictionary<string, List<string>>();
                    dict.Add(circuitName, new List<string> { rd });
                    BypassedRefDesMap.Add(pn, dict);
                }
            }
            catch (Exception e)
            {
                string s = string.Format("Error populating BypassedRefDes Map!\nParameters:\nPart Number : {0}\nRef Des : {1}\nCircuit Name : {2}\nError : {3}", pn, rd, circuitName, e.Message, MessageBoxButton.OK, MessageBoxImage.Error);
                MessageBox.Show(s, "AddToBypassedRdMap() Exception");
                ClearData();
                IsValid = false;
            }
        }

        private void PopulateRefDesMap(string line)
        {
            Regex rePatternStep = new Regex(@"^\s+Step\s:\s""[^""]+""\s(?:""([^""]+)""\s:\s""([^""]*)""\s?)+");
            Match pStepMatch = rePatternStep.Match(line);

            if (pStepMatch.Success && pStepMatch.Groups[2].Captures[0].Value.Equals("Place"))
            {
                bool stepIsBypassed = pStepMatch.Groups[2].Captures[6].Value.Equals("1") ? true : false;
                string rd = pStepMatch.Groups[2].Captures[1].Value;
                string circuitName = pStepMatch.Groups[2].Captures[5].Value;
                string pn = "";
                if (PlacementMap.ContainsKey(rd))
                {
                    pn = PlacementMap[rd][0];
                }
                if (stepIsBypassed)
                    AddToBypassedRdMap(pn, rd, circuitName);
                else
                    AddToRdMap(pn, rd, circuitName);
            }
        }

        private void PopulatePlacementMap(string line)
        {
            Regex rePlacementStep = new Regex(@"^Placement\s:\s""([^""]+)""\s(?:""([^""]+)""\s:\s""([^""]*)""\s?)+");
            Match placeMatch = rePlacementStep.Match(line);
            string rd = null;
            string pn = null;
            string x = null;
            string y = null;
            string t = null;
            try
            {
                if (placeMatch.Success && placeMatch.Groups[3].Captures[3].Value.Equals("Place"))
                {
                    rd = placeMatch.Groups[1].Captures[0].Value;
                    pn = placeMatch.Groups[3].Captures[1].Value;
                    x = placeMatch.Groups[3].Captures[6].Value;
                    y = placeMatch.Groups[3].Captures[7].Value;
                    t = placeMatch.Groups[3].Captures[5].Value;
                    if (string.IsNullOrEmpty(MainCircuitName))
                    {
                        MainCircuitName = placeMatch.Groups[3].Captures[0].Value;
                    }
                    PlacementMap.Add(rd, new List<string> { pn, x, y, t });
                }
            }
            catch (Exception e)
            {
                string s = string.Format("Error populateing Placements Map!\n refdes : {0}\n pn: {1}\n x: {2}\n y : {3}\n theta : {4}\n Error:{5}", rd, pn, x, y, t, e.Message, MessageBoxButton.OK, MessageBoxImage.Error);
                MessageBox.Show(s, "PopulatePlacemnetMap() Exception");
                ClearData();
                IsValid = false;
            }

        }

        private void ParseFile()
        {
            Regex reBoard = new Regex(@"\s+Circuit[\s:""]+Board");
            Regex reLength = new Regex(@"X""\s:\s""\d+\.\d+");
            Regex reWidth = new Regex(@"Y""\s:\s""\d+\.\d+");
            Regex reDate = new Regex(@"^// Created on (\d{2}/\d{2}/\d{4})");
            List<string> widths = new List<string>();
            List<string> lengths = new List<string>();
            bool prognamefound = false;
            bool custdbfound = false;
            bool circuitcountfound = false;
            bool datefound = false;
            try
            {
                foreach (string line in Lines)
                {
                    if (datefound == false)
                    {
                        Match m = reDate.Match(line);
                        if (m.Success)
                        {
                            DateCreated = m.Groups[1].Value;
                            datefound = true;
                            continue;
                        }
                    }
                    if (prognamefound == false)
                    {
                        prognamefound = FindProgramName(line);
                        if (prognamefound == true)
                            continue;
                    }
                    if (custdbfound == false)
                    {
                        custdbfound = FindCustDb(line);
                        if (custdbfound == true)
                            continue;
                    }
                    if (circuitcountfound == false)
                    {
                        circuitcountfound = FindCircuitCount(line).Item2;
                        if (circuitcountfound == true)
                            continue;
                    }
                    if(line.StartsWith("Slot : "))
                    {
                        PopulateFeederMap(line);
                        continue;
                    }
                    if (line.StartsWith("   Tray :"))
                    {
                        UpdateTray(line);
                        continue;
                    }
                    if (line.StartsWith("Placement :"))
                    {
                        PopulatePlacementMap(line);
                        continue;
                    }
                    if (line.StartsWith("      Step :"))
                    {
                        PopulateRefDesMap(line);
                        continue;
                    }
                    Match matchBoardDimension = reBoard.Match(line);
                    if (matchBoardDimension.Success)
                    {
                        foreach(Match match in reLength.Matches(line))
                        {
                            lengths.Add(match.Value.Split('\"')[2]);
                        }
                        foreach (Match match in reWidth.Matches(line))
                        {
                            widths.Add(match.Value.Split('\"')[2]);
                        }
                    }
                }
                lengths.Sort();
                widths.Sort();
                PanelLength = lengths.Last();
                PanelWidth = widths.Last();
            }
            catch (Exception e)
            {
                MessageBox.Show("Error parsing ci2 file.\nMake sure it's a valid ci2 and try again.\nParseFile()\n" + e.Message, "Ci2Parser.ParseFile()", MessageBoxButton.OK, MessageBoxImage.Error);
                ClearData();
                IsValid = false;
                throw;
            }  
 
        }
    }
}
