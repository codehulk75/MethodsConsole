using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;

namespace Methods_Console
{
    class SetupSheetParser : FileParser
    {
        public override string FileType { get; set; }
        public override string FileName { get; set; }
        public override string FullFilePath { get; set; }
        public bool IsValid { get; private set; }
        public List<string> Lines { get; private set; }

        //maybe make dupsrefs a map and append all the pns as a list as a value, so you can access the diff pn's later
        public SortedDictionary<string, List<string>> DuplicateRefs { get; private set; }

        public SortedDictionary<string, List<string>> FirstPassRefDesDict { get; private set; }
        public SortedDictionary<string, List<string>> SecondPassRefDesDict { get; private set; }
        public SortedDictionary<string, List<string>> FirstPassPnDict { get; private set; }
        public SortedDictionary<string, List<string>> FirstPassPnDups { get; private set; }
        public SortedDictionary<string, List<string>> SecondPassPnDict { get; private set; }
        public SortedDictionary<string, List<string>> SecondPassPnDups { get; private set; }


        public SetupSheetParser(string filepath)
        {
            FullFilePath = filepath;
            FileName = Path.GetFileName(FullFilePath);
            FileType = Path.GetExtension(FullFilePath).ToLower();
            FirstPassRefDesDict = new SortedDictionary<string, List<string>>();
            FirstPassPnDict = new SortedDictionary<string, List<string>>();
            FirstPassPnDups = new SortedDictionary<string, List<string>>();
            SecondPassRefDesDict = new SortedDictionary<string, List<string>>();
            SecondPassPnDict = new SortedDictionary<string, List<string>>();
            SecondPassPnDups = new SortedDictionary<string, List<string>>();
            DuplicateRefs = new SortedDictionary<string, List<string>>();
            Lines = new List<string>();
            LoadFile();
            ParseFile();
            SetValid();
        }

        private void LoadFile()
        {
            try
            {
                string line = null;
                using (StreamReader sr = new StreamReader(FullFilePath))
                {
                    while ((line = sr.ReadLine()) != null)
                        Lines.Add(line);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Error creating Setup Sheet Parser object.\nMake sure it's a valid setup sheet and try again.\n" + e.Message, "SetupSheetParser.LoadFile()", MessageBoxButton.OK, MessageBoxImage.Error);
                ClearData();
                IsValid = false;
            }
        }

        private void ClearData()
        {
            IsValid = false;
            FirstPassRefDesDict.Clear();
            FirstPassPnDict.Clear();
            FirstPassPnDups.Clear();
            SecondPassRefDesDict.Clear();
            SecondPassPnDict.Clear();
            SecondPassPnDups.Clear();
            Lines.Clear();
        }

        private void ParseFile()
        {
            Regex reProgramLine = new Regex(@"^\\par\s+Program:.*Date.*\\tab\s(\\tab\s)?([a-zA-z0-9\-_\s]+|Hand Place)\\tab\sSide:\s(SMT 1|SMT 2)");
            Regex rePartNumLine = new Regex(@"^\\par (\S+)\\tab");
            Regex reLocationLine = new Regex(@"( SL |Tray | SH |Hand Place)");
            Regex reMultiLineRds = new Regex(@"\\tab \\tab \\tab ");
            string machine = null;
            string pass = null;
            string tempPart = null;
            string firstRefs = null;
            bool bGetLocation = false;
            bool bRefDesSearch = false;

            foreach(string line in Lines)
            {
                string templine = line;
                string part = null;

                if (line.Contains("REFERENCE DESIGNATOR COUNT"))
                    break;
                if (bGetLocation)
                {
                    string[] sep = { @"\tab " };
                    List<string> LocAndRefInfo = line.Split(sep, StringSplitOptions.None).ToList();
                    Match match = reLocationLine.Match(line);
                    if (!match.Success)
                        continue;
                    List<string> lstLoc = GetLocAndRd(line);
                    ///Tack on slot location to end of pn info list
                    ///
                    if (pass.Equals("SMT 1"))
                        FirstPassPnDict[tempPart].Add(lstLoc.First());
                    else
                        SecondPassPnDict[tempPart].Add(lstLoc.First());
                    firstRefs = lstLoc.Last();
                    bRefDesSearch = true;
                }
                if (bRefDesSearch)
                {
                    if(bGetLocation)
                    {
                        bGetLocation = false;
                        List<string> fRefs = firstRefs.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries).ToList();
                        firstRefs = null;
                        if (pass.Equals("SMT 1"))
                        {
                            string fdrLocation = FirstPassPnDict[tempPart].Last();
                            foreach (var rd in fRefs)
                            {
                                if (FirstPassRefDesDict.ContainsKey(rd))
                                {
                                    if (!string.IsNullOrEmpty(rd) && !string.IsNullOrWhiteSpace(rd))
                                    {
                                        if (DuplicateRefs.ContainsKey(rd))
                                            DuplicateRefs[rd].Add(tempPart);
                                        else
                                            DuplicateRefs.Add(rd, new List<string> { FirstPassRefDesDict[rd][0], tempPart });
                                    }                                       
                                }                                  
                                else
                                    FirstPassRefDesDict[rd] = new List<string> { tempPart, pass, machine, fdrLocation };
                            }
                            continue;
                        }                            
                        else
                        {
                            string fdrLocation = SecondPassPnDict[tempPart].Last();
                            foreach (var rd in fRefs)
                            {
                                if (SecondPassRefDesDict.ContainsKey(rd))
                                {
                                    if (!string.IsNullOrEmpty(rd) && !string.IsNullOrWhiteSpace(rd))
                                    {
                                        if (DuplicateRefs.ContainsKey(rd))
                                            DuplicateRefs[rd].Add(tempPart);
                                        else
                                            DuplicateRefs.Add(rd, new List<string> { SecondPassRefDesDict[rd][0], tempPart });
                                    }
                                }                                  
                                else
                                    SecondPassRefDesDict[rd] = new List<string> { tempPart, pass, machine, fdrLocation };
                            }                           
                            continue;
                        }
                    }
                    else
                    {
                        Match matchRds = reMultiLineRds.Match(line);
                        if (!matchRds.Success && rePartNumLine.Match(line).Success)
                            bRefDesSearch = false;
                        else if (matchRds.Success)
                        {
                            List<string> lstRds = reMultiLineRds.Split(line).Last().Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries).ToList();
                            if (pass.Equals("SMT 1"))
                            {
                                string fdrLocation = FirstPassPnDict[tempPart].Last();
                                foreach (var rd in lstRds)
                                {
                                    if (FirstPassRefDesDict.ContainsKey(rd) && !string.IsNullOrWhiteSpace(rd))
                                    {
                                        if (!string.IsNullOrEmpty(rd) && !string.IsNullOrWhiteSpace(rd))
                                        {
                                            if (DuplicateRefs.ContainsKey(rd))
                                                DuplicateRefs[rd].Add(tempPart);
                                            else
                                                DuplicateRefs.Add(rd, new List<string> { FirstPassRefDesDict[rd][0], tempPart });
                                        }                                           
                                    }
                                    else
                                        FirstPassRefDesDict[rd] = new List<string> { tempPart, pass, machine, fdrLocation };
                                }
                                continue;
                            }
                            else
                            {
                                string fdrLocation = SecondPassPnDict[tempPart].Last();
                                foreach (var rd in lstRds)
                                {
                                    if (SecondPassRefDesDict.ContainsKey(rd) && !string.IsNullOrEmpty(rd))
                                    {
                                        if (!string.IsNullOrEmpty(rd) && !string.IsNullOrWhiteSpace(rd))
                                        {
                                            if (DuplicateRefs.ContainsKey(rd))
                                                DuplicateRefs[rd].Add(tempPart);
                                            else
                                                DuplicateRefs.Add(rd, new List<string> { SecondPassRefDesDict[rd][0], tempPart });
                                        }
                                    }
                                    else
                                        SecondPassRefDesDict[rd] = new List<string> { tempPart, pass, machine, fdrLocation };
                                }
                                continue;
                            }
                        }
                            
                    }                 
                }
                Match m = reProgramLine.Match(line);
                if (m.Success)
                {
                    ///Grab machine and pass whenever available
                    ///                
                    machine = m.Groups[2].Value;
                    pass = m.Groups[3].Value;
                    ///Strip hyphens and underscores to account for different naming styles
                    ///
                    string[] removechars = new string[] { "-", "_" };
                    foreach (string c in removechars)
                    {
                        machine = machine.Replace(c, string.Empty);
                    }
                    continue;
                }
                m = rePartNumLine.Match(line);
                if (m.Success)
                {
                    part = m.Groups[1].Value;
                    string feeder = GetFeederType(line);
                    if (pass.Equals("SMT 1"))
                    {
                        if (!FirstPassPnDict.ContainsKey(part))
                        {
                            FirstPassPnDict.Add(part, new List<string>(new string[] { feeder, pass, machine }));
                            tempPart = part;
                            bGetLocation = true;
                        }
                        else
                        {
                            if (!FirstPassPnDups.ContainsKey(part))
                                FirstPassPnDups.Add(part, new List<string>(new string[] { feeder, pass, machine }));
                            else
                                FirstPassPnDups[part].AddRange(new List<string>(new string[] { feeder, pass, machine }));
                        }
                            
                    }                      
                    else
                    {
                        if (!SecondPassPnDict.ContainsKey(part))
                        {
                            SecondPassPnDict.Add(part, new List<string>(new string[] { feeder, pass, machine }));
                            tempPart = part;
                            bGetLocation = true;
                        }
                        else
                        {
                            if (!SecondPassPnDups.ContainsKey(part))
                                SecondPassPnDups.Add(part, new List<string>(new string[] { feeder, pass, machine }));
                            else
                                SecondPassPnDups[part].AddRange(new List<string>(new string[] { feeder, pass, machine }));
                        }                           
                    }                         
                }
            }

        }

        private void SetValid()
        {
            if ((FirstPassPnDict.Count > 0 || SecondPassPnDict.Count > 0) && (FirstPassRefDesDict.Count > 0 || SecondPassRefDesDict.Count > 0))
                IsValid = true;
            else
                IsValid = false;
        }

        private string GetFeederType(string line)
        {
            string feeder = null;
            string[] sep = { @"\tab " };
            List<string> fields = line.Split(sep, StringSplitOptions.RemoveEmptyEntries).ToList();
            if (fields.Count == 3 && fields[2].Length > 3)
                feeder = fields[2];
            else
                feeder = "Feeder not found";
            return feeder;
        }

        private List<string> GetLocAndRd(string line)
        {
            string[] sep = { @"\tab " };
            List<string> list = line.Split(sep, StringSplitOptions.None).ToList();
            list.RemoveAt(0);
            if (string.IsNullOrEmpty(list[1]) || string.IsNullOrWhiteSpace(list[1]))
                list[1] = "qty not found";
            for (int i = 0; i < list.Count; ++i)
                list[i] = list[i].Trim();
            return list;
        }
    }
}
