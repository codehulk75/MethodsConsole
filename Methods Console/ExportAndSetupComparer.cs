using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;

namespace Methods_Console
{
    class ExportAndSetupComparer
    {
        public SetupSheetParser OldParser { get; private set; }
        public SetupSheetParser NewParser { get; private set; }
        public List<Ci2Parser> OldCi2List { get; private set; }
        public List<Ci2Parser> NewCi2List { get; private set; }
        public int ErrorCount { get; private set; }
        public bool Valid { get; private set; }

        private string ReportBanner;
        private string ReportHeader;
        private List<string> botCi2Dups;
        private List<string> topCi2Dups;
        private List<string> lstBotCi2Deleted;
        private List<string> lstBotCi2Added;
        private List<string> lstTopCi2Deleted;
        private List<string> lstTopCi2Added;
        private List<string> lstCi2JumpedBotToTop;
        private List<string> lstCi2JumpedTopToBot;
        private List<string> lstSetupRdAdds;
        private List<string> lstSetupRdDeletes;
        private List<string> lstSetupRdChangedPartNum;
        private List<string> lstSetupPassChanges;
        private List<string> lstSetupMachineChanges;
        private List<string> lstSetupSlotChanges;
        private List<string> lstFeederChangesFirstPass;
        private List<string> lstFeederChangesSecondPass;
        private int iExportErrorCount;
        private bool bBotCi2sMatch;
        private bool bTopCi2sMatch;
        private int iSetupErrorCount;

        public ExportAndSetupComparer(SetupSheetParser oldParser, SetupSheetParser newParser, List<Ci2Parser> oldCi2s, List<Ci2Parser> newCi2s)
        {
            Valid = false;
            ErrorCount = 0;
            OldParser = oldParser;
            NewParser = newParser;
            OldCi2List = oldCi2s;
            NewCi2List = newCi2s;
            ReportBanner = GetReportBanner();
            ReportHeader = GetReportHeader();
            botCi2Dups = new List<string>();
            topCi2Dups = new List<string>();
            lstBotCi2Deleted = new List<string>();
            lstBotCi2Added = new List<string>();
            lstTopCi2Deleted = new List<string>();
            lstTopCi2Added = new List<string>();
            lstCi2JumpedBotToTop = new List<string>();
            lstCi2JumpedTopToBot = new List<string>();
            iExportErrorCount = 0;
            bBotCi2sMatch = false;
            bTopCi2sMatch = false;
            lstSetupRdAdds = new List<string>();
            lstSetupRdDeletes = new List<string>();
            lstSetupRdChangedPartNum = new List<string>();
            lstSetupPassChanges = new List<string>();
            lstSetupMachineChanges = new List<string>();
            lstSetupSlotChanges = new List<string>();
            lstFeederChangesFirstPass = new List<string>();
            lstFeederChangesSecondPass = new List<string>();
            iSetupErrorCount = 0;
        }

        public void SaveResults()
        {
            if (ErrorCount == 0)
            {
                MessageBoxResult result = MessageBox.Show("The are no errors in the compare report.  Do you want to save it, anyway?",
                    "Setup and Program Check is Clean", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.No)
                    return;
                else
                {
                    Microsoft.Win32.SaveFileDialog savedlg = new Microsoft.Win32.SaveFileDialog();
                    savedlg.Filter = "Text Files (*.txt)|*.txt";
                    savedlg.DefaultExt = "txt";
                    savedlg.FileName = "Setup Sheet Compare";
                    savedlg.AddExtension = true;
                    savedlg.InitialDirectory = Path.GetDirectoryName(NewParser.FullFilePath);
                    savedlg.Title = "Save Report";
                    bool? saveResult = savedlg.ShowDialog();
                    if (saveResult == true)
                    {
                        using (StreamWriter writer = new StreamWriter(savedlg.FileName))
                        {
                            writer.WriteLine(ReportBanner);
                            writer.WriteLine(ReportHeader);
                            writer.WriteLine("Program export ref des check is clean." + Environment.NewLine + "There are no Auto SMT differences." + Environment.NewLine);
                        }

                        //Open report in maximized notepad window
                        ProcessStartInfo procinfo = new ProcessStartInfo("notepad.exe", savedlg.FileName);
                        procinfo.WindowStyle = ProcessWindowStyle.Maximized;
                        Process.Start(procinfo);
                    }
                }
            }
            else
            {
                Microsoft.Win32.SaveFileDialog savedlg = new Microsoft.Win32.SaveFileDialog();
                savedlg.Filter = "Text Files (*.txt)|*.txt";
                savedlg.DefaultExt = "txt";
                savedlg.FileName = "Setup Sheet Compare";
                savedlg.AddExtension = true;
                savedlg.InitialDirectory = Path.GetDirectoryName(NewParser.FullFilePath);
                savedlg.Title = "Save Report";
                bool? saveResult = savedlg.ShowDialog();
                if (saveResult == true)
                {
                    using (StreamWriter writer = new StreamWriter(savedlg.FileName))
                    {
                        writer.WriteLine(ReportBanner);
                        writer.WriteLine(ReportHeader);
                        writer.WriteLine("See Auto SMT differences below..." + Environment.NewLine);
                        if (iExportErrorCount > 0)
                            WriteExportErrors(writer);
                        else
                        {
                            writer.WriteLine(Environment.NewLine + Environment.NewLine + "########################################");
                            writer.WriteLine("#     Program Exports Error Report     #");
                            writer.WriteLine("########################################" + Environment.NewLine);
                            writer.WriteLine("Program export ref des check is clean.");
                        }
                        if (lstSetupPassChanges.Count > 0)
                            WriteSetupPassChanges(writer);
                        if (NewParser.DuplicateRefs.Count > 0)
                            WriteSetupDups(writer);
                        if (lstFeederChangesFirstPass.Count > 0 || lstFeederChangesSecondPass.Count > 0)
                            WriteFeederChanges(writer);
                        if (lstSetupMachineChanges.Count > 0 || lstSetupSlotChanges.Count > 0)
                            WriteLocationChanges(writer);
                        WriteHandPlaceChanges(writer);
                        if(lstSetupRdAdds.Count > 0 || lstSetupRdDeletes.Count > 0 || lstSetupRdChangedPartNum.Count > 0)
                        {
                            WriteEcoHeader(writer);
                            if (lstSetupRdAdds.Count > 0)
                                WriteSetupAdds(writer);
                            if (lstSetupRdDeletes.Count > 0)
                                WriteSetupDeletes(writer);
                            if (lstSetupRdChangedPartNum.Count > 0)
                                WriteSetupPartChanges(writer);
                        }
                    }
                    MessageBox.Show("SMT differences found.  See results file.", "Report Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                    //Open report in maximized notepad window
                    ProcessStartInfo procinfo = new ProcessStartInfo("notepad.exe", savedlg.FileName);
                    procinfo.WindowStyle = ProcessWindowStyle.Maximized;
                    Process.Start(procinfo);
                }
            }
           
        }
      
        public void CompareSetupSheets()
        {
            Valid = CheckFileTimes();
            if (!Valid)
            {
                MessageBox.Show("All of your new files must be newer than all of your original files."
                                + "\nYours aren't.\nThis is an error. Correct this and restart this tool.",
                                "Something is not right here...", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            ///Start with export placement data compare
            ///
            List<string> oldCi2PlacementBot = CombineCi2Placements(OldCi2List, "BOT");
            List<string> newCi2PlacementBot = CombineCi2Placements(NewCi2List, "BOT");
            List<string> oldCi2PlacementTop = CombineCi2Placements(OldCi2List, "TOP");
            List<string> newCi2PlacementTop = CombineCi2Placements(NewCi2List, "TOP");

            ///After combining exports into single sides, find any duplicate ref des (one might be bypassed, but it's still dirty)
            ///
            botCi2Dups = newCi2PlacementBot
                .GroupBy(i => i)
                .Where(g => g.Count() > 1)
                .Select(g => g.Key).ToList();
            topCi2Dups = newCi2PlacementTop
                .GroupBy(i => i)
                .Where(g => g.Count() > 1)
                .Select(g => g.Key).ToList();

            ///After getting dups, clean the dups from the placement lists.
            ///
            newCi2PlacementBot = newCi2PlacementBot.Distinct().ToList();
            newCi2PlacementTop = newCi2PlacementTop.Distinct().ToList();

            ///Build error lists by comparing export placement lists old vs. new
            ///
            lstBotCi2Deleted = oldCi2PlacementBot.Except(newCi2PlacementBot).ToList();
            lstBotCi2Added = newCi2PlacementBot.Except(oldCi2PlacementBot).ToList();
            lstTopCi2Deleted = oldCi2PlacementTop.Except(newCi2PlacementTop).ToList();
            lstTopCi2Added = newCi2PlacementTop.Except(oldCi2PlacementTop).ToList();
            lstCi2JumpedBotToTop = oldCi2PlacementBot.Intersect(newCi2PlacementTop).ToList();
            lstCi2JumpedTopToBot = oldCi2PlacementTop.Intersect(newCi2PlacementBot).ToList();

            ///Clean the add + delete lists...if they jumped sides, keep them in the jumped sides list, remove from del/add lists
            lstBotCi2Deleted = lstBotCi2Deleted.Except(lstCi2JumpedBotToTop).ToList();
            lstBotCi2Added = lstBotCi2Added.Except(lstCi2JumpedTopToBot).ToList();
            lstTopCi2Deleted = lstTopCi2Deleted.Except(lstCi2JumpedTopToBot).ToList();
            lstTopCi2Added = lstTopCi2Added.Except(lstCi2JumpedBotToTop).ToList();

            ///Tally 'em up and get some bools for simple tests
            iExportErrorCount = botCi2Dups.Count + topCi2Dups.Count + lstBotCi2Added.Count
                            + lstBotCi2Deleted.Count + lstTopCi2Added.Count + lstTopCi2Deleted.Count
                            + lstCi2JumpedBotToTop.Count + lstCi2JumpedTopToBot.Count;
            bBotCi2sMatch = oldCi2PlacementBot.SequenceEqual(newCi2PlacementBot);
            bTopCi2sMatch = oldCi2PlacementTop.SequenceEqual(newCi2PlacementTop);
            ///Done with Export Errors
            ///

            ///Start Setup Sheet Compare
            ///
            lstSetupRdAdds = GetSetupRdAdds();
            lstSetupRdDeletes = GetSetupRdDeletes();
            lstSetupRdChangedPartNum = GetPartNumberChanges();
            lstSetupPassChanges = GetSetupPassChanges();
            //clean up adds and deletes
            lstSetupRdAdds = lstSetupRdAdds.Except(lstSetupPassChanges).ToList();
            lstSetupRdDeletes = lstSetupRdDeletes.Except(lstSetupPassChanges).ToList();

            //machine changes
            lstSetupMachineChanges = GetSetupMachineChanges();

            //slot changes
            lstSetupSlotChanges = GetSetupSlotChanges();

            //first pass feeder changes --use PnDicts
            lstFeederChangesFirstPass = GetFeederChanges(1);

            //second pass feeder changes --use PnDicts
            lstFeederChangesSecondPass = GetFeederChanges(2);

            //total up error count 
            iSetupErrorCount = lstSetupRdAdds.Count + lstSetupRdDeletes.Count + lstSetupRdChangedPartNum.Count
                + lstSetupPassChanges.Count + lstSetupMachineChanges.Count + lstSetupSlotChanges.Count
                + lstFeederChangesFirstPass.Count + lstFeederChangesSecondPass.Count + NewParser.DuplicateRefs.Count;
            ErrorCount = iExportErrorCount + iSetupErrorCount;

        }

        private bool CheckFileTimes()
        {
            bool bValid = true;
            List<DateTime> lstOldFileTimes = new List<DateTime>();
            List<DateTime> lstNewFileTimes = new List<DateTime>();
            foreach (var ci2 in OldCi2List)
            {
                if ((DateTime.Now - File.GetLastWriteTime(ci2.FullFilePath)).Days > 1)
                {
                    MessageBox.Show("Your original exports are more than 1 day old."
                        + "\nIt's recommended to output original ci2's just prior to doing the new setup sheet."
                        + "\nConsider using more up-to-date ci2's. (But it's up to you.)",
                        "Aged CI2 Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    break;
                }
            }
            foreach (var ci2 in OldCi2List)
                lstOldFileTimes.Add(File.GetLastWriteTime(ci2.FullFilePath));
            foreach (var ci2 in NewCi2List)
                lstNewFileTimes.Add(File.GetLastWriteTime(ci2.FullFilePath));
            lstOldFileTimes.Sort((a, b) => b.CompareTo(a));
            lstNewFileTimes.Sort((a, b) => a.CompareTo(b));
            if (lstOldFileTimes.First() >= lstNewFileTimes.First())
                bValid = false;

            return bValid;
        }

        private List<string> CombineCi2Placements(List<Ci2Parser> ci2List, string side)
        {
            string regstr = null;
            if (side.Equals("TOP"))
                regstr = @"TOP";
            else
                regstr = @"BOT";
            Regex reSide = new Regex(regstr, RegexOptions.IgnoreCase);
            List<string> lstPlacements = new List<string>();

            foreach (var ci2 in ci2List)
            {
                if (reSide.Match(ci2.FileName).Success)
                {
                    foreach (var placement in ci2.PlacementMap)
                    {
                        lstPlacements.Add(placement.Key);
                    }
                }
            }
            lstPlacements.Sort();
            return lstPlacements;
        }

        private List<string> GetSetupRdAdds()
        {
            List<string> lstAdds = new List<string>();
            foreach (var newRd in NewParser.FirstPassRefDesDict)
            {
                if (!OldParser.FirstPassRefDesDict.ContainsKey(newRd.Key))
                    lstAdds.Add(newRd.Key);
            }
            foreach (var newRd in NewParser.SecondPassRefDesDict)
            {
                if (!OldParser.SecondPassRefDesDict.ContainsKey(newRd.Key))
                    lstAdds.Add(newRd.Key);
            }
            return lstAdds;
        }

        private List<string> GetSetupRdDeletes()
        {
            List<string> lstDeletes = new List<string>();
            foreach (var oldRd in OldParser.FirstPassRefDesDict)
            {
                if (!NewParser.FirstPassRefDesDict.ContainsKey(oldRd.Key))
                    lstDeletes.Add(oldRd.Key);
            }
            foreach (var oldRd in OldParser.SecondPassRefDesDict)
            {
                if (!NewParser.SecondPassRefDesDict.ContainsKey(oldRd.Key))
                    lstDeletes.Add(oldRd.Key);
            }
            return lstDeletes;
        }
        private List<string> GetPartNumberChanges()
        {
            List<string> lstChangedPns = new List<string>();
            foreach (var pair in OldParser.FirstPassRefDesDict)
            {
                if (NewParser.FirstPassRefDesDict.ContainsKey(pair.Key) && !NewParser.FirstPassRefDesDict[pair.Key][0].Equals(pair.Value[0]))
                    lstChangedPns.Add(pair.Key);
            }
            foreach (var pair in OldParser.SecondPassRefDesDict)
            {
                if (NewParser.SecondPassRefDesDict.ContainsKey(pair.Key) && !NewParser.SecondPassRefDesDict[pair.Key][0].Equals(pair.Value[0]))
                    lstChangedPns.Add(pair.Key);
            }
            return lstChangedPns;
        }

        private List<string> GetSetupPassChanges()
        {
            List<string> lstPassChanged = new List<string>();
            foreach (var oldrd in OldParser.FirstPassRefDesDict)
            {
                if (NewParser.SecondPassRefDesDict.ContainsKey(oldrd.Key))
                    lstPassChanged.Add(oldrd.Key);
            }
            foreach (var oldrd in OldParser.SecondPassRefDesDict)
            {
                if (NewParser.FirstPassRefDesDict.ContainsKey(oldrd.Key))
                    lstPassChanged.Add(oldrd.Key);
            }
            return lstPassChanged;
        }

        private List<string> GetSetupMachineChanges()
        {
            List<string> lstMachineChanges = new List<string>();
            foreach (var pair in OldParser.FirstPassRefDesDict)
            {
                if (NewParser.FirstPassRefDesDict.ContainsKey(pair.Key) && !NewParser.FirstPassRefDesDict[pair.Key][2].Equals(pair.Value[2]))
                    lstMachineChanges.Add(pair.Key);
            }
            foreach (var pair in OldParser.SecondPassRefDesDict)
            {
                if (NewParser.SecondPassRefDesDict.ContainsKey(pair.Key) && !NewParser.SecondPassRefDesDict[pair.Key][2].Equals(pair.Value[2]))
                    lstMachineChanges.Add(pair.Key);
            }

            return lstMachineChanges;
        }
        private List<string> GetSetupSlotChanges()
        {
            List<string> lstSlotChanges = new List<string>();
            foreach (var pair in OldParser.FirstPassRefDesDict)
            {
                if (NewParser.FirstPassRefDesDict.ContainsKey(pair.Key) && !NewParser.FirstPassRefDesDict[pair.Key][3].Equals(pair.Value[3]))
                    lstSlotChanges.Add(pair.Key);
            }
            foreach (var pair in OldParser.SecondPassRefDesDict)
            {
                if (NewParser.SecondPassRefDesDict.ContainsKey(pair.Key) && !NewParser.SecondPassRefDesDict[pair.Key][3].Equals(pair.Value[3]))
                    lstSlotChanges.Add(pair.Key);
            }
            return lstSlotChanges;
        }

        private List<string> GetFeederChanges(int pass)
        {
            List<string> lstFeederChanges = new List<string>();
            if (pass == 1)
            {
                foreach (var oldPn in OldParser.FirstPassPnDict)
                {
                    if (NewParser.FirstPassPnDict.ContainsKey(oldPn.Key) && !NewParser.FirstPassPnDict[oldPn.Key][0].Equals(oldPn.Value[0]))
                        lstFeederChanges.Add(oldPn.Key);
                }
            }
            else if (pass == 2)
            {
                foreach (var oldPn in OldParser.SecondPassPnDict)
                {
                    if (NewParser.SecondPassPnDict.ContainsKey(oldPn.Key) && !NewParser.SecondPassPnDict[oldPn.Key][0].Equals(oldPn.Value[0]))
                        lstFeederChanges.Add(oldPn.Key);
                }
            }
            return lstFeederChanges;
        }

        private string GetReportHeader()
        {
            string header = "Old Setup Sheet: " + OldParser.FullFilePath + Environment.NewLine
                + "New Setup Sheet: " + NewParser.FullFilePath + Environment.NewLine + Environment.NewLine + Environment.NewLine
                + "*****Note: This tool only concerns itself with Auto SMT differences." + Environment.NewLine
                + "Example: A connector that got moved from Auto to Post, will be flagged as a \"Delete\" or \"DNI\" here." + Environment.NewLine
                + "Verify your post-Auto sections with the VAP." + Environment.NewLine + Environment.NewLine;
            return header;
        }
        private string GetReportBanner()
        {
            string banner = "                 __    __                  __              ____                                    ___              " + Environment.NewLine
                + " /'\\_/`\\        /\\ \\__/\\ \\                /\\ \\            /\\  _`\\                                 /\\_ \\             " + Environment.NewLine
                + "/\\      \\     __\\ \\ ,_\\ \\ \\___     ___    \\_\\ \\    ____   \\ \\ \\/\\_\\    ___     ___     ____    ___\\//\\ \\      __    " + Environment.NewLine
                + "\\ \\ \\__\\ \\  /'__`\\ \\ \\/\\ \\  _ `\\  / __`\\  /'_` \\  /',__\\   \\ \\ \\/_/_  / __`\\ /' _ `\\  /',__\\  / __`\\\\ \\ \\   /'__`\\  " + Environment.NewLine
                + " \\ \\ \\_/\\ \\/\\  __/\\ \\ \\_\\ \\ \\ \\ \\/\\ \\L\\ \\/\\ \\L\\ \\/\\__, `\\   \\ \\ \\L\\ \\/\\ \\L\\ \\/\\ \\/\\ \\/\\__, `\\/\\ \\L\\ \\\\_\\ \\_/\\  __/  " + Environment.NewLine
                + "  \\ \\_\\\\ \\_\\ \\____\\\\ \\__\\\\ \\_\\ \\_\\ \\____/\\ \\___,_\\/\\____/    \\ \\____/\\ \\____/\\ \\_\\ \\_\\/\\____/\\ \\____//\\____\\ \\____\\ " + Environment.NewLine
                + "   \\/_/ \\/_/\\/____/ \\/__/ \\/_/\\/_/\\/___/  \\/__,_ /\\/___/      \\/___/  \\/___/  \\/_/\\/_/\\/___/  \\/___/ \\/____/\\/____/ " + Environment.NewLine
                + "                                                                                                                    " + Environment.NewLine
                + " " + Environment.NewLine
                + "                                             __  __                        ___              " + Environment.NewLine
                + "                                            /\\ \\/\\ \\                 __  /'___\\             " + Environment.NewLine
                + "                                            \\ \\ \\ \\ \\     __   _ __ /\\_\\/\\ \\__/  __  __     " + Environment.NewLine
                + "                                             \\ \\ \\ \\ \\  /'__`\\/\\`'__\\/\\ \\ \\ ,__\\/\\ \\/\\ \\    " + Environment.NewLine
                + "                                              \\ \\ \\_/ \\/\\  __/\\ \\ \\/ \\ \\ \\ \\ \\_/\\ \\ \\_\\ \\   " + Environment.NewLine
                + "                                               \\ `\\___/\\ \\____\\\\ \\_\\  \\ \\_\\ \\_\\  \\/`____ \\  " + Environment.NewLine
                + "                                                `\\/__/  \\/____/ \\/_/   \\/_/\\/_/   `/___/> \\ " + Environment.NewLine
                + "                                                                                     /\\___/ " + Environment.NewLine
                + "                                                        Craig Thomson 2018           \\/__/  " + Environment.NewLine
                + "																												    " + Environment.NewLine;

            return banner;
        }

        public void WriteExportErrors(StreamWriter writer)
        {
            writer.WriteLine(Environment.NewLine + Environment.NewLine + "########################################");
            writer.WriteLine("#     Program Exports Error Report     #");
            writer.WriteLine("########################################" + Environment.NewLine);
            writer.WriteLine("This section describes errors between original exports and final exports.");
            writer.WriteLine("It looks at ref des's only. It is not concerned with part numbers.");
            writer.WriteLine("All ref des's should match before and after. Bypassed and NOBOM ref des's are included.");
            writer.WriteLine("Ref des's changing pass are the most serious errors that could happen here,");
            writer.WriteLine("but technically all categories of errors in this section should be fixed so programs are clean." + Environment.NewLine);

            if (lstCi2JumpedTopToBot.Count > 0)
            {
                writer.WriteLine("The following ref des moved passes from TOP to BOT." + Environment.NewLine + "SERIOUS ERROR!!!!:");
                writer.WriteLine(string.Join(",", lstCi2JumpedTopToBot) + Environment.NewLine);
            }
            if (lstCi2JumpedBotToTop.Count > 0)
            {
                writer.WriteLine("The following ref des moved passes from BOT to TOP." + Environment.NewLine + "SERIOUS ERROR!!!!:");
                writer.WriteLine(string.Join(", ", lstCi2JumpedBotToTop) + Environment.NewLine);
            }
            if (botCi2Dups.Count > 0)
            {
                writer.WriteLine("The following ref des are on BOT more than once. (Did you copy a NOBOM and forget to go back and cut??)");
                writer.WriteLine(string.Join(", ", botCi2Dups) + Environment.NewLine);
            }
            if (topCi2Dups.Count > 0)
            {
                writer.WriteLine("The following ref des are on TOP more than once. (Did you copy a NOBOM and forget to go back and cut??)");
                writer.WriteLine(string.Join(", ", topCi2Dups) + Environment.NewLine);
            }
            if (lstBotCi2Deleted.Count > 0)
            {
                writer.WriteLine("The following ref des got deleted from BOT in the new programs. What happened??");
                writer.WriteLine(string.Join(", ", lstBotCi2Deleted) + Environment.NewLine);
            }
            if (lstTopCi2Deleted.Count > 0)
            {
                writer.WriteLine("The following ref des got deleted from TOP in the new programs. What happened??");
                writer.WriteLine(string.Join(", ", lstTopCi2Deleted) + Environment.NewLine);
            }
            if (lstBotCi2Added.Count > 0)
            {
                writer.WriteLine("The following ref des got added to BOT in the new programs. Where did they come from??");
                writer.WriteLine(string.Join(", ", lstBotCi2Added) + Environment.NewLine);
            }
            if (lstTopCi2Added.Count > 0)
            {
                writer.WriteLine("The following ref des got added to TOP in the new programs. Where did they come from??");
                writer.WriteLine(string.Join(", ", lstTopCi2Added) + Environment.NewLine);
            }
        }
        private void WriteSetupPassChanges(StreamWriter writer)
        {
            writer.WriteLine(Environment.NewLine + Environment.NewLine);
            writer.WriteLine("****************************************");
            writer.WriteLine("* REF DES CHANGED PASS -- FIX THESE!!! *");
            writer.WriteLine("****************************************" + Environment.NewLine);

            foreach(string refdes in lstSetupPassChanges)
            {
                List<string> olddata = new List<string>();
                List<string> newdata = new List<string>();
                writer.WriteLine(refdes);
                if (OldParser.FirstPassRefDesDict.ContainsKey(refdes))
                    olddata = OldParser.FirstPassRefDesDict[refdes];
                else if (OldParser.SecondPassRefDesDict.ContainsKey(refdes))
                    olddata = OldParser.SecondPassRefDesDict[refdes];
                if (NewParser.FirstPassRefDesDict.ContainsKey(refdes))
                    newdata = NewParser.FirstPassRefDesDict[refdes];
                else if (NewParser.SecondPassRefDesDict.ContainsKey(refdes))
                    newdata = NewParser.SecondPassRefDesDict[refdes];
                if(olddata != null & newdata != null)
                {
                    writer.WriteLine("Old : " + string.Join(" ", olddata));
                    writer.WriteLine("New : " + string.Join(" ", newdata) + Environment.NewLine);
                }
            }
        }
        private void WriteSetupDups(StreamWriter writer)
        {
            writer.WriteLine(Environment.NewLine + Environment.NewLine);
            writer.WriteLine("****************************************");
            writer.WriteLine("* REF DES DUPLICATES  --  FIX THESE!!! *");
            writer.WriteLine("****************************************" + Environment.NewLine);
            writer.WriteLine("The following ref des's are found more than once in " + NewParser.FileName + ":");
            writer.WriteLine(Environment.NewLine + string.Format("{0,-22}{1,-22}", "Ref Des", "Part Numbers"));
            writer.WriteLine(string.Format("{0,-22}{1,-22}", "=======", "============"));
            foreach(var entry in NewParser.DuplicateRefs)
            {
                writer.WriteLine(string.Format("{0,-22}{1,-22}", entry.Key, string.Join(", ", entry.Value)));
            }
        }

        private void WriteFeederChanges(StreamWriter writer)
        {
            writer.WriteLine(Environment.NewLine + Environment.NewLine);
            writer.WriteLine("**************************************");
            writer.WriteLine("* PART NUMBER(S) CHANGED FEEDER TYPE *");
            writer.WriteLine("**************************************" + Environment.NewLine);
            if(lstFeederChangesFirstPass.Count > 0)
                writer.WriteLine("1st Pass" + Environment.NewLine + "========");
            foreach(string pn in lstFeederChangesFirstPass)
            {
                writer.WriteLine(pn + " :");
                writer.WriteLine("Old : " + OldParser.FirstPassPnDict[pn].FirstOrDefault());
                writer.WriteLine("New : " + NewParser.FirstPassPnDict[pn].FirstOrDefault() + Environment.NewLine);
            }
            writer.WriteLine();
            if (lstFeederChangesSecondPass.Count > 0)
                writer.WriteLine("2nd Pass" + Environment.NewLine + "========");
            foreach (string pn in lstFeederChangesSecondPass)
            {
                writer.WriteLine(pn + " :");
                writer.WriteLine("Old : " + OldParser.SecondPassPnDict[pn].FirstOrDefault());
                writer.WriteLine("New : " + NewParser.SecondPassPnDict[pn].FirstOrDefault() + Environment.NewLine);
            }
        }

        private void WriteLocationChanges(StreamWriter writer)
        {
            writer.WriteLine(Environment.NewLine + Environment.NewLine);
            writer.WriteLine("***********************************");
            writer.WriteLine("* FEEDER(S) MOVED TO NEW LOCATION *");
            writer.WriteLine("***********************************" + Environment.NewLine);
            Dictionary<string, List<string>> dictFirstPassMoves = GetLocMovesDictionary(1);
            Dictionary<string, List<string>> dictSecondPassMoves = GetLocMovesDictionary(2);

            if (dictFirstPassMoves.Count > 0)
                writer.WriteLine("1st Pass" + Environment.NewLine + "========");
            foreach(var entry in dictFirstPassMoves)
            {
                writer.WriteLine(entry.Key + " : ");
                writer.WriteLine("Old Location : " + entry.Value[0] + " " + entry.Value[1]);
                writer.WriteLine("New Location : " + entry.Value[2] + " " + entry.Value[3] + Environment.NewLine);
            }
            writer.WriteLine();
            if (dictSecondPassMoves.Count > 0)
                writer.WriteLine("2nd Pass" + Environment.NewLine + "========");
            foreach (var entry in dictSecondPassMoves)
            {
                writer.WriteLine(entry.Key + " : ");
                writer.WriteLine("Old Location : " + entry.Value[0] + " " + entry.Value[1]);
                writer.WriteLine("New Location : " + entry.Value[2] + " " + entry.Value[3] + Environment.NewLine);
            }


        }

        private Dictionary<string, List<string>> GetLocMovesDictionary(int pass)
        {
            Dictionary<string, List<string>> dictLocationMoves = new Dictionary<string, List<string>>();
            List<string> lstUniqueRds = lstSetupSlotChanges.Union(lstSetupMachineChanges).ToList();
            if(pass == 1)
            {
                foreach(string rd in lstUniqueRds)
                {
                    if (OldParser.FirstPassRefDesDict.ContainsKey(rd))
                    {
                        dictLocationMoves[OldParser.FirstPassRefDesDict[rd][0]] = new List<string>(new string[]
                        {
                            OldParser.FirstPassRefDesDict[rd][2], OldParser.FirstPassRefDesDict[rd][3],
                            NewParser.FirstPassRefDesDict[rd][2], NewParser.FirstPassRefDesDict[rd][3]
                        });
                    }

                }
            }
            else if(pass == 2)
            {
                foreach (string rd in lstUniqueRds)
                {
                    if (OldParser.SecondPassRefDesDict.ContainsKey(rd))
                    {
                        dictLocationMoves[OldParser.SecondPassRefDesDict[rd][0]] = new List<string>(new string[]
                        {
                            OldParser.SecondPassRefDesDict[rd][2], OldParser.SecondPassRefDesDict[rd][3],
                            NewParser.SecondPassRefDesDict[rd][2], NewParser.SecondPassRefDesDict[rd][3]
                        });
                    }

                }
            }
            return dictLocationMoves;
        }  
        private void WriteHandPlaceChanges(StreamWriter writer)
        {
            var queryFirstPassHpAdds = from rd in lstSetupRdAdds
                                         where NewParser.FirstPassRefDesDict.ContainsKey(rd) && NewParser.FirstPassRefDesDict[rd][3].Equals("Hand Place")
                                         orderby rd ascending
                                         select rd;
            var querySecondPassHpAdds = from rd in lstSetupRdAdds
                                           where NewParser.SecondPassRefDesDict.ContainsKey(rd) && NewParser.SecondPassRefDesDict[rd][3].Equals("Hand Place")
                                        orderby rd ascending
                                           select rd;
            var queryFirstPassHpDeletes = from rd in lstSetupRdDeletes
                                          where OldParser.FirstPassRefDesDict.ContainsKey(rd) && OldParser.FirstPassRefDesDict[rd][3].Equals("Hand Place")
                                          orderby rd ascending
                                          select rd;
            var querySecondPassHpDeletes = from rd in lstSetupRdDeletes
                                           where OldParser.SecondPassRefDesDict.ContainsKey(rd) && OldParser.SecondPassRefDesDict[rd][3].Equals("Hand Place")
                                           orderby rd ascending
                                           select rd;
            //var queryHandPlaceChangedPass = from rd in lstSetupPassChanges
            //                                where ((OldParser.FirstPassRefDesDict.ContainsKey(rd) && OldParser.FirstPassRefDesDict[rd][3].Equals("Hand Place"))
            //                                || (OldParser.SecondPassRefDesDict.ContainsKey(rd) && OldParser.SecondPassRefDesDict[rd][3].Equals("Hand Place")))
            //                                && ((NewParser.FirstPassRefDesDict.ContainsKey(rd) && NewParser.FirstPassRefDesDict[rd][3].Equals("Hand Place"))
            //                                || (NewParser.SecondPassRefDesDict.ContainsKey(rd) && NewParser.SecondPassRefDesDict[rd][3].Equals("Hand Place")))
            //                                orderby rd ascending
            //                                select rd; 

            if(queryFirstPassHpAdds.ToList().Count > 0 || querySecondPassHpAdds.ToList().Count > 0)
            {
                writer.WriteLine(Environment.NewLine + Environment.NewLine);
                writer.WriteLine("************************************");
                writer.WriteLine("*    ADDS TO HANDPLACE SECTION     *");
                writer.WriteLine("************************************" + Environment.NewLine);
                if(queryFirstPassHpAdds.ToList().Count > 0)
                {
                    writer.WriteLine("1st Pass:" + Environment.NewLine);
                    writer.WriteLine(string.Format("{0,-22}{1,-22}", "Ref", @"P\N"));
                    writer.WriteLine(string.Format("{0,-22}{1,-22}", "===", "==="));
                    foreach (var rd in queryFirstPassHpAdds)
                    {
                        writer.WriteLine(string.Format("{0,-22}{1,-22}", rd, NewParser.FirstPassRefDesDict[rd][0]));
                    }
                }
                writer.WriteLine();
                if (querySecondPassHpAdds.ToList().Count > 0)
                {
                    writer.WriteLine("2nd Pass:" + Environment.NewLine);
                    writer.WriteLine(string.Format("{0,-22}{1,-22}", "Ref", @"P\N"));
                    writer.WriteLine(string.Format("{0,-22}{1,-22}", "===", "==="));
                    foreach (var rd in querySecondPassHpAdds)
                    {
                        writer.WriteLine(string.Format("{0,-22}{1,-22}", rd, NewParser.SecondPassRefDesDict[rd][0]));
                    }
                }
            }
            if (queryFirstPassHpDeletes.ToList().Count > 0 || querySecondPassHpDeletes.ToList().Count > 0)
            {
                writer.WriteLine(Environment.NewLine + Environment.NewLine);
                writer.WriteLine("*************************************");
                writer.WriteLine("*  DELETES FROM HANDPLACE SECTION   *");
                writer.WriteLine("*************************************" + Environment.NewLine);
                writer.WriteLine("NOTE: These may be legit DNI's now, or you might have accidentally cut it out. Verify with BOM." + Environment.NewLine);
                if (queryFirstPassHpDeletes.ToList().Count > 0)
                {
                    writer.WriteLine("1st Pass:" + Environment.NewLine);
                    writer.WriteLine(string.Format("{0,-22}{1,-22}", "Ref", @"P\N"));
                    writer.WriteLine(string.Format("{0,-22}{1,-22}", "===", "==="));
                    foreach (var rd in queryFirstPassHpDeletes)
                    {
                        writer.WriteLine(string.Format("{0,-22}{1,-22}", rd, OldParser.FirstPassRefDesDict[rd][0]));
                    }
                }
                writer.WriteLine();
                if (querySecondPassHpDeletes.ToList().Count > 0)
                {
                    writer.WriteLine("2nd Pass:" + Environment.NewLine);
                    writer.WriteLine(string.Format("{0,-22}{1,-22}", "Ref", @"P\N"));
                    writer.WriteLine(string.Format("{0,-22}{1,-22}", "===", "==="));
                    foreach (var rd in querySecondPassHpDeletes)
                    {
                        writer.WriteLine(string.Format("{0,-22}{1,-22}", rd, OldParser.SecondPassRefDesDict[rd][0]));
                    }
                }
            }
            //if (queryHandPlaceChangedPass.ToList().Count > 0)
            //{
            //    writer.WriteLine(Environment.NewLine + Environment.NewLine);
            //    writer.WriteLine("*************************************");
            //    writer.WriteLine("*    HANDPLACE PARTS MOVED PASS     *");
            //    writer.WriteLine("*************************************" + Environment.NewLine);
            //    writer.WriteLine(string.Format("{0,-22}{1,-22}{2,-22}{3,-22}", "Ref", @"P\N", "Old Pass", "New Pass"));
            //    writer.WriteLine(string.Format("{0,-22}{1,-22}{2,-22}{3,-22}", "===", "===", "========", "========"));
            //    string oldpass = null;
            //    string newpass = null;
            //    string pn = null;
            //    foreach(var rd in queryHandPlaceChangedPass)
            //    {
            //        if (OldParser.FirstPassRefDesDict.ContainsKey(rd))
            //        {
            //            oldpass = "1st Pass";
            //            newpass = "2nd Pass";
            //            pn = OldParser.FirstPassRefDesDict[rd][0];
            //        }
            //        else
            //        {
            //            oldpass = "2nd Pass";
            //            newpass = "1st Pass";
            //            pn = OldParser.SecondPassRefDesDict[rd][0];
            //        }
            //        writer.WriteLine(string.Format("{0,-22}{1,-22}{2,-22}{3,-22}", rd, pn, oldpass, newpass));
            //    }
            //}
            writer.WriteLine();
            writer.WriteLine();
        } 
        private void WriteEcoHeader(StreamWriter writer)
        {
            writer.WriteLine(Environment.NewLine + Environment.NewLine);
            writer.WriteLine("=======================================================================================");
            writer.WriteLine("# The remainder of this report contains information that is most likely ECO activity. #");
            writer.WriteLine("# None of these should be errors, as long as the setup sheet didn't call them out.    #");
            writer.WriteLine("# This section is provided for informational purposes only. Feel free to ignore it.   #");
            writer.WriteLine("=======================================================================================" + Environment.NewLine);
        }
        private void WriteSetupAdds(StreamWriter writer)
        {
            writer.WriteLine(Environment.NewLine + Environment.NewLine);
            writer.WriteLine("***************************************");
            writer.WriteLine("* REF DES'S ADDED IN NEW SETUP SHEET. *");
            writer.WriteLine("***************************************" + Environment.NewLine);
            writer.WriteLine(string.Format("{0,-22}{1,-22}{2,-22}", "Ref", "Old", "New"));
            writer.WriteLine(string.Format("{0,-22}{1,-22}{2,-22}", "===", "===", "==="));
            foreach(string rd in lstSetupRdAdds)
            {
                string pn = null;
                if (NewParser.FirstPassRefDesDict.ContainsKey(rd))
                    pn = NewParser.FirstPassRefDesDict[rd][0];
                else
                    pn = NewParser.SecondPassRefDesDict[rd][0];
                writer.WriteLine(string.Format("{0,-22}{1,-22}{2,-22}", rd, "NOBOM", pn));
            }
        }
        private void WriteSetupDeletes(StreamWriter writer)
        {
            writer.WriteLine(Environment.NewLine + Environment.NewLine);
            writer.WriteLine("*****************************************");
            writer.WriteLine("* REF DES'S NOW DNI IN NEW SETUP SHEET. *");
            writer.WriteLine("*****************************************" + Environment.NewLine);
            writer.WriteLine(string.Format("{0,-22}{1,-22}{2,-22}", "Ref", "Old", "New"));
            writer.WriteLine(string.Format("{0,-22}{1,-22}{2,-22}", "===", "===", "==="));

            foreach(string rd in lstSetupRdDeletes)
            {
                string pn = null;
                if (OldParser.FirstPassRefDesDict.ContainsKey(rd))
                    pn = OldParser.FirstPassRefDesDict[rd][0];
                else
                    pn = OldParser.SecondPassRefDesDict[rd][0];
                writer.WriteLine(string.Format("{0,-22}{1,-22}{2,-22}", rd, pn, "NOBOM"));
            }
        }
        private void WriteSetupPartChanges(StreamWriter writer)
        {
            writer.WriteLine(Environment.NewLine + Environment.NewLine);
            writer.WriteLine("*******************************************");
            writer.WriteLine("* PART NUMBER CHANGES IN NEW SETUP SHEET. *");
            writer.WriteLine("*******************************************" + Environment.NewLine);
            writer.WriteLine(string.Format("{0,-22}{1,-22}{2,-22}", "Ref", "Old", "New"));
            writer.WriteLine(string.Format("{0,-22}{1,-22}{2,-22}", "===", "===", "==="));

            foreach(string rd in lstSetupRdChangedPartNum)
            {
                string oldpn = null;
                string newpn = null;
                if (OldParser.FirstPassRefDesDict.ContainsKey(rd))
                {
                    oldpn = OldParser.FirstPassRefDesDict[rd][0];
                    newpn = NewParser.FirstPassRefDesDict[rd][0];
                }
                else
                {
                    oldpn = OldParser.SecondPassRefDesDict[rd][0];
                    newpn = NewParser.SecondPassRefDesDict[rd][0];
                }
                writer.WriteLine(string.Format("{0,-22}{1,-22}{2,-22}", rd, oldpn, newpn));
            }
        }
    }
}
