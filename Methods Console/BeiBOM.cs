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
    class BeiBOM
    {
        public string FileType { get; private set; }
        public string FileName { get; private set; }
        public string FullFilePath { get; private set; }
        public bool HasRouting { get; private set; }
        public bool IsValid { get; private set; }
        public string AssemblyName { get; private set; }
        public string DateOfListing { get; private set; }
        public string AssyDescription { get; private set; }
        public string Rev { get; private set; }
        public List<string> RouteList { get; private set; }
        public Dictionary<string, Tuple<string, string, string, string,string>> Bom { get; private set; } //Key = '<Findnum>:<Sequence>', Tuple.Item1=Part Number, Tuple.Item2=Operation, Tuple.Item3=Description, Tuple.Item4=comma-separated ref des's 


        public BeiBOM(string strBomFile)
        {
            IsValid = false;
            FullFilePath = strBomFile;
            FileName = Path.GetFileName(FullFilePath);
            FileType = Path.GetExtension(strBomFile).ToLower();
            Bom = new Dictionary<string, Tuple<string, string, string, string, string>>();
            HasRouting = false;
            if (FileType.Equals(".txt"))
            {
                if (IsValidBaanBOM())
                {
                    HasRouting = true;
                    BaanBomParserFactory factory = new BaanBomParserFactory(FullFilePath);
                    BaanBOMParser baanparser = null;
                    if (factory != null)
                        baanparser = (BaanBOMParser)factory.GetFileParser();
                    else
                        MessageBox.Show("Failed to create BAAN BOM Parser Object!", "BOMExplosionParser.ParseBAANBom()", MessageBoxButton.OK, MessageBoxImage.Error);

                    ///populate memebers with baanparser data
                    ///
                    AssemblyName = baanparser.AssemblyName;
                    AssyDescription = baanparser.AssyDescription;
                    DateOfListing = baanparser.DateOfListing;
                    Rev = baanparser.Rev;
                    RouteList = baanparser.RouteList;
                    foreach (var record in baanparser.BomMap)
                    {
                        string pn = baanparser.BomMap[record.Key][0];
                        string op = baanparser.BomMap[record.Key][1];
                        string desc = baanparser.BomMap[record.Key][2];
                        string refdes = baanparser.BomMap[record.Key][3];
                        string bomqty = baanparser.BomMap[record.Key][4];
                        Bom.Add(record.Key, Tuple.Create(pn, op, desc, refdes, bomqty));
                    }

                    IsValid = baanparser.IsValid;
                }

                else
                {
                    ClearData();
                    MessageBox.Show("Not a valid BAAN BOM!\nPlease check your bom file and try again.", "BAAN BOM Format Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }

            }
            else if (FileType.Equals(".csv") || FileType.Equals(".xls") || FileType.Equals(".xlsx"))
            {
                HasRouting = false;
                BOMExplosionParserFactory factory = new BOMExplosionParserFactory(FullFilePath);
                BOMExplosionParser agileparser = null;
                if (factory != null)
                    agileparser = (BOMExplosionParser)factory.GetFileParser();
                else
                    MessageBox.Show("BAAN BOM Parsing Failed", "BOMExplosionParser.ParseBAANBom()", MessageBoxButton.OK, MessageBoxImage.Error);

                AssemblyName = agileparser.AssemblyName;
                AssyDescription = agileparser.AssyDescription;
                DateOfListing = agileparser.DateOfListing;
                Rev = agileparser.Rev;
                foreach (var record in agileparser.BomMap)
                {
                    string pn = agileparser.BomMap[record.Key][0];
                    string op = null;
                    string desc = agileparser.BomMap[record.Key][1];
                    string refdes = agileparser.BomMap[record.Key][2];
                    string bomqty = agileparser.BomMap[record.Key][3];
                    Bom.Add(record.Key, Tuple.Create(pn, op, desc, refdes, bomqty));
                }

                IsValid = agileparser.IsValid;

            }
            else
            {
                MessageBox.Show("Not a valid BOM!\nPlease check your bom file and try again.", "BOM Format Error-BeiBOM.BeiBom()", MessageBoxButton.OK, MessageBoxImage.Error);
                ClearData();
                IsValid = false;
            }
        }

        private void ClearData()
        {
            FileType = null;
            FileName = null;
            FullFilePath = null;
            HasRouting = false;
            IsValid = false;
            AssemblyName = null;
            DateOfListing = null;
            AssyDescription = null;
            Rev = null;
            RouteList.Clear();
            Bom.Clear();
         }

        private bool IsValidBaanBOM()
        {
            bool isvalid = false;
            bool hasbomline = false;
            bool hasroutingline = false;
            bool hasroutingitem = false;

            string line;
            Regex reBomLine = new Regex(@"BILLS OF MATERIAL\s+\(MULTILEVEL\)\s+\(WITH BOM QUANTITIES\)");
            Regex reRoutingLine = new Regex(@"\(Including\sRouting\)");
            Regex reRoutingItem = new Regex(@"Routing Item");
            try
            {
                using (StreamReader sr = new StreamReader(FullFilePath))
                {
                    while ((line = sr.ReadLine()) != null)
                    {
                        if (reBomLine.Match(line).Success)
                            hasbomline = true;
                        if (reRoutingLine.Match(line).Success)
                            hasroutingline = true;
                        if (reRoutingItem.Match(line).Success)
                            hasroutingitem = true;
                    }
                }
                if (hasbomline && hasroutingline && hasroutingitem)
                    isvalid = true;
            }
            catch (Exception e)
            {
                MessageBox.Show("Error reading BAAN BOM. Could not load BOM.\n" + e.Message, "IsValidBaanBOM()", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            return isvalid;
        }
    }
}
