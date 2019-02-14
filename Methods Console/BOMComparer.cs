using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;


namespace Methods_Console
{
    class BOMComparer
    {
        Assembly ThisProgram;
        AssemblyName ThisProgramName;
        Version ThisProgramVersion;
        public BeiBOM BomOne { get; private set; }
        public BeiBOM BomTwo { get; private set; }
        public SortedDictionary<string, List<string>> BomOnePartsAndRefs { get; private set; }
        public SortedDictionary<string, List<string>> BomTwoPartsAndRefs { get; private set; }
        public SortedDictionary<string, List<string>> InBomOneButNotInBomTwo { get; private set; }
        public SortedDictionary<string, List<string>> InBomTwoButNotInBomOne { get; private set; }

        public BOMComparer(BeiBOM bomOne, BeiBOM bomTwo)
        {
            ThisProgram = Assembly.GetEntryAssembly();
            ThisProgramName = ThisProgram.GetName();
            ThisProgramVersion = ThisProgramName.Version;
            BomOne = bomOne;
            BomTwo = bomTwo;
            BomOnePartsAndRefs = PopulatePartsAndRefs(BomOne);
            BomTwoPartsAndRefs = PopulatePartsAndRefs(BomTwo);
            InBomOneButNotInBomTwo = new SortedDictionary<string, List<string>>();
            InBomTwoButNotInBomOne = new SortedDictionary<string, List<string>>();
        }

        public bool CompareBomData()
        {
            ///return true if part numbers and ref des's match between BomOne and BomTwo, otherwise return false
            ///
            bool BomsMatch = false;
            InBomOneButNotInBomTwo.Clear();
            InBomTwoButNotInBomOne.Clear();

            foreach(var bomoneentry in BomOnePartsAndRefs)
            {
                if (BomTwoPartsAndRefs.ContainsKey(bomoneentry.Key))
                {
                    List<string> RefsNotInBomTwo = bomoneentry.Value.Except(BomTwoPartsAndRefs[bomoneentry.Key]).ToList();
                    if (RefsNotInBomTwo.Count > 0)
                        InBomOneButNotInBomTwo.Add(bomoneentry.Key, new List<string>(RefsNotInBomTwo));
                }
                else
                {
                    InBomOneButNotInBomTwo.Add(bomoneentry.Key, new List<string>(bomoneentry.Value));
                }
            }

            foreach (var bomtwoentry in BomTwoPartsAndRefs)
            {
                if (BomOnePartsAndRefs.ContainsKey(bomtwoentry.Key))
                {
                    List<string> RefsNotInBomOne = bomtwoentry.Value.Except(BomOnePartsAndRefs[bomtwoentry.Key]).ToList();
                    if (RefsNotInBomOne.Count > 0)
                        InBomTwoButNotInBomOne.Add(bomtwoentry.Key, new List<string>(RefsNotInBomOne));
                }
                else
                {
                    InBomTwoButNotInBomOne.Add(bomtwoentry.Key, new List<string>(bomtwoentry.Value));
                }
            }
            if (InBomOneButNotInBomTwo.Count == 0 && InBomTwoButNotInBomOne.Count == 0)
                BomsMatch = true;
            Microsoft.Win32.SaveFileDialog savedlg = new Microsoft.Win32.SaveFileDialog();
            savedlg.Filter = "Text Files (*.txt)|*.txt";
            savedlg.DefaultExt = "txt";
            savedlg.FileName = "BOM Compare";
            savedlg.AddExtension = true;
            savedlg.InitialDirectory = Path.GetDirectoryName(BomTwo.FullFilePath);
            savedlg.Title = "Save BOM Compare Report";
            bool? saveResult = savedlg.ShowDialog();
            using (StreamWriter writer = new StreamWriter(savedlg.FileName))
            {
                writer.WriteLine("Methods Console - BOM Compare Module: software version " + ThisProgramVersion);
            }
                return BomsMatch;
        }
        private SortedDictionary<string, List<string>> PopulatePartsAndRefs(BeiBOM bom)
        {
            SortedDictionary<string, List<string>> PartsAndRefs = new SortedDictionary<string, List<string>>();
            AlphanumComparator AlphaNumCompare = new AlphanumComparator();
            foreach(var entry in bom.Bom)
            {
                if (PartsAndRefs.ContainsKey(entry.Value.Item1))
                {
                    PartsAndRefs[entry.Value.Item1].AddRange(entry.Value.Item4.Split(',').ToList());
                    PartsAndRefs[entry.Value.Item1].Sort(AlphaNumCompare);
                    
                }
                else
                {
                    PartsAndRefs.Add(entry.Value.Item1, entry.Value.Item4.Split(',').ToList());
                    PartsAndRefs[entry.Value.Item1].Sort(AlphaNumCompare);
                }
            }
            return PartsAndRefs;
        }

    }
}
