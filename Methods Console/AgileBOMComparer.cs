using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Methods_Console
{
    class AgileBOMComparer
    {
        public BeiBOM BomOne { get; private set; }
        public BeiBOM BomTwo { get; private set; }
        public SortedDictionary<string, List<string>> BomOnePartsAndRefs { get; private set; }
        public SortedDictionary<string, List<string>> BomTwoPartsAndRefs { get; private set; }

        public AgileBOMComparer(BeiBOM bomOne, BeiBOM bomTwo)
        {
            BomOne = bomOne;
            BomTwo = bomTwo;
            BomOnePartsAndRefs = PopulatePartsAndRefs(BomOne);
            BomTwoPartsAndRefs = PopulatePartsAndRefs(BomTwo);
        }

        public bool CompareBomData()
        {
            ///return true if part numbers and ref des's match between BomOne and BomTwo, otherwise return false
            ///
            bool BomsMatch = false;



            return BomsMatch;
        }
        private SortedDictionary<string, List<string>> PopulatePartsAndRefs(BeiBOM bom)
        {
            SortedDictionary<string, List<string>> PartsAndRefs = new SortedDictionary<string, List<string>>();
            foreach(var entry in bom.Bom)
            {
                if (PartsAndRefs.ContainsKey(entry.Value.Item1))
                {
                    PartsAndRefs[entry.Value.Item1].AddRange(entry.Value.Item4.Split(',').ToList());
                    PartsAndRefs[entry.Value.Item1].Sort();
                }
                else
                {
                    PartsAndRefs.Add(entry.Value.Item1, entry.Value.Item4.Split(',').ToList());
                }
            }
            return PartsAndRefs;
        }

    }
}
