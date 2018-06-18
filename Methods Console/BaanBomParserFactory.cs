using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.IO;

namespace Methods_Console
{
    class BaanBomParserFactory : FileParserFactory
    {
        private string _filename;

        public BaanBomParserFactory(string file)
        {
            _filename = file;
            if (!System.IO.File.Exists(_filename) || !Path.GetExtension(file).ToLower().Equals(".txt"))
            {
                _filename = null;
            }
        }


        public override FileParser GetFileParser()
        {
            if (_filename == null)
            {
                MessageBox.Show("Error opening BAAN BOM file.\nCheck file format (does it have Operations assigned?).", "BaanBomParserFactory.GetFileParser()", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
            else
            {
                return new BaanBOMParser(_filename);
            }
        }
    }
}
