using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.IO;

namespace Methods_Console
{
    class BOMExplosionParserFactory : FileParserFactory
    {
        private string _filename;
        private string _ext;

        public BOMExplosionParserFactory(string filepath)
        {
            _filename = filepath;
            _ext = Path.GetExtension(filepath).ToLower();
            if (!System.IO.File.Exists(_filename))
            {
                _filename = null;
            }
            else if(!_ext.Equals(".xls") && !_ext.Equals(".txt") && !_ext.Equals(".csv"))
            {               
                _filename = null;
            }
        }

        public override FileParser GetFileParser()
        {
            if (_filename == null)
            {
                MessageBox.Show("Error opening BOM file.\nBOM should be an Agile BOM Explosion or BAAN BOM file. Try again.", "BOMExplosionParserFactory.GetFileParser()");
                return null;
            }
            else
            {
                return new BOMExplosionParser(_filename, _ext);
            }
        }
    }
}