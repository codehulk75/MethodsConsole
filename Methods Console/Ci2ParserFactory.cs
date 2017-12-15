using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.IO;

namespace Methods_Console
{
    class Ci2ParserFactory : FileParserFactory
    {
        private string _filename;

        public Ci2ParserFactory(string filepath)
        {
            _filename = filepath;   
            if (!System.IO.File.Exists(_filename))
            {
                _filename = null;
            }
            else if (!Path.GetExtension(_filename).ToLower().Equals(".ci2"))
            {
                _filename = null;
            }
        }

        public override FileParser GetFileParser()
        {
            if (_filename == null)
            {
                MessageBox.Show("Error opening ci2 export file.\nCheck to make sure file is a valid ci2 and you have permission to open it.", "Ci2FileParserFactory.GetFileParser()");
                return null;
            }
            else
            {
                return new Ci2Parser(_filename);
            }          
        }
    }
}
