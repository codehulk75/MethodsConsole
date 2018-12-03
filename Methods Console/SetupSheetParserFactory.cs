using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Methods_Console
{
    class SetupSheetParserFactory : FileParserFactory
    {
        private string _filename;

        public SetupSheetParserFactory(string filepath)
        {
            _filename = filepath;
            if (!System.IO.File.Exists(_filename))
            {
                _filename = null;
            }
            else if (!Path.GetExtension(_filename).ToLower().Equals(".rtf"))
            {
                _filename = null;
            }
        }
        public override FileParser GetFileParser()
        {
            if (_filename == null)
            {
                MessageBox.Show("Error opening setup sheet file.\nCheck to make sure file is a valid setup sheet and you have permission to open it.", "SetupSheetParserFactory.GetFileParser()", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
            else
            {
                return new SetupSheetParser(_filename);
            }
        }
    }
}
