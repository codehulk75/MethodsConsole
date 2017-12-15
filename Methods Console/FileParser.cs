using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Methods_Console
{
    abstract class FileParser
    {
        public abstract string FileType { get; set; }
        public abstract string FileName { get; set; }
        public abstract string FullFilePath { get; set; }

    }
}
