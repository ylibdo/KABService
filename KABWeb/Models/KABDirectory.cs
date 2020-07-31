using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace KABWeb.Models
{
    public class KABDirectory
    {
        public int SearchDepth { get; set; }
        public string Name { get; set; }
        public IEnumerable<KABDirectory> SubDirectoryList { get; set; }
        public IEnumerable<FileInfo> FileList { get; set; }
    }
}
