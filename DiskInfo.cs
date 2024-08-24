using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CheckDiskSpace
{
    public class DiskInfo
    {
        public DiskInfo() { }
        public string DiskName { get; set; }
        public string FreeSpace { get; set; }
    }
}
