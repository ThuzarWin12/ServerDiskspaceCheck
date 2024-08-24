using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CheckDiskSpace
{
    public class ServerInfo
    {
        public ServerInfo() { }
        public List<DiskInfo> DiskInfo  = new List<DiskInfo>();
        public string ServerName { get; set; }
        public DateTime RunTime { get; set; }

    }
}
