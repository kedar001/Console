using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace eDocDN_Error_Log
{
    public enum LogType
    {
        E = 0,
        D
    }

    public class LogRecord
    {
        public string LogText { get; set; }
        public LogType eLogType { get; set; }
    }
}
