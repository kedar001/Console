using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace eDocsDN_File_Operations
{
    public interface ISave_File_as_Blob
    {
        bool Save_File_In_Database(string szTable_Name, List<File_Data> lstFile_Data);
    }
}
