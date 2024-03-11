using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace eDocsDN_Get_Directory_Info
{
    public class Directory_Attributes : IDisposable
    {
        public string Directory_Path { get; set; }
        public bool Files_To_Be_Encrypted { get; set; }
        public bool Physical_Directory { get; set; }
        public bool Database_Storage { get; set; }
        public string Table_Name { get; set; }

        #region .... IDISPOSABLE ....
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                Directory_Path = string.Empty;
                Table_Name = string.Empty;
            }
            else
            {

            }
        }

        ~Directory_Attributes()
        {
            Dispose(false);
        }


        #endregion
    }
}
