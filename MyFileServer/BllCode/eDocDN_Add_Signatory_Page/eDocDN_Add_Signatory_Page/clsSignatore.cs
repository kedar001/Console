using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace eDocDN_Add_Signatory_Page
{
    public class clsPossitionOfSignature
    {
        public bool AutoGenerateSignature { get; set; }
        public bool isFirstPageSignature { get; set; }
    }
    public class clsSignatore
    {
        public string UserID { get; set; }
        public string UserDepartment { get; set; }
        public string UserDesignation { get; set; }
        public string UserType { get; set; }
        public string Sequence { get; set; }
        public string UserFullName { get; set; }
        public DateTime UserDate { get; set; }
        public DateTime UserTime { get; set; }
    }
}
