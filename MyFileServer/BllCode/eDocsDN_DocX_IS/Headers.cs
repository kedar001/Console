using System.Collections.Generic;
namespace eDocsDN_DocX_IS
{
    public class Headers
    {
        public string Id
        {
            get;
            set;
        }

        internal Headers()
        {
        }

        public Header odd;
        public Header even;
        public Header first;

        public List<Header> lstodd;
        public List<Header> lsteven;
        public List<Header> lstfirst;

    }




}
