using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestConsole
{
    public class clsGetVariables
    {
        IGet_Variables _IGetVariables;
        public clsGetVariables(IGet_Variables GetVariables)
        {
            _IGetVariables = GetVariables;
        }
        public List<string> Get_Variable_For_Updation()
        {
            return _IGetVariables.Get_Variables();
        }


    }

    public class BackEnd_CL : IGet_Variables
    {
        public List<string> Get_Variables()
        {
            List<string> obj = new List<string>();
            obj.Add("Company");
            obj.Add("Location");
            return obj;
        }
    }
    public class BackEnd_TR : IGet_Variables
    {
        public List<string> Get_Variables()
        {
            List<string> obj = new List<string>();
            obj.Add("Author");
            obj.Add("Author Date");
            return obj;
        }
    }


    public interface IGet_Variables
    {
        List<string> Get_Variables();
    }
}
