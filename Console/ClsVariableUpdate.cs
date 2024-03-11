using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestConsole
{
    public class ClsVariable_Update
    {
        IUpdate_Variable _IUpdateVariables;
        public ClsVariable_Update(IUpdate_Variable iUpdateVariables)
        {

            _IUpdateVariables = iUpdateVariables;
        }
        public void Update_Variable(List<string> lst)
        {
            _IUpdateVariables.Update_Variables(lst);
        }

    }

    public class Update_Using_OpenXml : IUpdate_Variable
    {
        public void Update_Variables(List<string> lst)
        {
            foreach (var item in lst)
            {
                Console.WriteLine("OPenXML : " + item);
            }

        }
    }
    public class Update_Using_SuncFusion : IUpdate_Variable
    {
        public void Update_Variables(List<string> lst)
        {
            foreach (var item in lst)
            {
                Console.WriteLine("SuncFusion  : " + item);
            }
        }
    }
    public class Update_Using_Word : IUpdate_Variable
    {
        public void Update_Variables(List<string> lst)
        {
            foreach (var item in lst)
            {
                Console.WriteLine("Word  : " + item);
            }
        }
    }

    public interface IUpdate_Variable
    {
        void Update_Variables(List<string> lst);
    }



}
