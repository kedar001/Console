using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;


using System.ServiceModel.Channels;
using System.Security.Cryptography.X509Certificates;
using System.Net;
using System.Net.Security;
using System.ServiceModel;
using TestConsole.ServiceReference1;
using System.IO;
using System.Text.RegularExpressions;

namespace TestConsole
{
    class Program
    {
        static WSHttpBinding _binding = null;
        static EndpointAddress _endpoint = null;
        public static int iProcess = 1;

        public static void Main(string[] args)
        {




            ///...bug fixes
            ///...bug fixes1
            ///bug fix 3 
            Check_Docver("1.101", "NN.NN");







            //ClsVariable_Update objUpdate;
            //IUpdate_Variable IUp;
            //IGet_Variables IGet;

            //clsGetVariables objGetVariables = new clsGetVariables(new BackEnd_CL());
            //List<string> lst = objGetVariables.Get_Variable_For_Updation();
            //switch (iProcess)
            //{
            //    case 0:
            //        objUpdate = new ClsVariable_Update(new Update_Using_OpenXml());
            //        objUpdate.Update_Variable(lst);
            //        break;

            //    case 1:
            //        objUpdate = new ClsVariable_Update(new Update_Using_SuncFusion());
            //        objUpdate.Update_Variable(lst);
            //        break;
            //    case 2:
            //        objUpdate = new ClsVariable_Update(new Update_Using_Word());
            //        objUpdate.Update_Variable(lst);
            //        break;
            //}
            Console.ReadLine();
        }

        private static bool Check_Docver(string szDocver, string szDocVer_MaskType)
        {

            string szLNum = "", szRNum = "";
            string szMLNum = "", szMRNum = "";
            string[] arrVal;
            string[] arrMVal;
            char[] chSplit = { '.' };
            bool bReturn = true;

            //int iVer = Convert.ToInt32(szDocver);
            //if (iVer == 0)
            //{
            //    bReturn = false;
            //    MsgError = "Migration Failed because Version should not be 0";
            //}


            arrVal = szDocver.Split(chSplit);
            arrMVal = szDocVer_MaskType.Split(chSplit);

            if (arrVal.Length > 1)
            {
                szLNum = arrVal[0];
                szRNum = arrVal[1];
            }
            if (arrMVal.Length > 1)
            {
                szMLNum = arrMVal[0];
                szMRNum = arrMVal[1];
            }

            if (szLNum.Equals("0"))
            {
                bReturn = false;
                throw new Exception("Migration Failed because Numeric part in doc version is not an Valid value");
            }

            if (szDocVer_MaskType.ToUpper() == "NN.XX")
            {
                if (IsNumeric(szLNum) == false)
                {
                    bReturn = false;
                    throw new Exception("Migration Failed because Numeric part in doc version is not an numeric value");
                }
                if (szLNum.Equals("00.00") == false)
                {
                    bReturn = false;
                    throw new Exception("Migration Failed because Numeric part in doc version is not an numeric value");
                }
            }
            else
            {
                if (arrVal.Length == 1)
                {
                    bReturn = false;
                    throw new Exception("Migration Failed because Document version is not in correct format");
                }
                else if (IsNumeric(szLNum) == false || IsNumeric(szRNum) == false)
                {
                    bReturn = false;
                    throw new Exception("Migration Failed because Document version is not numeric");
                }

                if (szDocVer_MaskType.Equals("N.N"))
                {
                    if (szRNum.Length != 1)
                    {
                        bReturn = false;
                        throw new Exception("Migration Failed because Document version is not in correct format");
                    }
                }
                else
                {

                    var regExp = "^[0-9]{1," + szMLNum.Length.ToString() + "}\\.[0-9]{" + szMRNum.Length.ToString() + "}$";
                    Regex re = new Regex(regExp);
                    if (!re.IsMatch(szDocver))
                    {
                        bReturn = false;
                        throw new Exception("Migration Failed because Document version is not in correct format");
                    }
                }





                //if (szLNum.Length != szMLNum.Length)
                //{
                //    if (Convert.ToInt32(szLNum) < 10)
                //    { }
                //    else
                //    {
                //        bReturn = false;
                //        MsgError = "Migration Failed because Document version is not in correct format";
                //    }
                //}
                //if (szRNum.Length != szMRNum.Length)
                //{
                //    bReturn = false;
                //    MsgError = "Migration Failed because Document version is not in correct format";
                //}
            }

            return bReturn;
        }
        private static bool IsNumeric(string szValue)
        {
            bool bMatch;
            Match oMatch;
            Regex isNumeric = new Regex(@"^\d+$");
            oMatch = isNumeric.Match(szValue);
            bMatch = oMatch.Success;

            if (bMatch == true)
                return true;
            else
                return false;
        }

    }




    //public class Update_Variables
    //{
    //    IUpdateVariables _IUpdateVariables;
    //    IGetVariables _IGetVariables;
    //    public Update_Variables(IUpdateVariables UpdateVariables, IGetVariables GetVariables)
    //    {
    //        _IUpdateVariables = UpdateVariables;
    //        _IGetVariables = GetVariables;
    //    }
    //    public void Update_Variable()
    //    {
    //        _IUpdateVariables.Update_Variable(_IGetVariables.Get_Variable());
    //    }

    //}

    //public class CL : IGetVariables, IUpdateVariables
    //{
    //    List<string> obj = null;
    //    //IGetVariables _IGetVariables;
    //    public List<string> Get_Variable()
    //    {
    //        obj = new List<string>();
    //        obj.Add("Company");
    //        obj.Add("Location");
    //        return obj;
    //    }

    //    public void Update_Variable(List<string> lst)
    //    {
    //        foreach (var item in lst)
    //        {
    //            Console.WriteLine("CL:  " + item.ToString());
    //        }
    //    }

    //    //public void Update_Variable(IGetVariables _IGetVariables)
    //    //{
    //    //    foreach (var item in _IGetVariables.Get_Variable())
    //    //    {
    //    //        Console.WriteLine(item.ToString());
    //    //    }
    //    //}
    //}

    //public class TR : IGetVariables, IUpdateVariables, IUpdateVariablesWord
    //{

    //    int iUpdateMethod = 0;
    //    List<string> obj = null;
    //    public List<string> Get_Variable()
    //    {
    //        obj = new List<string>();
    //        obj.Add("author");
    //        obj.Add("author name");
    //        return obj;
    //    }
    //    public void Update_Variable(List<string> lst)
    //    {
    //        foreach (var item in lst)
    //        {
    //            Console.WriteLine("TR:  " + item.ToString());
    //        }
    //    }

    //    //public void Update_Variable(List<string> lst)
    //    //{
    //    //    foreach (var item in lst)
    //    //    {
    //    //        Console.WriteLine("TR  Word:  " + item.ToString());
    //    //    }
    //    //}



    //    //public void Update_Variable(IGetVariables _IGetVariables)
    //    //{
    //    //    foreach (var item in _IGetVariables.Get_Variable())
    //    //    {
    //    //        Console.WriteLine(item.ToString());
    //    //    }
    //    //}
    //}

    //public interface IUpdateVariables
    //{
    //    //IGetVariables _IGetVariables;
    //    //void Update_Variable(IGetVariables _IGetVariables);
    //    void Update_Variable(List<string> lst);

    //    //List<string> Get_Variable();
    //}

    //public interface IUpdateVariablesWord
    //{
    //    //IGetVariables _IGetVariables;
    //    //void Update_Variable(IGetVariables _IGetVariables);
    //    void Update_Variable(List<string> lst);

    //    //List<string> Get_Variable();
    //}




    //public interface IGetVariables
    //{
    //    List<string> Get_Variable();
    //}



    public interface ILogger
    {
        void LogMessage(string aString);
    }

    public class DbLogger : ILogger
    {
        public void LogMessage(string aMessage)
        {
            Console.WriteLine("db :" + aMessage);
        }
    }
    public class FileLogger : ILogger
    {
        public void LogMessage(string aStackTrace)
        {
            Console.WriteLine("file: " + aStackTrace);
        }
    }
    public class ExceptionLogger
    {
        private ILogger _logger;
        public ExceptionLogger(ILogger aLogger)
        {
            this._logger = aLogger;
        }
        public void LogException(string aException)
        {
            string strMessage = GetUserReadableMessage(aException);
            this._logger.LogMessage(strMessage);
        }
        private string GetUserReadableMessage(string aException)
        {
            //string strMessage = string.Empty;
            //code to convert Exception's stack trace and message to user readable format.

            return aException;
        }
    }
    public class DataExporter
    {
        public void ExportDataFromFile()
        {
            ExceptionLogger _exceptionLogger;
            try
            {
                _exceptionLogger = new ExceptionLogger(new DbLogger());
                _exceptionLogger.LogException(" for db");

                //_exceptionLogger = new ExceptionLogger(new FileLogger());
                //_exceptionLogger.LogException(" for file");
            }
            catch (IOException ex)
            {

            }
            catch (Exception ex)
            {

            }
        }
    }

}
