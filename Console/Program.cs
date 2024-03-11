﻿using System;
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

namespace TestConsole
{
    class Program
    {
        static WSHttpBinding _binding = null;
        static EndpointAddress _endpoint = null;
        public static int iProcess = 1;

        public static void Main(string[] args)
        {
            ClsVariable_Update objUpdate;
            IUpdate_Variable IUp;
            IGet_Variables IGet;

            clsGetVariables objGetVariables = new clsGetVariables(new BackEnd_CL());
            List<string> lst = objGetVariables.Get_Variable_For_Updation();
            switch (iProcess)
            {
                case 0:
                    objUpdate = new ClsVariable_Update(new Update_Using_OpenXml());
                    objUpdate.Update_Variable(lst);
                    break;

                case 1:
                    objUpdate = new ClsVariable_Update(new Update_Using_SuncFusion());
                    objUpdate.Update_Variable(lst);
                    break;
                case 2:
                    objUpdate = new ClsVariable_Update(new Update_Using_Word());
                    objUpdate.Update_Variable(lst);
                    break;
            }
            Console.ReadLine();
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
