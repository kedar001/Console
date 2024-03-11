using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;
using eDocsDN_ReadAppXml;
using eDocsDN_File_Operations;
using eDocsDN_Get_Directory_Info;
using System.IO;
using System.Data;
using System.ServiceModel.Activation;


namespace MyFileServer
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "Service1" in code, svc and config file together.
    // NOTE: In order to launch WCF Test Client for testing this service, please select Service1.svc or Service1.svc.cs at the Solution Explorer and start debugging.
    [ServiceBehavior(InstanceContextMode = InstanceContextMode.PerCall,
        ConcurrencyMode = ConcurrencyMode.Multiple,
        AddressFilterMode = AddressFilterMode.Any,
        ReleaseServiceInstanceOnTransactionComplete = false,
         TransactionTimeout = "00:30:00")]
    public class Service1 : IService1, IDisposable
    {
        #region ..... Variable Declaration ....
        clsReadAppXml _objINI;
        ClsCopy_Files _objFile_Operations = null;
        string _szAppXmlpath = string.Empty;
        string _szAppLocation = string.Empty;
        string _szLocation = string.Empty;
        string _szDBNAme = string.Empty;
        string _szLogFileName = "";
        bool _bisDebug = false;

        #endregion

        #region .... Constructor ....
        public Service1()
        {
            try
            {
                _szAppXmlpath = AppDomain.CurrentDomain.BaseDirectory + "DllApps\\Application.xml";
                if (!File.Exists(_szAppXmlpath))
                    throw new Exception("Application.xml File not found. Path :" + _szAppXmlpath);

                _objINI = new clsReadAppXml(_szAppXmlpath);
                _szAppLocation = _objINI.GetAppLocation().Trim();

                _szLocation = _objINI.GetCurrentLocation().Trim();
                _szDBNAme = _objINI.GetLocationVariable(_szLocation, "", "Database");
                _bisDebug = _objINI.GetLocationVariable(_szLocation, "", "DebugStatus").ToUpper() == "TRUE" ? true : false;
                _objINI.Dispose();
                _objINI = null;
                _szLogFileName = AppDomain.CurrentDomain.BaseDirectory + @"\Log\DetailLog.txt";
                if (!Directory.Exists(AppDomain.CurrentDomain.BaseDirectory + @"\Log\"))
                    Directory.CreateDirectory(AppDomain.CurrentDomain.BaseDirectory + @"\Log\");

            }
            catch (Exception ex)
            {
                System.Diagnostics.EventLog.WriteEntry("FileServer", ex.StackTrace, System.Diagnostics.EventLogEntryType.Error);
                throw new FaultException(ex.Message);
            }
        }

        public Service1(string szAppXmlPath)
        {
            try
            {
                _szAppXmlpath = szAppXmlPath + "DllApps\\Application.xml";
                if (!File.Exists(_szAppXmlpath))
                    throw new Exception("Application.xml File not found. Path :" + _szAppXmlpath);

                _objINI = new clsReadAppXml(_szAppXmlpath);
                _szAppLocation = _objINI.GetAppLocation().Trim();

                _szLocation = _objINI.GetCurrentLocation().Trim();
                _szDBNAme = _objINI.GetLocationVariable(_szLocation, "", "Database");
                _bisDebug = _objINI.GetLocationVariable(_szLocation, "", "DebugStatus").ToUpper() == "TRUE" ? true : false;
                _objINI.Dispose();
                _objINI = null;
                _szLogFileName = AppDomain.CurrentDomain.BaseDirectory + @"\Log\DetailLog.txt";
                if (!Directory.Exists(AppDomain.CurrentDomain.BaseDirectory + @"\Log\"))
                {

                }
                Directory.CreateDirectory(AppDomain.CurrentDomain.BaseDirectory + @"\Log\");
            }
            catch (Exception ex)
            {
                throw new FaultException(ex.Message);
            }
        }
        #endregion

        #region ..... Public Methods ...

        [OperationBehavior(AutoDisposeParameters = false, ReleaseInstanceMode = ReleaseInstanceMode.BeforeAndAfterCall)]
        public File_Data CopyFile(File_Data oFileData)
        {
            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : " + _szLogFileName + Environment.NewLine);
            try
            {
                _objFile_Operations = new ClsCopy_Files(_szAppXmlpath, _szDBNAme, _szLocation);
                if (!string.IsNullOrEmpty(oFileData.Source_Directory))
                {
                    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : string.IsNullOrEmpty(oFileData.Source_Directory) : " + Environment.NewLine);
                    oFileData = _objFile_Operations.Copy_File(oFileData.Source_Directory, oFileData.Destination_Directory, oFileData);
                    if (_objFile_Operations.msgError != "")
                        throw new Exception(_objFile_Operations.msgError);
                }
                else
                {
                    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : oFileData.Source_Directory : " + Environment.NewLine);
                    oFileData = _objFile_Operations.Copy_File(oFileData);
                    if (_objFile_Operations.msgError != "")
                        throw new Exception(_objFile_Operations.msgError);
                }
                if (!oFileData.Need_File_Blob)
                    oFileData.Data = null;

                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : END " + Environment.NewLine);

            }
            catch (Exception ex)
            {
                string message =
       "Exception type " + ex.GetType() + Environment.NewLine +
       "Exception message: " + ex.Message + Environment.NewLine +
       "Stack trace: " + ex.StackTrace + Environment.NewLine;
                if (ex.InnerException != null)
                {
                    message += "---BEGIN InnerException--- " + Environment.NewLine +
                               "Exception type " + ex.InnerException.GetType() + Environment.NewLine +
                               "Exception message: " + ex.InnerException.Message + Environment.NewLine +
                               "Stack trace: " + ex.InnerException.StackTrace + Environment.NewLine +
                               "---END Inner Exception";
                }
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " CopyFile(File_Data oFileData): " + message + Environment.NewLine);
                throw new FaultException(message);
            }
            finally
            {
                //GC.Collect();
            }
            return oFileData;
        }

        [OperationBehavior(AutoDisposeParameters = false, ReleaseInstanceMode = ReleaseInstanceMode.BeforeAndAfterCall)]
        public List<File_Data> CopyFile(List<File_Data> lstFileData)
        {
            try
            {
                _objFile_Operations = new ClsCopy_Files(_szAppXmlpath, _szDBNAme, _szLocation);
                for (int iIndex = 0; iIndex < lstFileData.Count; iIndex++)
                {
                    if (!string.IsNullOrEmpty(lstFileData[iIndex].Source_Directory))
                    {
                        lstFileData[iIndex] = _objFile_Operations.Copy_File(lstFileData[iIndex].Source_Directory, lstFileData[iIndex].Destination_Directory, lstFileData[iIndex]);
                        if (_objFile_Operations.msgError != "")
                            throw new Exception(_objFile_Operations.msgError);
                    }
                    else
                    {
                        lstFileData[iIndex] = _objFile_Operations.Copy_File(lstFileData[iIndex]);
                        if (_objFile_Operations.msgError != "")
                            throw new Exception(_objFile_Operations.msgError);
                    }
                    if (!lstFileData[iIndex].Need_File_Blob)
                        lstFileData[iIndex].Data = null;
                }
            }
            catch (Exception ex)
            {
                string message =
       "Exception type " + ex.GetType() + Environment.NewLine +
       "Exception message: " + ex.Message + Environment.NewLine +
       "Stack trace: " + ex.StackTrace + Environment.NewLine;
                if (ex.InnerException != null)
                {
                    message += "---BEGIN InnerException--- " + Environment.NewLine +
                               "Exception type " + ex.InnerException.GetType() + Environment.NewLine +
                               "Exception message: " + ex.InnerException.Message + Environment.NewLine +
                               "Stack trace: " + ex.InnerException.StackTrace + Environment.NewLine +
                               "---END Inner Exception";
                }
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : " + message + Environment.NewLine);
                throw new FaultException(message);
            }
            finally
            {
                //..GC.Collect();
            }
            return lstFileData;
        }


        [OperationBehavior(AutoDisposeParameters = false, ReleaseInstanceMode = ReleaseInstanceMode.BeforeAndAfterCall)]
        public bool Delete_File(File_Data oFileData)
        {
            bool bResult = true;
            try
            {
                _objFile_Operations = new ClsCopy_Files(_szAppXmlpath, _szDBNAme, _szLocation);
                bResult = _objFile_Operations.Delete_File(oFileData);
                if (_objFile_Operations.msgError != "")
                    throw new Exception(_objFile_Operations.msgError);
            }
            catch (Exception ex)
            {
                bResult = false;
                string message =
       "Exception type " + ex.GetType() + Environment.NewLine +
       "Exception message: " + ex.Message + Environment.NewLine +
       "Stack trace: " + ex.StackTrace + Environment.NewLine;
                if (ex.InnerException != null)
                {
                    message += "---BEGIN InnerException--- " + Environment.NewLine +
                               "Exception type " + ex.InnerException.GetType() + Environment.NewLine +
                               "Exception message: " + ex.InnerException.Message + Environment.NewLine +
                               "Stack trace: " + ex.InnerException.StackTrace + Environment.NewLine +
                               "---END Inner Exception";
                }
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : " + message + Environment.NewLine);
                throw new FaultException(message);
            }
            finally
            {
                if (_objFile_Operations != null)
                    _objFile_Operations.Dispose();
                _objFile_Operations = null;

                if (oFileData != null)
                    oFileData.Dispose();
                oFileData = null;

            }
            return bResult;
        }

        [OperationBehavior(AutoDisposeParameters = false, ReleaseInstanceMode = ReleaseInstanceMode.BeforeAndAfterCall)]
        public bool Check_File_Exist(File_Data oFileData)
        {
            bool bResult = true;
            try
            {
                _objFile_Operations = new ClsCopy_Files(_szAppXmlpath, _szDBNAme, _szLocation);
                bResult = _objFile_Operations.Check_File_Exist_In_Source(oFileData);
                if (_objFile_Operations.msgError != "")
                    throw new Exception(_objFile_Operations.msgError);
            }
            catch (Exception ex)
            {
                bResult = false;
                string message =
       "Exception type " + ex.GetType() + Environment.NewLine +
       "Exception message: " + ex.Message + Environment.NewLine +
       "Stack trace: " + ex.StackTrace + Environment.NewLine;
                if (ex.InnerException != null)
                {
                    message += "---BEGIN InnerException--- " + Environment.NewLine +
                               "Exception type " + ex.InnerException.GetType() + Environment.NewLine +
                               "Exception message: " + ex.InnerException.Message + Environment.NewLine +
                               "Stack trace: " + ex.InnerException.StackTrace + Environment.NewLine +
                               "---END Inner Exception";
                }
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Check_File_Exist: " + message + Environment.NewLine);
                throw new FaultException(message);
            }
            finally
            {
                if (_objFile_Operations != null)
                    _objFile_Operations.Dispose();
                _objFile_Operations = null;

                if (oFileData != null)
                    oFileData.Dispose();
                oFileData = null;
            }
            return bResult;
        }

        [OperationBehavior(AutoDisposeParameters = false, ReleaseInstanceMode = ReleaseInstanceMode.BeforeAndAfterCall)]
        public File_Data Get_File_Information(File_Data oFileData)
        {
            try
            {
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Get_File_Information : " + Environment.NewLine);
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " _szAppXmlpath : " + _szAppXmlpath + Environment.NewLine);
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " _szDBNAme : " + _szDBNAme + Environment.NewLine);
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " _szLocation : " + _szLocation + Environment.NewLine);
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " _objFile_Operations : " + Environment.NewLine);
                _objFile_Operations = new ClsCopy_Files(_szAppXmlpath, _szDBNAme, _szLocation);
                if (_objFile_Operations.msgError != "")
                    throw new Exception("Error Message from Constructor :" + _objFile_Operations.msgError);

                oFileData = _objFile_Operations.Get_File_Information(oFileData);
                if (_objFile_Operations.msgError != "")
                    throw new Exception(_objFile_Operations.msgError);
            }
            catch (Exception ex)
            {
                string message =
       "Exception type " + ex.GetType() + Environment.NewLine +
       "Exception message: " + ex.Message + Environment.NewLine +
       "Stack trace: " + ex.StackTrace + Environment.NewLine;
                if (ex.InnerException != null)
                {
                    message += "---BEGIN InnerException--- " + Environment.NewLine +
                               "Exception type " + ex.InnerException.GetType() + Environment.NewLine +
                               "Exception message: " + ex.InnerException.Message + Environment.NewLine +
                               "Stack trace: " + ex.InnerException.StackTrace + Environment.NewLine +
                               "---END Inner Exception";
                }
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Get_File_Information : " + message + Environment.NewLine);
                throw new FaultException(message);
            }
            finally
            {
                if (_objFile_Operations != null)
                    _objFile_Operations.Dispose();
                _objFile_Operations = null;
            }
            return oFileData;
        }

        [OperationBehavior(AutoDisposeParameters = false, ReleaseInstanceMode = ReleaseInstanceMode.BeforeAndAfterCall)]
        public List<File_Data> Get_Documents(File_Data oFileData)
        {
            List<File_Data> lstFile_Data = new List<File_Data>();
            try
            {
                _objFile_Operations = new ClsCopy_Files(_szAppXmlpath, _szDBNAme, _szLocation);
                lstFile_Data = _objFile_Operations.Get_Documents(oFileData);
                if (_objFile_Operations.msgError != "")
                    throw new Exception(_objFile_Operations.msgError);
            }
            catch (Exception ex)
            {
                string message =
       "Exception type " + ex.GetType() + Environment.NewLine +
       "Exception message: " + ex.Message + Environment.NewLine +
       "Stack trace: " + ex.StackTrace + Environment.NewLine;
                if (ex.InnerException != null)
                {
                    message += "---BEGIN InnerException--- " + Environment.NewLine +
                               "Exception type " + ex.InnerException.GetType() + Environment.NewLine +
                               "Exception message: " + ex.InnerException.Message + Environment.NewLine +
                               "Stack trace: " + ex.InnerException.StackTrace + Environment.NewLine +
                               "---END Inner Exception";
                }
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Get_Documents: " + message + Environment.NewLine);
                throw new FaultException(message);
            }
            finally
            {
                if (_objFile_Operations != null)
                    _objFile_Operations.Dispose();
                _objFile_Operations = null;
            }
            return lstFile_Data;
        }

        public bool Check_File_Is_Locked(string szFileName, string szOfficeVersion)
        {
            bool bResult = false;
            try
            {
                _objFile_Operations = new ClsCopy_Files(_szAppXmlpath, _szDBNAme, _szLocation);
                bResult = _objFile_Operations.Check_File_is_Locked(szFileName, szOfficeVersion);
                if (_objFile_Operations.msgError != "")
                    throw new Exception(_objFile_Operations.msgError);
            }
            catch (Exception ex)
            {
                bResult = true;
                string message =
       "Exception type " + ex.GetType() + Environment.NewLine +
       "Exception message: " + ex.Message + Environment.NewLine +
       "Stack trace: " + ex.StackTrace + Environment.NewLine;
                if (ex.InnerException != null)
                {
                    message += "---BEGIN InnerException--- " + Environment.NewLine +
                               "Exception type " + ex.InnerException.GetType() + Environment.NewLine +
                               "Exception message: " + ex.InnerException.Message + Environment.NewLine +
                               "Stack trace: " + ex.InnerException.StackTrace + Environment.NewLine +
                               "---END Inner Exception";
                }
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Check_File_Is_Locked: " + message + Environment.NewLine);
                throw new FaultException(ex.Message);
            }
            finally
            {
                if (_objFile_Operations != null)
                    _objFile_Operations.Dispose();
                _objFile_Operations = null;
            }
            return bResult;
        }
        public bool Check_File_Is_Locked(string szFileName)
        {
            bool bResult = false;
            try
            {
                _objFile_Operations = new ClsCopy_Files(_szAppXmlpath, _szDBNAme, _szLocation);
                bResult = _objFile_Operations.Check_File_is_Locked(szFileName);
                if (_objFile_Operations.msgError != "")
                    throw new Exception(_objFile_Operations.msgError);
            }
            catch (Exception ex)
            {
                bResult = true;
                string message =
       "Exception type " + ex.GetType() + Environment.NewLine +
       "Exception message: " + ex.Message + Environment.NewLine +
       "Stack trace: " + ex.StackTrace + Environment.NewLine;
                if (ex.InnerException != null)
                {
                    message += "---BEGIN InnerException--- " + Environment.NewLine +
                               "Exception type " + ex.InnerException.GetType() + Environment.NewLine +
                               "Exception message: " + ex.InnerException.Message + Environment.NewLine +
                               "Stack trace: " + ex.InnerException.StackTrace + Environment.NewLine +
                               "---END Inner Exception";
                }
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Check_File_Is_Locked: " + message + Environment.NewLine);
                throw new FaultException(ex.Message);
            }
            finally
            {
                if (_objFile_Operations != null)
                    _objFile_Operations.Dispose();
                _objFile_Operations = null;
            }
            return bResult;
        }


        public bool Pre_Check_File(string szFileName)
        {
            bool bResult = true;
            try
            {
                _objFile_Operations = new ClsCopy_Files(_szAppXmlpath, _szDBNAme, _szLocation);
                bResult = _objFile_Operations.Pre_Check_File(szFileName);
                if (_objFile_Operations.msgError != "")
                    throw new Exception(_objFile_Operations.msgError);
            }
            catch (Exception ex)
            {
                bResult = false;
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Pre_Check_File: " + ex.Message + Environment.NewLine);
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Pre_Check_File: " + ex.StackTrace + Environment.NewLine);
                throw new FaultException(ex.Message);
            }
            finally
            {
                if (_objFile_Operations != null)
                    _objFile_Operations.Dispose();
                _objFile_Operations = null;
            }
            return bResult;
        }
        public bool Pre_Check_File(byte[] arrFileData)
        {
            bool bResult = true;
            try
            {
                _objFile_Operations = new ClsCopy_Files(_szAppXmlpath, _szDBNAme, _szLocation);
                bResult = _objFile_Operations.Pre_Check_File(arrFileData);
                if (_objFile_Operations.msgError != "")
                    throw new Exception(_objFile_Operations.msgError);
            }
            catch (Exception ex)
            {
                bResult = false;
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Pre_Check_File: " + ex.Message + Environment.NewLine);
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Pre_Check_File: " + ex.StackTrace + Environment.NewLine);
                throw new FaultException(ex.Message);
            }
            finally
            {
                if (_objFile_Operations != null)
                    _objFile_Operations.Dispose();
                _objFile_Operations = null;
            }
            return bResult;
        }
        public bool Pre_Check_File(File_Data oFileData)
        {
            bool bResult = true;
            try
            {
                _objFile_Operations = new ClsCopy_Files(_szAppXmlpath, _szDBNAme, _szLocation);
                bResult = _objFile_Operations.Pre_Check_File(oFileData);
                if (_objFile_Operations.msgError != "")
                    throw new Exception(_objFile_Operations.msgError);
            }
            catch (Exception ex)
            {
                bResult = false;
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Pre_Check_File: " + ex.Message + Environment.NewLine);
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Pre_Check_File: " + ex.StackTrace + Environment.NewLine);
                throw new FaultException(ex.Message);
            }
            finally
            {
                if (_objFile_Operations != null)
                    _objFile_Operations.Dispose();
                _objFile_Operations = null;
            }
            return bResult;
        }



        public string Get_Server_Date_Time()
        {
            return System.DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss");
        }

        [OperationBehavior(TransactionScopeRequired = true)]
        public string Get_Server_Time()
        {
            return System.DateTime.Now.ToString("T");
        }

        [OperationBehavior(AutoDisposeParameters = false, ReleaseInstanceMode = ReleaseInstanceMode.BeforeAndAfterCall)]
        public string Get_Document_CheckSum(File_Data oFileData)
        {
            string szDocumentCheckSum = string.Empty;
            try
            {
                _objFile_Operations = new ClsCopy_Files(_szAppXmlpath, _szDBNAme, _szLocation);
                szDocumentCheckSum = _objFile_Operations.Get_File_Checksum(oFileData);
                if (_objFile_Operations.msgError != "")
                    throw new Exception(_objFile_Operations.msgError);
            }
            catch (Exception ex)
            {
                string message =
       "Exception type " + ex.GetType() + Environment.NewLine +
       "Exception message: " + ex.Message + Environment.NewLine +
       "Stack trace: " + ex.StackTrace + Environment.NewLine;
                if (ex.InnerException != null)
                {
                    message += "---BEGIN InnerException--- " + Environment.NewLine +
                               "Exception type " + ex.InnerException.GetType() + Environment.NewLine +
                               "Exception message: " + ex.InnerException.Message + Environment.NewLine +
                               "Stack trace: " + ex.InnerException.StackTrace + Environment.NewLine +
                               "---END Inner Exception";
                }
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Get_Document_CheckSum: " + message + Environment.NewLine);
                throw new FaultException(message);
            }
            finally
            {
                if (_objFile_Operations != null)
                    _objFile_Operations.Dispose();
                _objFile_Operations = null;
            }
            return szDocumentCheckSum;
        }

        [OperationBehavior(AutoDisposeParameters = false, ReleaseInstanceMode = ReleaseInstanceMode.BeforeAndAfterCall)]
        public bool Prepare_Document_For_Print(File_Data oFileData)
        {
            bool bResult = true;
            try
            {
                _objFile_Operations = new ClsCopy_Files(_szAppXmlpath, _szDBNAme, _szLocation);
                _objFile_Operations.Get_File_Checksum(oFileData);
                if (_objFile_Operations.msgError != "")
                    throw new Exception(_objFile_Operations.msgError);
            }
            catch (Exception ex)
            {
                bResult = false;
                string message =
       "Exception type " + ex.GetType() + Environment.NewLine +
       "Exception message: " + ex.Message + Environment.NewLine +
       "Stack trace: " + ex.StackTrace + Environment.NewLine;
                if (ex.InnerException != null)
                {
                    message += "---BEGIN InnerException--- " + Environment.NewLine +
                               "Exception type " + ex.InnerException.GetType() + Environment.NewLine +
                               "Exception message: " + ex.InnerException.Message + Environment.NewLine +
                               "Stack trace: " + ex.InnerException.StackTrace + Environment.NewLine +
                               "---END Inner Exception";
                }
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : " + message + Environment.NewLine);
                throw new FaultException(message);
            }
            finally
            {
                if (_objFile_Operations != null)
                    _objFile_Operations.Dispose();
                _objFile_Operations = null;
            }
            return bResult;
        }

        [OperationBehavior(AutoDisposeParameters = false, ReleaseInstanceMode = ReleaseInstanceMode.BeforeAndAfterCall)]
        public bool Convert_Document_To_PDF(File_Data oFileData)
        {
            bool bResult = true;
            try
            {
                _objFile_Operations = new ClsCopy_Files(_szAppXmlpath, _szDBNAme, _szLocation);
                _objFile_Operations.Convert_Document_to_Pdf(oFileData);
                if (_objFile_Operations.msgError != "")
                    throw new Exception(_objFile_Operations.msgError);
            }
            catch (Exception ex)
            {
                bResult = false;
                string message =
       "Exception type " + ex.GetType() + Environment.NewLine +
       "Exception message: " + ex.Message + Environment.NewLine +
       "Stack trace: " + ex.StackTrace + Environment.NewLine;
                if (ex.InnerException != null)
                {
                    message += "---BEGIN InnerException--- " + Environment.NewLine +
                               "Exception type " + ex.InnerException.GetType() + Environment.NewLine +
                               "Exception message: " + ex.InnerException.Message + Environment.NewLine +
                               "Stack trace: " + ex.InnerException.StackTrace + Environment.NewLine +
                               "---END Inner Exception";
                }
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : " + message + Environment.NewLine);
                throw new FaultException(message);
            }
            finally
            {
                if (_objFile_Operations != null)
                    _objFile_Operations.Dispose();
                _objFile_Operations = null;
            }
            return bResult;
        }



        //private bool isDebugEnable(string szAppXmlPath)
        //{
        //    bool bIsDebugEnable = false;
        //    string szDebugStatus = string.Empty;
        //    try
        //    {
        //        _objINI = new clsReadAppXml(szAppXmlPath);
        //        szDebugStatus = _objINI.GetLocationVariable(_szLocation, "", "DebugStatus");
        //        if (_objINI.ErrorMsg != "")
        //            throw new Exception(_objINI.ErrorMsg);
        //        if (szDebugStatus.Equals("TRUE", StringComparison.InvariantCultureIgnoreCase))
        //            bIsDebugEnable = true;

        //    }
        //    finally { }
        //    return bIsDebugEnable;
        //}

        #endregion

        #region .... IDISPOSABLE ....

        public void Dispose()
        {
            Dispose(true);
           //.. GC.Collect();
           //.. GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_objINI != null)
                    _objINI.Dispose();
                _objINI = null;
                if (_objFile_Operations != null)
                    _objFile_Operations.Dispose();
                _objFile_Operations = null;
                _szAppXmlpath = string.Empty;
                _szAppLocation = string.Empty;
                _szLocation = string.Empty;
                _szDBNAme = string.Empty;
            }
            else
            {

            }
        }

        ~Service1()
        {
            Dispose(false);
        }


        #endregion

    }
}
