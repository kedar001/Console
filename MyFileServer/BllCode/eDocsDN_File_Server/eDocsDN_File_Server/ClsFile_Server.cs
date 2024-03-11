using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using eDocsDN_ReadAppXml;
using System.ServiceModel;
using eDocsDN_File_Server_Operations.File_Server;
using System.ServiceModel.Channels;

namespace eDocsDN_File_Server_Operations
{
    public class ClsFile_Server_Operations : IDisposable
    {
        #region .... Variable Declaration ...
        clsReadAppXml _objINI = null;

        WSHttpBinding _binding = null;
        string _szAppXmlPath = string.Empty;
        string _szDBName = string.Empty;
        string _szServerIP = string.Empty;
        string _szLocation = string.Empty;
        string _szTempSharedPath = string.Empty;
        bool _bIsBackEndProcess = false;
        EndpointAddress _endpoint = null;
        File_Data objFile_Data = null;
        #endregion

        #region .... Property ...
        public string msgError { get; set; }
        #endregion

        #region .... Constructor ....
        //public ClsFile_Server_Operations(string szAppXmlPath, string szLocation)
        //{
        //    msgError = string.Empty;
        //    try
        //    {
        //        #region ... Read Application Variables .....

        //        _szAppXmlPath = szAppXmlPath;
        //        _szLocation = szLocation;
        //        _objINI = new clsReadAppXml(_szAppXmlPath);
        //        _szServerIP = _objINI.GetLocationVariable(_szLocation, "", "Selfip");
        //        if (string.IsNullOrEmpty(_szServerIP))
        //            _szServerIP = _objINI.GetLocationVariable(_objINI.GetCurrentLocation(), "", "Selfip");

        //        #endregion

        //        #region ..... Binding Configuration .....

        //        _binding = new NetTcpBinding();
        //        _binding.SendTimeout = new TimeSpan(0, 10, 0);
        //        _binding.ReceiveTimeout = new TimeSpan(0, 10, 0);
        //        _binding.OpenTimeout = new TimeSpan(0, 10, 0);
        //        _binding.CloseTimeout = new TimeSpan(0, 10, 0);
        //        _binding.MaxBufferPoolSize = Int32.MaxValue;
        //        _binding.MaxBufferSize = Int32.MaxValue;
        //        _binding.MaxReceivedMessageSize = Int32.MaxValue;
        //        _binding.ReaderQuotas.MaxDepth = Int32.MaxValue;
        //        _binding.ReaderQuotas.MaxStringContentLength = Int32.MaxValue;
        //        _binding.ReaderQuotas.MaxArrayLength = Int32.MaxValue;
        //        _binding.ReaderQuotas.MaxBytesPerRead = Int32.MaxValue;
        //        _binding.ReaderQuotas.MaxNameTableCharCount = Int32.MaxValue;

        //        #endregion

        //        #region .... End point Configuration ....

        //        _endpoint = new EndpointAddress("net.tcp://" + _szServerIP + "//" + _szLocation + "_FileServer");

        //        #endregion
        //    }
        //    catch (Exception ex)
        //    {
        //        msgError = ex.Message;
        //    }
        //    finally
        //    {
        //        if (_objINI != null)
        //            _objINI.Dispose();
        //        _objINI = null;
        //    }


        //}
        //public ClsFile_Server_Operations(string szDBName, string szAppXmlPath, string szLocation)
        //{
        //    msgError = string.Empty;
        //    try
        //    {
        //        #region ..... get Application Variables .....

        //        _szAppXmlPath = szAppXmlPath;
        //        _szDBName = szDBName;
        //        _szLocation = szLocation;
        //        _bIsBackEndProcess = true;
        //        _objINI = new clsReadAppXml(_szAppXmlPath);
        //        _objINI.IsBackEnd = _bIsBackEndProcess;
        //        _szServerIP = _objINI.GetLocationVariable(_szLocation, "", "Selfip");
        //        if (string.IsNullOrEmpty(_szServerIP))
        //            _szServerIP = _objINI.GetLocationVariable(_objINI.GetCurrentLocation(), "", "Selfip");

        //        #endregion

        //        #region ..... Binding Configuration .....
        //        _binding = new NetTcpBinding();
        //        _binding.SendTimeout = new TimeSpan(0, 10, 0);
        //        _binding.ReceiveTimeout = new TimeSpan(0, 10, 0);
        //        _binding.OpenTimeout = new TimeSpan(0, 10, 0);
        //        _binding.CloseTimeout = new TimeSpan(0, 10, 0);
        //        _binding.MaxBufferPoolSize = Int32.MaxValue;
        //        _binding.MaxBufferSize = Int32.MaxValue;
        //        _binding.MaxReceivedMessageSize = Int32.MaxValue;
        //        _binding.ReaderQuotas.MaxDepth = Int32.MaxValue;
        //        _binding.ReaderQuotas.MaxStringContentLength = Int32.MaxValue;
        //        _binding.ReaderQuotas.MaxArrayLength = Int32.MaxValue;
        //        _binding.ReaderQuotas.MaxBytesPerRead = Int32.MaxValue;
        //        _binding.ReaderQuotas.MaxNameTableCharCount = Int32.MaxValue;
        //        #endregion

        //        #region .... EndPoints ...
        //        _endpoint = new EndpointAddress("net.tcp://" + _szServerIP + "//" + _szLocation + "_FileServer");
        //        #endregion

        //    }
        //    catch (Exception ex)
        //    {
        //        msgError = ex.Message;
        //    }
        //    finally
        //    {
        //        if (_objINI != null)
        //            _objINI.Dispose();
        //        _objINI = null;
        //    }


        //}

        public ClsFile_Server_Operations(string szAppXmlPath, string szLocation)
        {
            msgError = string.Empty;
            try
            {
                #region ... Read Application Variables .....

                _szAppXmlPath = szAppXmlPath;
                _szLocation = szLocation;
                _objINI = new clsReadAppXml(_szAppXmlPath);
                _szServerIP = _objINI.GetLocationVariable(_szLocation, "", "FileServer");
                if (string.IsNullOrEmpty(_szServerIP))
                    _szServerIP = _objINI.GetLocationVariable(_objINI.GetCurrentLocation(), "", "FileServer");

                #endregion

                #region ..... Binding Configuration .....

                _binding = new WSHttpBinding();
                _binding.SendTimeout = new TimeSpan(0, 30, 0);
                _binding.ReceiveTimeout = new TimeSpan(0, 30, 0);
                _binding.OpenTimeout = new TimeSpan(0, 30, 0);
                _binding.CloseTimeout = new TimeSpan(0, 30, 0);
                _binding.MaxBufferPoolSize = Int32.MaxValue;
                _binding.MaxReceivedMessageSize = Int32.MaxValue;
                
                _binding.ReaderQuotas.MaxDepth = Int32.MaxValue;
                _binding.ReaderQuotas.MaxStringContentLength = Int32.MaxValue;
                _binding.ReaderQuotas.MaxArrayLength = Int32.MaxValue;
                _binding.ReaderQuotas.MaxBytesPerRead = Int32.MaxValue;
                _binding.ReaderQuotas.MaxNameTableCharCount = Int32.MaxValue;

                _binding.Security.Mode = SecurityMode.None;
                _binding.TransactionFlow = true;
                //_binding.Security.Transport.ClientCredentialType = HttpClientCredentialType.None;


                #endregion

                #region .... End point Configuration ....

                _endpoint = new EndpointAddress(_szServerIP + "\\" + "Service1.svc");

                #endregion
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            finally
            {
                if (_objINI != null)
                    _objINI.Dispose();
                _objINI = null;
            }


        }
        public ClsFile_Server_Operations(string szDBName, string szAppXmlPath, string szLocation)
        {
            msgError = string.Empty;
            try
            {
                #region ..... get Application Variables .....

                _szAppXmlPath = szAppXmlPath;
                _szDBName = szDBName;
                _szLocation = szLocation;
                _bIsBackEndProcess = true;
                _objINI = new clsReadAppXml(_szAppXmlPath);
                _objINI.IsBackEnd = _bIsBackEndProcess;
                _szServerIP = _objINI.GetLocationVariable(_szLocation, "", "FileServer");
                if (string.IsNullOrEmpty(_szServerIP))
                    _szServerIP = _objINI.GetLocationVariable(_objINI.GetCurrentLocation(), "", "FileServer");

                #endregion

                #region ..... Binding Configuration .....
                _binding = new WSHttpBinding();
                _binding.SendTimeout = new TimeSpan(0, 30, 0);
                _binding.ReceiveTimeout = new TimeSpan(0, 30, 0);
                _binding.OpenTimeout = new TimeSpan(0, 30, 0);
                _binding.CloseTimeout = new TimeSpan(0, 30, 0);
                _binding.MaxBufferPoolSize = Int32.MaxValue;
                _binding.MaxReceivedMessageSize = Int32.MaxValue;
                _binding.ReaderQuotas.MaxDepth = Int32.MaxValue;
                _binding.ReaderQuotas.MaxStringContentLength = Int32.MaxValue;
                _binding.ReaderQuotas.MaxArrayLength = Int32.MaxValue;
                _binding.ReaderQuotas.MaxBytesPerRead = Int32.MaxValue;
                _binding.ReaderQuotas.MaxNameTableCharCount = Int32.MaxValue;
                _binding.Security.Mode = SecurityMode.None;
                _binding.TransactionFlow = true;




                #endregion

                #region .... EndPoints ...
                _endpoint = new EndpointAddress(_szServerIP + "\\" + "Service1.svc");
                #endregion

            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            finally
            {
                if (_objINI != null)
                    _objINI.Dispose();
                _objINI = null;
            }


        }

        #endregion

        #region .... Public Functions ...
        public List<File_Data> Copy_File(List<File_Data> lstFile_Data)
        {
            ChannelFactory<IService1> Channel = null;
            File_Data[] arrFileData = lstFile_Data.ToArray();
            try
            {
                Channel = new ChannelFactory<IService1>(_binding, _endpoint);
                IService1 proxy = Channel.CreateChannel();
                arrFileData = proxy.ListOfFiles(arrFileData);
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            finally
            {
                if (Channel != null)
                {
                    Channel.Close();
                    Channel.Abort();
                }
                Channel = null;
            }
            return lstFile_Data;
        }

        public File_Data Copy_File(File_Data objFile_Data)
        {
            ChannelFactory<IService1> Channel = null;
            try
            {
                Channel = new ChannelFactory<IService1>(_binding, _endpoint);
                IService1 proxy = Channel.CreateChannel();
                objFile_Data = proxy.CopyFile(objFile_Data);
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            finally
            {
                if (Channel != null)
                {
                    Channel.Close();
                    Channel.Abort();
                }
                Channel = null;
            }
            return objFile_Data;
        }
        public bool Check_File_Exist(File_Data objFile_Data)
        {
            bool bResult = false;
            ChannelFactory<IService1> Channel = null;
            try
            {
                if (string.IsNullOrEmpty(objFile_Data.Source_Directory))
                    throw new Exception("Please Provide Source Directory");

                Channel = new ChannelFactory<IService1>(_binding, _endpoint);
                IService1 proxy = Channel.CreateChannel();
                bResult = proxy.Check_File_Exist(objFile_Data);
            }
            catch (Exception ex)
            {
                bResult = false;
                msgError = ex.Message;
            }
            finally
            {
                if (Channel != null)
                {
                    Channel.Close();
                    Channel.Abort();
                }
                Channel = null;
            }
            return bResult;

        }
        public bool Check_File_Is_Locked(string szFileName)
        {
            bool bResult = false;
            ChannelFactory<IService1> Channel = null;
            try
            {
                Channel = new ChannelFactory<IService1>(_binding, _endpoint);
                IService1 proxy = Channel.CreateChannel();
                bResult = proxy.Check_File_Is_Locked(szFileName);
            }
            catch (Exception ex)
            {
                bResult = true;
                msgError = ex.Message;
            }
            finally
            {
                if (Channel != null)
                {
                    Channel.Close();
                    Channel.Abort();
                }
                Channel = null;
            }
            return bResult;

        }
        public bool Pre_Check_File(string szFileName)
        {
            bool bResult = true;
            ChannelFactory<IService1> Channel = null;
            try
            {
                Channel = new ChannelFactory<IService1>(_binding, _endpoint);
                IService1 proxy = Channel.CreateChannel();
                bResult = proxy.Pre_Check_File(szFileName);
            }
            catch (Exception ex)
            {
                bResult = false;
                msgError = ex.Message;
            }
            finally
            {
                if (Channel != null)
                {
                    Channel.Close();
                    Channel.Abort();
                }
                Channel = null;
            }
            return bResult;

        }
        public bool Pre_Check_File(byte[] arrFile)
        {
            bool bResult = true;
            ChannelFactory<IService1> Channel = null;
            try
            {
                Channel = new ChannelFactory<IService1>(_binding, _endpoint);
                IService1 proxy = Channel.CreateChannel();
                bResult = proxy.Pre_Check_File_Blob(arrFile);
            }
            catch (Exception ex)
            {
                bResult = false;
                msgError = ex.Message;
            }
            finally
            {
                if (Channel != null)
                {
                    Channel.Close();
                    Channel.Abort();
                }
                Channel = null;
            }
            return bResult;

        }

        public bool Delete_File(File_Data objFile_Data)
        {
            bool bResult = false;
            ChannelFactory<IService1> Channel = null;
            try
            {
                if (string.IsNullOrEmpty(objFile_Data.Source_Directory))
                    throw new Exception("Please Provide Source Directory");


                Channel = new ChannelFactory<IService1>(_binding, _endpoint);
                IService1 proxy = Channel.CreateChannel();
                bResult = proxy.Delete_File(objFile_Data);

            }
            catch (Exception ex)
            {
                bResult = false;
                msgError = ex.Message;
            }
            finally
            {
                if (Channel != null)
                {
                    Channel.Close();
                    Channel.Abort();
                }

                Channel = null;
            }
            return bResult;

        }
        public File_Data Get_File_Information(File_Data objFile_Data)
        {
            ChannelFactory<IService1> Channel = null;
            try
            {
                if (string.IsNullOrEmpty(objFile_Data.Source_Directory))
                    throw new Exception("Please Provide Source Directory");

                Channel = new ChannelFactory<IService1>(_binding, _endpoint);
                IService1 proxy = Channel.CreateChannel();
                objFile_Data = proxy.Get_File_Information(objFile_Data);
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            finally
            {
                if (Channel != null)
                {
                    Channel.Close();
                    Channel.Abort();
                }
                Channel = null;
            }
            return objFile_Data;

        }
        public List<File_Data> Get_Documents(File_Data objFile_Data)
        {
            ChannelFactory<IService1> Channel = null;
            List<File_Data> lstFile_Data = new List<File_Data>();
            try
            {
                if (string.IsNullOrEmpty(objFile_Data.Source_Directory))
                    throw new Exception("Please Provide Source Directory");

                Channel = new ChannelFactory<IService1>(_binding, _endpoint);
                IService1 proxy = Channel.CreateChannel();
                lstFile_Data = proxy.Get_Documents(objFile_Data).Cast<File_Data>().ToList();
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            finally
            {
                if (Channel != null)
                {
                    Channel.Close();
                    Channel.Abort();
                }
                Channel = null;
            }
            return lstFile_Data;

        }

        public string Get_Server_Date_Time()
        {
            ChannelFactory<IService1> Channel = null;
            string szServer_Date_Time = string.Empty;
            try
            {
                Channel = new ChannelFactory<IService1>(_binding, _endpoint);
                IService1 proxy = Channel.CreateChannel();
                szServer_Date_Time = proxy.Get_Server_Date_Time();
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            finally
            {
                if (Channel != null)
                {
                    Channel.Close();
                    Channel.Abort();
                }
                Channel = null;
            }
            return szServer_Date_Time;
        }
        public string Get_Server_Time()
        {
            ChannelFactory<IService1> Channel = null;
            string szServer_Date_Time = string.Empty;
            try
            {
                Channel = new ChannelFactory<IService1>(_binding, _endpoint);
                IService1 proxy = Channel.CreateChannel();
                szServer_Date_Time = proxy.Get_Server_Time();
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            finally
            {
                if (Channel != null)
                {
                    Channel.Close();
                    Channel.Abort();
                }
                Channel = null;
            }
            return szServer_Date_Time;
        }

        public string Get_Document_Checksum(File_Data objFile_Data)
        {
            ChannelFactory<IService1> Channel = null;
            string szDocumentCheckSum = string.Empty;
            try
            {
                if (string.IsNullOrEmpty(objFile_Data.Source_Directory))
                    throw new Exception("Please Provide Source Directory");

                Channel = new ChannelFactory<IService1>(_binding, _endpoint);
                IService1 proxy = Channel.CreateChannel();
                szDocumentCheckSum = proxy.Get_Document_CheckSum(objFile_Data);
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            finally
            {
                if (Channel != null)
                {
                    Channel.Close();
                    Channel.Abort();
                }
                Channel = null;
            }
            return szDocumentCheckSum;

        }

        #endregion

        #region .... IDISPOSABLE ....
        public void Dispose()
        {
            Dispose(true);
            GC.Collect();
            GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_objINI != null)
                    _objINI.Dispose();
                _objINI = null;
                _binding = null;
                _endpoint = null;
                if (objFile_Data != null)
                    objFile_Data.Data = null;
                objFile_Data = null;
                _szAppXmlPath = string.Empty;
                _szDBName = string.Empty;
                _szServerIP = string.Empty;
                _szLocation = string.Empty;
                _szTempSharedPath = string.Empty;
            }
            else
            {

            }
        }

        ~ClsFile_Server_Operations()
        {
            Dispose(false);
        }


        #endregion
    }
}
