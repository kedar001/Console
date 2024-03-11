using System;
using System.Collections.Generic;
using System.Text;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
//using System.ComponentModel;

namespace GetConInfo
{
    public class ClsConInfo
    {
        #region .... Variable Declaration ....

        //GetIniInfo.ClsGetIniInfo objGetIni;
        eDocsDN_ReadAppXml.clsReadAppXml objReadXml;
        Uri _uriConnAppSite;

        string _szAppXmlPath;
        string _szConnUrlPath;
        string _szError;

        #endregion

        #region .... Constructor ....

        public ClsConInfo(Uri uriConnAppSite)
        {
            //IntPtr iPtrHandle,
            //this.handle = iPtrHandle;
            msgError = "";
            //----------------------------------------------------------------------------------------------//
            this._szAppXmlPath = "";
            this._uriConnAppSite = uriConnAppSite;

            if (_uriConnAppSite != null)
                _szConnUrlPath = _uriConnAppSite.OriginalString;

            if (!_szConnUrlPath.Trim().EndsWith("/"))
                _szConnUrlPath = _szConnUrlPath + "/";
            //----------------------------------------------------------------------------------------------//
        }

        public ClsConInfo(string szAppXmlPath)
        {
            //IntPtr iPtrHandle,
            //this.handle = iPtrHandle;
            msgError = "";
            this._szAppXmlPath = szAppXmlPath;
            this._uriConnAppSite = null;
            try
            {
                //----------------------------------------------------------------------------------------------//
                if (string.IsNullOrEmpty(_szAppXmlPath))
                    throw new Exception("Application Xml Path is Null or Empty: " + _szAppXmlPath);

                //objGetIni = new GetIniInfo.ClsGetIniInfo(_szAppXmlPath);
                //_szConnUrlPath = objGetIni.GetIniInfo("APPLICATION INFO", "COMMON_APPLICATION_PATH");
                //objGetIni.Dispose();
                //objGetIni = null;

                objReadXml = new eDocsDN_ReadAppXml.clsReadAppXml(_szAppXmlPath);
                _szConnUrlPath = objReadXml.GetApplicationVariable("CommonAppPath");
                if (objReadXml.ErrorMsg != "")
                    throw new Exception("Error while reading Common Application Path :" + objReadXml.ErrorMsg);
                objReadXml.Dispose();
                objReadXml = null;

                if (string.IsNullOrEmpty(_szConnUrlPath) || _szConnUrlPath.Trim() == "")
                    throw new Exception("Connection Path is Null or Empty: " + _szConnUrlPath);

                if (!_szConnUrlPath.Trim().EndsWith("/"))
                    _szConnUrlPath = _szConnUrlPath + "/";
                //----------------------------------------------------------------------------------------------//
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            finally
            {
                //if (objGetIni != null)
                //    objGetIni.Dispose();
                //objGetIni = null;

                if (objReadXml != null)
                    objReadXml.Dispose();
                objReadXml = null;
            }
        }

        #endregion

        # region .... Property ....

        public string msgError
        {
            get { return _szError; }
            set { _szError = value; }
        }

        #endregion

        #region .... Function Definition ....

        public string GetConnectionString(string szProductName, string szRegion, string szConnectionType)
        {
            msgError = "";
            string szConnString = "";
            string szErrorMsg = "";
            WR_GetAppInfo.ClsGetAppInfo objWrGetAppInfo = null;
            try
            {
                if (string.IsNullOrEmpty(szProductName) || szProductName.Trim() == "")
                    throw new Exception("Product Name is Null or Empty: " + szProductName);

                if (!IsWebPathExist(_szConnUrlPath))
                    throw new Exception(msgError);

                objWrGetAppInfo = new WR_GetAppInfo.ClsGetAppInfo();
                //... Code Changed by manav on 13-12-2012 for DOE DRT-4136 ...
                //objWrGetAppInfo.Url = _szConnUrlPath + "code/WS_GetAppInfo.asmx";
                ServicePointManager.ServerCertificateValidationCallback += delegate (Object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
                {
                    //... Check certificate ...
                    return true;
                };

                objWrGetAppInfo.Url = _szConnUrlPath + "WS_GetAppInfo.asmx";
                szConnString = objWrGetAppInfo.GetConnectionString(szProductName, szRegion, szConnectionType, out szErrorMsg);
                if (szErrorMsg.Trim() != "")
                    throw new Exception(szErrorMsg);

                objWrGetAppInfo.Abort();
                objWrGetAppInfo.Dispose();
                objWrGetAppInfo = null;
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            finally
            {
                if (objWrGetAppInfo != null)
                {
                    objWrGetAppInfo.Dispose();
                    objWrGetAppInfo = null;
                }
            }
            return szConnString;
        }

        //.... Function commneted and added for DRT- on 21/8/2014 by Harshad ....
        //private bool IsWebPathExist(string szUrlPath)
        //{
        //    bool bFileExist = true;
        //    HttpWebRequest wbRequest = null;
        //    HttpWebResponse wbResponse = null;
        //    try
        //    {
        //        if (string.IsNullOrEmpty(szUrlPath) || szUrlPath.Trim() == "")
        //            throw new Exception("Invalid URL - " + szUrlPath);

        //        wbRequest = (HttpWebRequest)WebRequest.Create(szUrlPath);
        //        wbResponse = (HttpWebResponse)wbRequest.GetResponse();
        //        wbRequest.Abort();
        //        wbRequest = null;
        //        wbResponse.Close();
        //        wbResponse = null;
        //    }
        //    catch (Exception ex)
        //    {
        //        msgError = "Connection Path not found: " + ex.Message + " Stacktrace: " + ex.StackTrace;
        //        msgError = msgError.Replace(Convert.ToChar(13), ' ');
        //        msgError = msgError.Replace(Convert.ToChar(10), ' ');

        //        bFileExist = false;
        //    }
        //    finally
        //    {
        //        if (wbRequest != null)
        //            wbRequest.Abort();

        //        if (wbResponse != null)
        //            wbResponse.Close();

        //        wbRequest = null;
        //        wbResponse = null;
        //    }
        //    return bFileExist;
        //}


        //.... Function added for DRT- on 21/8/2014 by Harshad ....
        private bool IsWebPathExist(string szUrlPath)
        {

            bool bFileExist = true;

            HttpWebRequest wbRequest = null;

            HttpWebResponse wbResponse = null;

            try
            {

                if (string.IsNullOrEmpty(szUrlPath))

                    throw new Exception("Invalid URL - " + szUrlPath);

                ServicePointManager.ServerCertificateValidationCallback += delegate (Object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
                {
                    //... Check certificate ...
                    return true;
                };

                wbRequest = (HttpWebRequest)WebRequest.Create(szUrlPath);

                wbRequest.Method = "HEAD";

                wbResponse = (HttpWebResponse)wbRequest.GetResponse();

                wbRequest.Abort();

                wbRequest = null;

                wbResponse.Close();

                wbResponse = null;

            }

            catch (Exception ex)
            {

                msgError = "Connection Path not found: " + ex.Message + " Stacktrace: " + ex.StackTrace;

                msgError = msgError.Replace(Convert.ToChar(13), ' ');

                msgError = msgError.Replace(Convert.ToChar(10), ' ');



                bFileExist = false;

            }

            finally
            {

                if (wbRequest != null)

                    wbRequest.Abort();



                if (wbResponse != null)

                    wbResponse.Close();



                wbRequest = null;

                wbResponse = null;

            }

            return bFileExist;

        }

        #endregion

        #region .... Functions for IDisposable Interface ....

        #region Variable Declaration for Disposable Object ...

        //private IntPtr handle;
        //private Component component = new Component();
        private bool bDisposed = false;

        #endregion

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        //[System.Runtime.InteropServices.DllImport("Kernel32")]
        //private extern static Boolean CloseHandle(IntPtr handle);

        ~ClsConInfo()
        {
            Dispose(false);
        }

        private void Dispose(bool bDisposing)
        {
            if (!this.bDisposed)
            {
                if (bDisposing)
                {
                    //if (objGetIni != null)
                    //    objGetIni.Dispose();

                    if (objReadXml != null)
                        objReadXml.Dispose();
                    //component.Dispose();
                }

                //objGetIni = null;
                objReadXml = null;
                msgError = null;
                _szAppXmlPath = null;
                _szConnUrlPath = null;
                _uriConnAppSite = null;
                //CloseHandle(handle);
                //handle = IntPtr.Zero;
                bDisposed = true;
            }
        }

        #endregion
    }
}