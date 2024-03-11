//... Last Code Changed by manav on 01-03-2014 for DR-919351 ...

using System.IO;
using System;
using System.Xml;
using System.Collections;
using System.Data;
using System.Web;
using DDLLCS;
using eDocsDN_ReadAppXml;

namespace eDocsDN_Common_LogWriter
{
    public class clsLogWriter
    {
        # region .... Variable Declaration ....

        clsReadAppXml objReadXml = null;
        ClsBuildQuery objDal = null;
        StreamWriter strWritter;
        TextWriter txtWritter;
        DateTime dtStartDateTime;
        DateTime dtEndDateTime;

        bool _bResult;
        long _iSurrogateKey = 0;
        string _szLogFolderPath;
        string _szError;
        string _szDbName = "";
        string _szAppXmlPath = "";
        string _szQuery = "";

        public enum FileMode
        {
            CreateNew = 1,
            AppendText = 2
        }

        public string szSurrKey = "", szlogfolderpath = "";
        public bool bConsoleFlag = false;

        #endregion

        # region ...... Property .......

        public string msgError
        {
            get { return _szError; }
            set { _szError = value; }
        }

        public long SurrogateKey
        {
            get { return _iSurrogateKey; }
        }

        #endregion

        # region .... Constructor ....

        public clsLogWriter(string szAppXmlPath)
        {
            msgError = "";
            this._szAppXmlPath = szAppXmlPath;
            objDal = new ClsBuildQuery(szAppXmlPath);
            objDal.OpenConnection();

            objReadXml = new clsReadAppXml(szAppXmlPath);
            _szLogFolderPath = objReadXml.GetApplicationVariable("LogDir");
            string szAppPath = HttpContext.Current.Request.PhysicalApplicationPath;
            DirectoryInfo DirInfo = new DirectoryInfo(szAppPath);
            string szDrive = DirInfo.Root.ToString();
            _szLogFolderPath = szDrive + _szLogFolderPath;

            if (!_szLogFolderPath.Trim().EndsWith("\\"))
                _szLogFolderPath = _szLogFolderPath + "\\";

            szlogfolderpath = _szLogFolderPath;
        }

        public clsLogWriter(string szDBName, string szAppXmlPath)
        {
            msgError = "";
            this._szDbName = szDBName;
            this._szAppXmlPath = szAppXmlPath;
            objDal = new ClsBuildQuery(szDBName, szAppXmlPath);
            objDal.OpenConnection();

            objReadXml = new clsReadAppXml(szAppXmlPath);
            _szLogFolderPath = objReadXml.GetApplicationVariable("LogDir");
            string szAppPath = HttpContext.Current.Request.PhysicalApplicationPath;
            DirectoryInfo DirInfo = new DirectoryInfo(szAppPath);
            string szDrive = DirInfo.Root.ToString();
            _szLogFolderPath = szDrive + _szLogFolderPath;

            if (!_szLogFolderPath.Trim().EndsWith("\\"))
                _szLogFolderPath = _szLogFolderPath + "\\";

            szlogfolderpath = _szLogFolderPath;
        }

        /// <summary>
        /// Will be used only for backend
        /// </summary>
        /// <param name="objDAL"></param>
        /// <param name="szApplicationPath"></param>
        public clsLogWriter(ClsBuildQuery objDAL, string szApplicationPath)
        {
            msgError = "";
            this.objDal = objDAL;
            //objDAL.OpenConnection();

            objReadXml = new clsReadAppXml(szApplicationPath);
            _szLogFolderPath = objReadXml.GetApplicationVariable("LogDir");
            ////string szAppPath = HttpContext.Current.Request.PhysicalApplicationPath;
            ////DirectoryInfo DirInfo = new DirectoryInfo(szAppPath);
            ////string szDrive = DirInfo.Root.ToString();
            string szDrive = Directory.GetDirectoryRoot(Environment.CurrentDirectory);
            _szLogFolderPath = szDrive + _szLogFolderPath;

            if (!_szLogFolderPath.Trim().EndsWith("\\"))
                _szLogFolderPath = _szLogFolderPath + "\\";

            szlogfolderpath = _szLogFolderPath;
        }

        #endregion

        #region ..... Functions Definition ......

        public bool AddLogHeader(string szProgType, string szProgKey, string szProgDesc, string szSessionId, string szStartDate, string szStartTime, string szUserID, string szUserPwd)
        {
            _bResult = true;
            _iSurrogateKey = 0;
            msgError = "";
            try
            {
                _szQuery = "INSERT INTO zespl_redaeh_gol (epyt_gorp, yek_gorp, csed_gorp, di_noisses, di_resu, dwp_resu, td_trats, mt_trats) " +
                           "VALUES ('" + szProgType + "','" + szProgKey + "','" + szProgDesc + "','" + szSessionId + "','" + szUserID + "','" + szUserPwd + "','" + szStartDate + "','" + szStartTime + "') Select @@IDENTITY AS 'Identity'";
                _iSurrogateKey = objDal.GetSingleValue(_szQuery);
                if (objDal.msgError != "")
                    throw new Exception("Error while inserting data into log header : " + objDal.msgError);

                if (SurrogateKey == 0)
                    throw new Exception("Log header key not found : " + objDal.msgError);

                File.AppendAllText(_szLogFolderPath + SurrogateKey.ToString() + ".txt", SurrogateKey.ToString() + "-$-D-$-Log Header Record Added Successfully" + Environment.NewLine);
            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = ex.Message + " StackTrace: " + ex.StackTrace;
                msgError = msgError.Replace((char)10, ' ').Replace((char)13, ' ');

                try { File.AppendAllText(_szLogFolderPath + "Error.txt", "(" + DateTime.Now.ToShortDateString() + ")-$-E-$-Error while inserting record into Log Header :- " + msgError); }
                catch (Exception) { }
            }
            finally
            {
                szSurrKey = _iSurrogateKey.ToString();
            }
            return _bResult;
        }

        public void addLogHeader(string szApplicationPath, string szProgType, string szProgKey, string szProgDesc, string szSessionId, string szStartDate, string szStartTime, string szUserID, string szUserPwd)
        {
            AddLogHeader(szProgType, szProgKey, szProgDesc, szSessionId, szStartDate, szStartTime, szUserID, szUserPwd);
        }

        public void StartTime()
        {
            dtStartDateTime = DateTime.Now;

            //txtWritter.WriteLine("Start Time Is ----->> " + dtStartDateTime.ToString());
            txtWritter.WriteLine(SurrogateKey + "-$-D-$-Start Time Is ----->> " + dtStartDateTime.ToString());
            txtWritter.Flush();
        }

        public void EndTime()
        {
            dtEndDateTime = DateTime.Now;

            //txtWritter.WriteLine("End Time Is ----->> " + dtEndDateTime.ToString());
            txtWritter.WriteLine(SurrogateKey + "-$-D-$-End Time Is ----->> " + dtEndDateTime.ToString());
            txtWritter.Flush();
        }

        public void TotalTime()
        {
            //txtWritter.WriteLine("Total Time Is ---->> " + dtEndDateTime.Subtract(dtStartDateTime));
            txtWritter.WriteLine(SurrogateKey + "-$-D-$-Total Time Is ---->> " + dtEndDateTime.Subtract(dtStartDateTime));
            txtWritter.Flush();
        }

        public void writeToLog(string szString)
        {
            szString = szString.Replace((char)10, ' ').Replace((char)13, ' ');
            txtWritter.WriteLine(szString);
            txtWritter.Flush();
        }

        private bool CloseFile()
        {
            _bResult = true;
            try
            {
                txtWritter.Flush();
                strWritter.Flush();

                txtWritter.Close();
                strWritter.Close();

                txtWritter.Dispose();
                strWritter.Dispose();
            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = "Error occured while Closing file: " + ex.Message + " StackTrace: " + ex.StackTrace;
                msgError = msgError.Replace((char)10, ' ').Replace((char)13, ' ');
            }
            finally
            {
                txtWritter = null;
                strWritter = null;
            }
            return _bResult;
        }

        public void closeFile()
        {
            CloseFile();
        }

        public TextWriter openFile(string szFileName, int iMode)
        {
            // iMode = 1  - - - - - - - - - > Output Mode
            // iMode = 2  - - - - - - - - - > Append Mode
            ////// iMode = 3  - - - - - - - - - > Append Mode For Error that is not handled,Blank SurrKey.
            try
            {
                if (iMode == 1)
                {
                    // CreateNew = 1  - - - - - - - - - > Output Mode
                    strWritter = new StreamWriter(_szLogFolderPath + szFileName);
                }
                else
                {
                    // AppendText = 2  - - - - - - - - - > Append Mode
                    strWritter = File.AppendText(_szLogFolderPath + szFileName);
                }
                txtWritter = strWritter;
                txtWritter.Flush();
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
                return (null);
            }
            return (txtWritter);
        }

        public TextWriter openTimeFile(string szFilePath, int iMode)
        {
            // iMode = 1  - - - - - - - - - > Output Mode
            // iMode = 2  - - - - - - - - - > Append Mode
            try
            {
                if (iMode == 1)
                {
                    // CreateNew = 1  - - - - - - - - - > Output Mode
                    strWritter = new StreamWriter(szFilePath);
                }
                else
                {
                    // AppendText = 2  - - - - - - - - - > Append Mode
                    strWritter = File.AppendText(szFilePath);
                }
                txtWritter = strWritter;
                txtWritter.Flush();
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
                return (null);
            }
            return (txtWritter);
        }

        public TextWriter OpenLogFile(long iSurrogateKey, FileMode eFileMode)
        {
            try
            {
                switch (eFileMode)
                {
                    case FileMode.CreateNew: // CreateNew = 1  - - - - - - - - - > Output Mode
                        strWritter = new StreamWriter(_szLogFolderPath + iSurrogateKey.ToString() + ".txt");

                        break;
                    case FileMode.AppendText: // AppendText = 2  - - - - - - - - - > Append Mode
                        strWritter = File.AppendText(_szLogFolderPath + iSurrogateKey.ToString() + ".txt");
                        break;
                }
                txtWritter = strWritter;
                txtWritter.Flush();
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
                return (null);
            }
            return (txtWritter);
        }

        public TextWriter OpenTextFile(string szFileName, FileMode eFileMode)
        {
            try
            {
                switch (eFileMode)
                {
                    case FileMode.CreateNew: // CreateNew = 1  - - - - - - - - - > Output Mode
                        strWritter = new StreamWriter(_szLogFolderPath + szFileName);
                        break;
                    case FileMode.AppendText: // AppendText = 2  - - - - - - - - - > Append Mode
                        strWritter = File.AppendText(_szLogFolderPath + szFileName);
                        break;
                }
                txtWritter = strWritter;
                txtWritter.Flush();
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
                return (null);
            }
            return (txtWritter);
        }

        public string replaceString(string szOriginalString, string szFindString, string szReplaceString)
        {
            string szTargetString = "";
            //char ch;
            int i = 0;

            if (szOriginalString != null)
            {
                while (i < szOriginalString.Length)
                {
                    int iindex = szOriginalString.IndexOf(szFindString, i);

                    if (iindex != i)
                    {
                        szTargetString = szTargetString + szOriginalString[i];
                        i++;
                    }
                    else
                    {
                        szTargetString = szTargetString + szReplaceString;
                        i += szFindString.Length;
                    }//if (iindex != 0)
                }//while (i < szOriginalString.length())
                return (szTargetString);
            }//if (szOriginalString != null)
            else
                return ("");
        }

        public bool UpdateLogDetail()
        {
            _bResult = true;
            msgError = "";
            TextReader txtReader = null;
            string szFileName = "";
            try
            {
                // ===== Open LogSurrKey.Txt File ===== //
                szFileName = _szLogFolderPath + SurrogateKey + ".txt";
                txtReader = new StreamReader(szFileName);

                int iSrNo = 1;
                string szLogText = "";

                #region ... Get Max SrNo ...

                _szQuery = "Select Max(on_rs) as MaxSRNO From zespl_liated_gol where yek_rrus_gol = " + SurrogateKey.ToString();
                iSrNo = objDal.GetSingleValue(_szQuery);
                if (objDal.msgError != "")
                    throw new Exception("Error while getting max SrNo from log details : " + objDal.msgError);

                if (iSrNo <= 0)
                    iSrNo = 1;
                else
                    iSrNo = iSrNo + 1;

                #endregion

                //... Add Log Detail Header ...
                szLogText = "Start Of Log For Log-Surrogate Key: " + SurrogateKey.ToString();
                _szQuery = "Insert into zespl_liated_gol (yek_rrus_gol, on_rs, epyt_gol, txet_gol) Values (" + SurrogateKey + "," + iSrNo + ",'D','" + szLogText + "')";
                if (!objDal.ExecuteQuery(_szQuery))
                    throw new Exception("Error while updating log details : " + objDal.msgError);

                #region ... Update Log details from file ...

                // ===== Read File Line By Line While Not End Of File ===== //
                bool bErrorFlag = false;
                bool bWarningFlag = false;
                string szLogType = "";
                string szLogDescLine;
                string[] szSeperator = new string[] { "-$-" };
                string[] stLogInfo;
                while ((szLogDescLine = txtReader.ReadLine()) != null)
                {
                    iSrNo++;
                    stLogInfo = szLogDescLine.Split(szSeperator, StringSplitOptions.None);

                    //... Error Flag ...
                    if (stLogInfo[1] != null)
                        szLogType = stLogInfo[1];           // ==== 2nd Token ==== //

                    //... Code Changed By manav on 21-05-2012 for DRT-3871 ...
                    //if (szLogType.Trim().ToUpper() == "E")
                    //    bErrorFlag = true;
                    switch (szLogType.Trim().ToUpper())
                    {
                        case "E":
                            bErrorFlag = true;
                            break;
                        case "W":
                            bWarningFlag = true;
                            break;
                    }
                    //...

                    //... Log Description ...
                    if (stLogInfo[2] != null)
                    {
                        szLogText = stLogInfo[2];            // ==== 3rd Token ==== //
                        szLogText = szLogText.Replace("\'", "");
                    }

                    _szQuery = "Insert into zespl_liated_gol (yek_rrus_gol, on_rs, epyt_gol, txet_gol) Values (" + SurrogateKey + "," + iSrNo + ",'" + szLogType + "','" + szLogText + "')";
                    if (!objDal.ExecuteQuery(_szQuery))
                    {
                        if (szLogText.Length > 950)
                            szLogText = szLogText.Substring(0, 950);

                        _szQuery = "Insert into zespl_liated_gol (yek_rrus_gol, on_rs, epyt_gol, txet_gol) Values (" + SurrogateKey + "," + iSrNo + ",'" + szLogType + "','" + szLogText + "')";
                        if (!objDal.ExecuteQuery(_szQuery))
                            throw new Exception("Error while updating log details : " + objDal.msgError);
                    }
                }
                txtReader.Close();
                txtReader.Dispose();
                txtReader = null;

                #endregion

                iSrNo++;
                szLogText = "End Of Log For Log-Surrogate Key: " + SurrogateKey.ToString();
                _szQuery = "Insert into zespl_liated_gol (yek_rrus_gol, on_rs, epyt_gol, txet_gol) Values (" + SurrogateKey + "," + iSrNo + ",'D','" + szLogText + "')";
                objDal.ExecuteQuery(_szQuery);

                // ===== Update End Time in Log Header Table ===== //
                //_szQuery = "Update zespl_redaeh_gol Set epyt_gol = 'E', td_dne = '" + DateTime.Now.ToString("MM-dd-yyyy HH:mm:ss tt") + "', mt_dne = '" + DateTime.Now.ToString("HH:mm:ss tt") + "' Where yek_rrus_gol = " + SurrogateKey;
                if (bErrorFlag)
                    _szQuery = "Update zespl_redaeh_gol Set epyt_gol = 'E', td_dne = '" + DateTime.Now.ToString("MM-dd-yyyy HH:mm:ss tt") + "', mt_dne = '" + DateTime.Now.ToString("HH:mm:ss") + "' Where yek_rrus_gol = " + SurrogateKey;
                else if (bWarningFlag)
                    _szQuery = "Update zespl_redaeh_gol Set epyt_gol = 'W', td_dne = '" + DateTime.Now.ToString("MM-dd-yyyy HH:mm:ss tt") + "', mt_dne = '" + DateTime.Now.ToString("HH:mm:ss") + "' Where yek_rrus_gol = " + SurrogateKey;
                else
                    _szQuery = "Update zespl_redaeh_gol Set td_dne = '" + DateTime.Now.ToString("MM-dd-yyyy HH:mm:ss tt") + "', mt_dne = '" + DateTime.Now.ToString("HH:mm:ss") + "' Where yek_rrus_gol = " + SurrogateKey;
                if (!objDal.ExecuteQuery(_szQuery))
                    throw new Exception("Error while updating log header : " + objDal.msgError);

                if (File.Exists(szFileName)) File.Delete(szFileName);
            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = ex.Message + " StackTrace: " + ex.StackTrace;
                msgError = msgError.Replace((char)10, ' ').Replace((char)13, ' ');

                OpenLogFile(SurrogateKey, FileMode.AppendText);
                writeToLog(SurrogateKey + "-$-E-$-Error While Updatating Log : " + msgError);
                CloseFile();
            }
            finally
            {
                if (txtReader != null)
                {
                    txtReader.Close();
                    txtReader.Dispose();
                }
                txtReader = null;
                szFileName = null;
            }
            return _bResult;
        }

        public string updateLogDetail(string szApplicationPath, long iSurrogateKey)
        {
            msgError = "";
            _iSurrogateKey = iSurrogateKey;
            UpdateLogDetail();
            return msgError;
        }

        #endregion
    }
}