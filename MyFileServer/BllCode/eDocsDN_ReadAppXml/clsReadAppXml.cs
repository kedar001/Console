//... Changed by manavya on 22-07-2011 For DRT-1459 ...
using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;
using System.Web;
//using System.ComponentModel;

namespace eDocsDN_ReadAppXml
{
    public class clsReadAppXml : IDisposable
    {
        #region .... Variable Declaration ....

        XmlReader _xReader;
        ArrayList _lstMapExt = new ArrayList();
        Dictionary<string, string> _dicAppVariables = new Dictionary<string, string>();

        bool _bIsfs = false;
        bool _bIssecfs = false;
        bool _bProLocations;
        bool _bResult;
        bool _bIsBackEnd;
        string _szAppXmlPath;
        string _szError;

        #endregion

        #region .... Property ....

        public string ErrorMsg
        {
            get { return _szError; }
            set { _szError = value; }
        }

        public bool IsBackEnd
        {
            get { return _bIsBackEnd; }
            set { _bIsBackEnd = value; }
        }

        public Dictionary<string, string> AppVariables
        {
            get { return _dicAppVariables; }
        }

        public ArrayList IsWordDocument
        {
            get
            {
                _lstMapExt.Clear();
                char[] chSplit = { ',' };
                string szExt = "";
                szExt = GetApplicationVariable("Word");
                string[] arrExt = szExt.Split(chSplit);
                for (int iCnt = 0; iCnt < arrExt.Length; iCnt++)
                {
                    _lstMapExt.Add(arrExt[iCnt].ToUpper());
                }
                return _lstMapExt;
            }
        }

        public ArrayList IsExcelDocument
        {
            get
            {
                char[] chSplit = { ',' };
                string szExt = "";
                szExt = GetApplicationVariable("Excel");
                string[] arrExt = szExt.Split(chSplit);
                _lstMapExt.Clear();
                for (int iCnt = 0; iCnt < arrExt.Length; iCnt++)
                {
                    _lstMapExt.Add(arrExt[iCnt].ToUpper());
                }
                return _lstMapExt;
            }
        }

        #endregion

        #region .... Constructor ....

        public clsReadAppXml(string szAppXmlPath)
        {
            //IntPtr handle
            //this.handle = handle;
            ErrorMsg = "";
            _szAppXmlPath = szAppXmlPath;
        }

        #endregion

        #region .... Functions for reading all values ....

        public bool ReadApplicationVariable()
        {
            ErrorMsg = "";
            _xReader = null;

            XmlReader xmlAppReader = null;
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.IgnoreWhitespace = true;
            bool bAppvar = false;
            try
            {
                _xReader = XmlReader.Create(_szAppXmlPath, settings);

                while (_xReader.Read())
                {
                    _xReader.MoveToContent();
                    if (_xReader.NodeType == XmlNodeType.Element)
                    {
                        bAppvar = _xReader.ReadToFollowing("ApplicationVariable");
                        if (bAppvar)
                        {
                            xmlAppReader = _xReader.ReadSubtree();
                            ReadAppVariable(xmlAppReader);
                            break;
                        }
                    }
                }

                _bResult = true;
            }
            catch (Exception ex)
            {
                _bResult = false;
                ErrorMsg = ex.Message;
            }
            finally
            {
                if (xmlAppReader != null)
                    xmlAppReader.Close();
                xmlAppReader = null;

                if (_xReader != null)
                    _xReader.Close();
                _xReader = null;
                settings = null;
            }
            return _bResult;
        }

        public bool ReadLocationSetting(string szLocation, string szDepartment)
        {
            ErrorMsg = "";
            _xReader = null;
            XmlReader xmlLoc = null;
            XmlReader xmlLocalSetting = null;
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.IgnoreWhitespace = true;
            bool bLocvar = false, bLocal = false, bSecFs = false, bDfs = false;
            try
            {
                _xReader = XmlReader.Create(_szAppXmlPath, settings);
                while (_xReader.Read())
                {
                    _xReader.MoveToContent();
                    bLocvar = _xReader.ReadToFollowing("Location");
                    xmlLoc = _xReader.ReadSubtree();
                    if (xmlLoc.Read())
                    {
                        if (xmlLoc.GetAttribute("name").ToUpper() == szLocation.ToUpper())
                        {
                            if (xmlLoc.GetAttribute("isfs").ToUpper() == "Y")
                                _bIsfs = true;

                            if (xmlLoc.GetAttribute("issecfs").ToUpper() == "Y")
                                _bIssecfs = true;

                            bLocal = _xReader.ReadToFollowing("LocalSetting");
                            if (bLocal)
                            {
                                xmlLocalSetting = _xReader.ReadSubtree();
                                ReadAppVariable(xmlLocalSetting);
                            }

                            bDfs = _xReader.ReadToFollowing("DefaultFs");
                            if (bDfs)
                            {
                                xmlLocalSetting = _xReader.ReadSubtree();
                                ReadAppVariable(xmlLocalSetting);
                            }

                            if (_bIssecfs && szDepartment != "")
                            {
                                bSecFs = _xReader.ReadToFollowing("SecondaryFs");
                                if (bSecFs)
                                {
                                    ReadSecFs(_xReader, szDepartment);
                                }
                            }
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
            }
            finally
            {
                if (xmlLoc != null)
                    xmlLoc.Close();
                xmlLoc = null;

                if (xmlLocalSetting != null)
                    xmlLocalSetting.Close();
                xmlLocalSetting = null;

                if (_xReader != null)
                    _xReader.Close();
                _xReader = null;
                settings = null;
            }
            return _bResult;
        }

        public bool ReadAppDirectoryPath(string szLocation, string szDepartment)
        {
            ErrorMsg = "";
            XmlReader xReader_1 = null;
            XmlReader xmlLocalSetting = null;
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.IgnoreWhitespace = true;
            bool bAppDir = false;
            try
            {
                xReader_1 = XmlReader.Create(_szAppXmlPath, settings);
                ReadLocationSetting(szLocation, szDepartment);
                bAppDir = xReader_1.ReadToFollowing("AppDirectory");
                if (bAppDir)
                {
                    xmlLocalSetting = xReader_1.ReadSubtree();
                    ReadAppDir(xmlLocalSetting);
                }
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
            }
            finally
            {
                if (xmlLocalSetting != null)
                    xmlLocalSetting.Close();
                xmlLocalSetting = null;

                if (xReader_1 != null)
                    xReader_1.Close();
                xReader_1 = null;
                settings = null;
            }
            return _bResult;
        }

        private void ReadAppVariable(XmlReader xRead)
        {
            try
            {
                while (xRead.Read())
                {
                    xRead.MoveToContent();

                    if (xRead.HasAttributes)
                    {
                        _dicAppVariables.Add(xRead.Name, xRead.GetAttribute("value"));
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
            }
            finally
            {
                if (xRead != null)
                    xRead.Close();
                xRead = null;
            }
        }

        private void ReadSecFs(XmlReader xRead, string szDept)
        {
            try
            {
                while (xRead.Read())
                {
                    xRead.MoveToContent();

                    if (xRead.Name.ToUpper() == szDept.ToUpper())
                    {
                        _dicAppVariables.Add("SecFs", xRead.Name);
                        while (xRead.MoveToNextAttribute())
                        {
                            _dicAppVariables.Add(xRead.Name, xRead.Value);
                        }
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
            }
            finally
            {
                if (xRead != null)
                    xRead.Close();
                xRead = null;
            }
        }

        private void ReadAppDir(XmlReader xRead)
        {
            Dictionary<string, string> dicAppDir = new Dictionary<string, string>();
            Dictionary<string, string> dicAllAppDir = new Dictionary<string, string>();
            string szDfsPath = "", szSfsPath = "", szPhysicalPath = "";
            try
            {
                while (xRead.Read())
                {
                    xRead.MoveToContent();
                    if (xRead.HasAttributes)
                    {
                        dicAppDir.Add(xRead.Name, xRead.GetAttribute("value"));
                    }
                }
                xRead.Close();
                xRead = null;

                szDfsPath = "\\\\" + _dicAppVariables["DfsIp"] + "\\" + _dicAppVariables["Dfssharedfolder"];

                //... Code Changed By manav on 29/12/2011 for DRD-917953 ...
                szPhysicalPath = Get_DriveName() + _dicAppVariables["RootDir"];

                foreach (KeyValuePair<string, string> kv in dicAppDir)
                {
                    dicAllAppDir.Add(kv.Key, szPhysicalPath + kv.Value);
                    dicAllAppDir.Add("Dfs_" + kv.Key, szDfsPath + kv.Value);

                    if (_bIssecfs)
                    {
                        szSfsPath = "\\\\" + _dicAppVariables["sfsip"] + "\\" + _dicAppVariables["sharedfolder"];
                        dicAllAppDir.Add("Sfs_" + kv.Key, szSfsPath + kv.Value);
                    }
                }
                _dicAppVariables.Clear();
                _dicAppVariables = dicAllAppDir;
                dicAppDir = null;
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
            }
            finally
            {
                if (xRead != null)
                    xRead.Close();
                xRead = null;
            }
        }

        #endregion

        #region .... Function for getting key value ....

        public string GetApplicationVariable(string szReadValueOf)
        {
            ErrorMsg = "";
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.IgnoreWhitespace = true;
            XmlReader xreader = null;
            bool bAppvar = false;
            string szReturn = "";
            try
            {
                xreader = XmlReader.Create(_szAppXmlPath, settings);
                while (xreader.Read())
                {
                    xreader.MoveToContent();
                    if (xreader.NodeType == XmlNodeType.Element)
                    {
                        bAppvar = xreader.ReadToFollowing(szReadValueOf);
                        if (bAppvar)
                        {
                            szReturn = xreader.GetAttribute("value");
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _bResult = false;
                ErrorMsg = ex.Message;
            }
            finally
            {
                if (xreader != null)
                    xreader.Close();
                xreader = null;
                settings = null;
            }
            return szReturn;
        }

        public string GetLocationVariable(string szLocation, string szDepartment, string szReadValueOf)
        {
            ErrorMsg = "";
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.IgnoreWhitespace = true;
            XmlReader xreader = null;
            XmlReader xmlLoc = null;
            bool bLocvar = false;
            string szReturn = "";
            try
            {
                if (!CheckLocationExist(szLocation))
                {
                    szLocation = GetCurrentLocation();
                }

                xreader = XmlReader.Create(_szAppXmlPath, settings);
                while (xreader.Read())
                {
                    xreader.MoveToContent();
                    if (xreader.NodeType == XmlNodeType.Element)
                    {
                        bLocvar = xreader.ReadToFollowing("Location");
                        if (bLocvar)
                        {
                            xmlLoc = xreader.ReadSubtree();
                            if (xmlLoc.Read())
                            {
                                if (xmlLoc.GetAttribute("name").ToUpper() == szLocation.ToUpper())
                                {
                                    bLocvar = xreader.ReadToFollowing(szReadValueOf);
                                    if (bLocvar)
                                    {
                                        szReturn = xreader.GetAttribute("value");
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
            }
            finally
            {
                if (xmlLoc != null)
                    xmlLoc.Close();
                xmlLoc = null;

                if (xreader != null)
                    xreader.Close();
                xreader = null;
                settings = null;
            }
            return szReturn;
        }

        public string GetCurrentLocation()
        {
            ErrorMsg = "";
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.IgnoreWhitespace = true;
            XmlReader xreader = null;// XmlReader.Create(szAppXmlPath, settings);
            bool bAppvar = false;
            string szReturn = "";
            try
            {
                xreader = XmlReader.Create(_szAppXmlPath, settings);
                while (xreader.Read())
                {
                    xreader.MoveToContent();
                    if (xreader.NodeType == XmlNodeType.Element)
                    {
                        bAppvar = xreader.ReadToFollowing("CurrentLocation");
                        if (bAppvar)
                        {
                            szReturn = xreader.GetAttribute("name");
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _bResult = false;
                ErrorMsg = ex.Message;
            }
            finally
            {
                if (xreader != null)
                    xreader.Close();
                xreader = null;
                settings = null;
            }
            return szReturn;
        }

        public string GetAppLocation()
        {
            ErrorMsg = "";
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.IgnoreWhitespace = true;
            XmlReader xreader = null;// XmlReader.Create(szAppXmlPath, settings);
            bool bAppvar = false;
            string szReturn = "";
            try
            {
                xreader = XmlReader.Create(_szAppXmlPath, settings);
                while (xreader.Read())
                {
                    xreader.MoveToContent();
                    if (xreader.NodeType == XmlNodeType.Element)
                    {
                        bAppvar = xreader.ReadToFollowing("AppLocation");
                        if (bAppvar)
                        {
                            szReturn = xreader.GetAttribute("name");
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _bResult = false;
                ErrorMsg = ex.Message;
            }
            finally
            {
                if (xreader != null)
                    xreader.Close();
                xreader = null;
                settings = null;
            }
            return szReturn;
        }

        public string GetProcessLocations(string szLocation)
        {
            ErrorMsg = "";
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.IgnoreWhitespace = true;
            XmlReader xreader = null; // = XmlReader.Create(szAppXmlPath, settings);
            XmlReader xmlLoc = null;
            bool bLocvar = false;
            string szReturn = "";
            try
            {
                xreader = XmlReader.Create(_szAppXmlPath, settings);
                while (xreader.Read())
                {
                    xreader.MoveToContent();
                    bLocvar = xreader.ReadToFollowing("Location");
                    xmlLoc = xreader.ReadSubtree();
                    if (xmlLoc.Read())
                    {
                        if (xmlLoc.GetAttribute("name").ToUpper() == szLocation.ToUpper() &&
                            xmlLoc.GetAttribute("processotherloc").ToUpper() == "Y")
                        {
                            bLocvar = xmlLoc.ReadToFollowing("ProcessLocations");
                            if (bLocvar)
                            {
                                while (xmlLoc.Read())
                                {
                                    xmlLoc.MoveToContent();
                                    if (xmlLoc.NodeType == XmlNodeType.Element)
                                    {
                                        if (szReturn == "")
                                        {
                                            szReturn = "'" + xmlLoc.Name + "'";
                                        }
                                        else
                                        {
                                            szReturn = szReturn + " , '" + xmlLoc.Name + "'";
                                        }
                                    }
                                }
                                if (xmlLoc != null)
                                    xmlLoc.Close();
                                xmlLoc = null;
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
            }
            finally
            {
                if (xmlLoc != null)
                    xmlLoc.Close();
                xmlLoc = null;

                if (xreader != null)
                    xreader.Close();
                xreader = null;
                settings = null;
            }
            return szReturn;
        }

        private bool CheckLocationExist(string szLocation)
        {
            ErrorMsg = "";
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.IgnoreWhitespace = true;
            XmlReader xreader = null; // = XmlReader.Create(szAppXmlPath, settings);
            bool bLocvar = false;
            _bResult = false;
            try
            {
                xreader = XmlReader.Create(_szAppXmlPath, settings);
                while (xreader.Read())
                {
                    xreader.MoveToContent();
                    if (xreader.NodeType == XmlNodeType.Element)
                    {
                        bLocvar = xreader.ReadToFollowing("Location");
                        if (bLocvar)
                        {
                            if (xreader.GetAttribute("name").ToUpper() == szLocation.ToUpper())
                            {
                                _bResult = true;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
            }
            finally
            {
                if (xreader != null)
                    xreader.Close();
                xreader = null;
                settings = null;
            }
            return _bResult;
        }

        #endregion

        #region .... Functions To get Bridge Values ....

        public bool IsElememtTagExist(string szElementName)
        {
            ErrorMsg = "";
            _xReader = null;
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.IgnoreWhitespace = true;
            bool bAppvar = false;
            bool bElementExist = false;
            try
            {
                _xReader = XmlReader.Create(_szAppXmlPath, settings);

                while (_xReader.Read())
                {
                    _xReader.MoveToContent();
                    if (_xReader.NodeType == XmlNodeType.Element)
                    {
                        bAppvar = _xReader.ReadToFollowing(szElementName);
                        if (bAppvar)
                        {
                            bElementExist = true;
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
            }
            finally
            {
                if (_xReader != null)
                    _xReader.Close();
                _xReader = null;
                settings = null;
            }
            return bElementExist;
        }

        public bool IsDocsDossierBridgeExist()
        {
            ErrorMsg = "";
            _xReader = null;
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.IgnoreWhitespace = true;
            XmlReader xmlBrgReader = null;
            bool bAppvar = false;
            string szReturn = "";
            bool bBridgeExist = false;
            try
            {
                _xReader = XmlReader.Create(_szAppXmlPath, settings);

                while (_xReader.Read())
                {
                    _xReader.MoveToContent();
                    if (_xReader.NodeType == XmlNodeType.Element)
                    {
                        //... Code Changed by manav on 26-12-2012 for DRT-4136 ...
                        //bAppvar = _xReader.ReadToFollowing("Bridge");
                        //if (bAppvar)
                        //{
                        //    xmlBrgReader = _xReader.ReadSubtree();
                        //    while (xmlBrgReader.Read())
                        //    {
                        //        xmlBrgReader.MoveToContent();
                        //        if (xmlBrgReader.Name.Trim().Equals("IsBridgeExist", StringComparison.InvariantCultureIgnoreCase))
                        //        {
                        //            if (xmlBrgReader.HasAttributes)
                        //            {
                        //                szReturn = xmlBrgReader.GetAttribute("value");
                        //                break;
                        //            }
                        //        }
                        //    }
                        //    break;
                        //}

                        bAppvar = _xReader.ReadToFollowing("ApplicationVariable");
                        if (bAppvar)
                        {
                            xmlBrgReader = _xReader.ReadSubtree();
                            while (xmlBrgReader.Read())
                            {
                                xmlBrgReader.MoveToContent();
                                if (xmlBrgReader.Name.Trim().Equals("IsBridgeExist", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    if (xmlBrgReader.HasAttributes)
                                    {
                                        szReturn = xmlBrgReader.GetAttribute("value");
                                        break;
                                    }
                                }
                            }
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
            }
            finally
            {
                if (xmlBrgReader != null)
                    xmlBrgReader.Close();
                xmlBrgReader = null;

                if (_xReader != null)
                    _xReader.Close();
                _xReader = null;
                settings = null;

                if (szReturn.Trim() != "")
                {
                    if (szReturn.Trim().Equals("true", StringComparison.InvariantCultureIgnoreCase))
                        bBridgeExist = true;
                }
            }
            return bBridgeExist;
        }

        //public string GetBridgeVariable(string szReadValueOf)
        //{
        //    ErrorMsg = "";
        //    _xReader = null;
        //    XmlReaderSettings settings = new XmlReaderSettings();
        //    settings.IgnoreWhitespace = true;
        //    XmlReader xmlBrgReader = null;
        //    bool bAppvar = false;
        //    string szReturn = "";
        //    try
        //    {
        //        _xReader = XmlReader.Create(_szAppXmlPath, settings);

        //        while (_xReader.Read())
        //        {
        //            _xReader.MoveToContent();
        //            if (_xReader.NodeType == XmlNodeType.Element)
        //            {
        //                bAppvar = _xReader.ReadToFollowing("Bridge");
        //                if (bAppvar)
        //                {
        //                    xmlBrgReader = _xReader.ReadSubtree();
        //                    while (xmlBrgReader.Read())
        //                    {
        //                        xmlBrgReader.MoveToContent();
        //                        if (xmlBrgReader.Name.Trim().Equals(szReadValueOf.Trim(), StringComparison.InvariantCultureIgnoreCase))
        //                        {
        //                            if (xmlBrgReader.HasAttributes)
        //                            {
        //                                szReturn = xmlBrgReader.GetAttribute("value");
        //                                break;
        //                            }
        //                        }
        //                    }
        //                    break;
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        ErrorMsg = ex.Message;
        //    }
        //    finally
        //    {
        //        if (xmlBrgReader != null)
        //            xmlBrgReader.Close();
        //        xmlBrgReader = null;

        //        if (_xReader != null)
        //            _xReader.Close();
        //        _xReader = null;
        //        settings = null;
        //    }
        //    return szReturn;
        //}

        //public Dictionary<string, string> GetBridgeValues()
        //{
        //    ErrorMsg = "";
        //    _xReader = null;

        //    XmlReader xmlBrgReader = null;
        //    XmlReaderSettings settings = new XmlReaderSettings();
        //    settings.IgnoreWhitespace = true;
        //    bool bAppvar = false;
        //    Dictionary<string, string> Dic_BridgValues = new Dictionary<string, string>();
        //    try
        //    {
        //        _xReader = XmlReader.Create(_szAppXmlPath, settings);

        //        while (_xReader.Read())
        //        {
        //            _xReader.MoveToContent();
        //            if (_xReader.NodeType == XmlNodeType.Element)
        //            {
        //                bAppvar = _xReader.ReadToFollowing("Bridge");
        //                if (bAppvar)
        //                {
        //                    xmlBrgReader = _xReader.ReadSubtree();
        //                    while (xmlBrgReader.Read())
        //                    {
        //                        xmlBrgReader.MoveToContent();

        //                        if (xmlBrgReader.HasAttributes)
        //                        {
        //                            Dic_BridgValues.Add(xmlBrgReader.Name, xmlBrgReader.GetAttribute("value"));
        //                        }
        //                    }
        //                    break;
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        ErrorMsg = ex.Message;
        //    }
        //    finally
        //    {
        //        if (xmlBrgReader != null)
        //            xmlBrgReader.Close();
        //        xmlBrgReader = null;

        //        if (_xReader != null)
        //            _xReader.Close();
        //        _xReader = null;
        //        settings = null;
        //    }
        //    return Dic_BridgValues;
        //}

        public string GetBridgeApplicationPath()
        {
            ErrorMsg = "";
            _xReader = null;
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.IgnoreWhitespace = true;
            XmlReader xmlBrgReader = null;
            bool bAppvar = false;
            string szReturn = "";
            try
            {
                _xReader = XmlReader.Create(_szAppXmlPath, settings);

                while (_xReader.Read())
                {
                    _xReader.MoveToContent();
                    if (_xReader.NodeType == XmlNodeType.Element)
                    {
                        bAppvar = _xReader.ReadToFollowing("ApplicationVariable");
                        if (bAppvar)
                        {
                            xmlBrgReader = _xReader.ReadSubtree();
                            while (xmlBrgReader.Read())
                            {
                                xmlBrgReader.MoveToContent();
                                if (xmlBrgReader.Name.Trim().Equals("BridgeAppPath", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    if (xmlBrgReader.HasAttributes)
                                    {
                                        szReturn = xmlBrgReader.GetAttribute("value");
                                        break;
                                    }
                                }
                            }
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
            }
            finally
            {
                if (xmlBrgReader != null)
                    xmlBrgReader.Close();
                xmlBrgReader = null;

                if (_xReader != null)
                    _xReader.Close();
                _xReader = null;
                settings = null;
            }

            if (szReturn.Trim() != "" && !szReturn.Trim().EndsWith("/"))
                szReturn = szReturn.Trim() + "/";

            return szReturn;
        }

        public string GetCommonApplicationPath()
        {
            ErrorMsg = "";
            _xReader = null;
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.IgnoreWhitespace = true;
            XmlReader xmlBrgReader = null;
            bool bAppvar = false;
            string szReturn = "";
            try
            {
                _xReader = XmlReader.Create(_szAppXmlPath, settings);

                while (_xReader.Read())
                {
                    _xReader.MoveToContent();
                    if (_xReader.NodeType == XmlNodeType.Element)
                    {
                        bAppvar = _xReader.ReadToFollowing("ApplicationVariable");
                        if (bAppvar)
                        {
                            xmlBrgReader = _xReader.ReadSubtree();
                            while (xmlBrgReader.Read())
                            {
                                xmlBrgReader.MoveToContent();
                                if (xmlBrgReader.Name.Trim().Equals("CommonAppPath", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    if (xmlBrgReader.HasAttributes)
                                    {
                                        szReturn = xmlBrgReader.GetAttribute("value");
                                        break;
                                    }
                                }
                            }
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
            }
            finally
            {
                if (xmlBrgReader != null)
                    xmlBrgReader.Close();
                xmlBrgReader = null;

                if (_xReader != null)
                    _xReader.Close();
                _xReader = null;
                settings = null;
            }

            if (szReturn.Trim() != "" && !szReturn.Trim().EndsWith("/"))
                szReturn = szReturn.Trim() + "/";

            return szReturn;
        }

        #endregion

        #region .... New AppDirPathFunction Function ....

        public string GetAppDirPath(string szLocation, string szDepartment, string szReadValueOf)
        {
            ErrorMsg = "";
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.IgnoreWhitespace = true;
            XmlReader xreader = null;// XmlReader.Create(szAppXmlPath, settings);
            XmlReader xmlLoc = null;
            XmlReader xmlAppDir = null;

            bool bAppDir = false, bFalg = false, bLocvar = false;
            string szDrive = "";
            string szCurrentLocation = "", szRootDir = "";
            string szReturnPath = "";
            string szSecFsPath = "";
            try
            {
                szCurrentLocation = GetCurrentLocation();
                szRootDir = GetApplicationVariable("RootDir");

                //... code Changed By manav on 29/12/2011 for DRD-917953 ...
                szDrive = Get_DriveName();

                if (!CheckLocationExist(szLocation))
                {
                    szLocation = szCurrentLocation;
                }

                xreader = XmlReader.Create(_szAppXmlPath, settings);
                while (xreader.Read())
                {
                    xreader.MoveToContent();
                    bLocvar = xreader.ReadToFollowing("Location");
                    xmlLoc = xreader.ReadSubtree();
                    if (xmlLoc.Read())
                    {
                        if (xmlLoc.GetAttribute("name").ToUpper() == szCurrentLocation.ToUpper())
                        {
                            _bIsfs = false;
                            _bIssecfs = false;
                            _bProLocations = false;

                            if (xmlLoc.GetAttribute("isfs").ToUpper() == "Y")
                                _bIsfs = true;

                            if (xmlLoc.GetAttribute("issecfs").ToUpper() == "Y")
                                _bIssecfs = true;

                            if (xmlLoc.GetAttribute("processotherloc").ToUpper() == "Y")
                                _bProLocations = true;

                            //--------- Check Process another locations for backend ----------------------

                            if (IsBackEnd)
                            {
                                if (_bProLocations && szCurrentLocation.ToUpper() != szLocation.ToUpper())
                                {
                                    bLocvar = xmlLoc.ReadToFollowing("ProcessLocations");
                                    if (bLocvar)
                                    {
                                        szReturnPath = ReadProcessLocation(xmlLoc, szLocation);
                                        if (xmlLoc != null)
                                            xmlLoc.Close();
                                        xmlLoc = null;
                                    }
                                }
                            }
                            //-----------------------------------------------------------------

                            if (szReturnPath == "")
                            {
                                //------------- Get Physical path --------------------------------
                                if (_bIsfs && _bIssecfs == false)
                                {
                                    szReturnPath = szDrive + szRootDir;
                                    if (xmlLoc != null)
                                        xmlLoc.Close();
                                    xmlLoc = null;
                                }//if(bIsfs == true && bIssecfs == false)
                                // if department has sec fs then return that path else return physical path 
                                else if (_bIsfs && _bIssecfs)
                                {
                                    bLocvar = xmlLoc.ReadToFollowing("SecondaryFs");
                                    if (bLocvar)
                                    {
                                        szSecFsPath = ReadSecondaryFs(xmlLoc, szDepartment);
                                    }

                                    if (szSecFsPath != "")
                                        szReturnPath = szSecFsPath;
                                    else
                                        szReturnPath = szDrive + szRootDir;
                                    if (xmlLoc != null)
                                        xmlLoc.Close();
                                    xmlLoc = null;
                                } //if (bIsfs == true && bIssecfs == true)
                                // if department has no sec fs and current location is not file server then return default fs path 
                                else if (_bIsfs == false && _bIssecfs == false)
                                {
                                    bLocvar = xmlLoc.ReadToFollowing("DefaultFs");
                                    if (bLocvar)
                                    {
                                        szReturnPath = "\\\\" + xmlLoc.GetAttribute("ip") + "\\" + xmlLoc.GetAttribute("sharedfolder");
                                    }
                                    if (xmlLoc != null)
                                        xmlLoc.Close();
                                    xmlLoc = null;

                                }//if (bIsfs == false && bIssecfs == false)
                                else if (_bIsfs == false && _bIssecfs)
                                {
                                    bLocvar = xmlLoc.ReadToFollowing("DefaultFs");
                                    if (bLocvar)
                                    {
                                        szReturnPath = "\\\\" + xmlLoc.GetAttribute("ip") + "\\" + xmlLoc.GetAttribute("sharedfolder");
                                    }

                                    bLocvar = xmlLoc.ReadToFollowing("SecondaryFs");
                                    if (bLocvar)
                                    {
                                        szSecFsPath = ReadSecondaryFs(xmlLoc, szDepartment);
                                    }

                                    if (szSecFsPath != "")
                                    {
                                        szReturnPath = szSecFsPath;
                                    }
                                    if (xmlLoc != null)
                                        xmlLoc.Close();
                                    xmlLoc = null;
                                }// if (bIsfs == false && bIssecfs == true)
                            }

                            //----- Read Application Directory values from xml -----------------
                            bAppDir = xreader.ReadToFollowing("AppDirectory");
                            if (bAppDir)
                            {
                                xmlAppDir = xreader.ReadSubtree();
                                bFalg = xmlAppDir.ReadToFollowing(szReadValueOf);
                                if (bFalg)
                                {
                                    szReturnPath = szReturnPath + xmlAppDir.GetAttribute("value");
                                }
                                if (xmlAppDir != null)
                                    xmlAppDir.Close();
                                xmlAppDir = null;
                                break;
                            }
                            //-------------------------------------------------------------------

                        } //if (xmlLoc.GetAttribute("name").ToUpper() == szCurrentLocation.ToUpper())
                    }// if (xmlLoc.Read())
                } // while (xreader.Read())
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
            }
            finally
            {
                if (xmlAppDir != null)
                    xmlAppDir.Close();
                xmlAppDir = null;

                if (xreader != null)
                    xreader.Close();
                xreader = null;
                settings = null;
            }
            return szReturnPath;
        }

        #region ... For Getting DDSM Dir Path ...

        public string GetDDSMDirPath(string szLocation, string szDepartment, string szReadValueOf)
        {
            IsBackEnd = true;
            ErrorMsg = "";
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.IgnoreWhitespace = true;
            XmlReader xreader = null;// XmlReader.Create(szAppXmlPath, settings);
            XmlReader xmlLoc = null;
            XmlReader xmlAppDir = null;

            bool bAppDir = false, bFalg = false, bLocvar = false;
            string szDrive = "";
            string szCurrentLocation = "", szRootDir = "";
            string szReturnPath = "";
            string szSecFsPath = "";
            try
            {
                szCurrentLocation = szLocation;
                szRootDir = GetLocationVariable(szCurrentLocation, szDepartment, "RootDir");

                //... code Changed By manav on 29/12/2011 for DRD-917953 ...
                szDrive = Get_DriveName();

                if (!CheckLocationExist(szLocation))
                {
                    szLocation = szCurrentLocation;
                }

                xreader = XmlReader.Create(_szAppXmlPath, settings);
                while (xreader.Read())
                {
                    xreader.MoveToContent();
                    bLocvar = xreader.ReadToFollowing("Location");
                    xmlLoc = xreader.ReadSubtree();
                    if (xmlLoc.Read())
                    {
                        if (xmlLoc.GetAttribute("name").ToUpper() == szCurrentLocation.ToUpper())
                        {
                            _bIsfs = false;
                            _bIssecfs = false;
                            _bProLocations = false;

                            if (xmlLoc.GetAttribute("isfs").ToUpper() == "Y")
                                _bIsfs = true;

                            if (xmlLoc.GetAttribute("issecfs").ToUpper() == "Y")
                                _bIssecfs = true;

                            if (xmlLoc.GetAttribute("processotherloc").ToUpper() == "Y")
                                _bProLocations = true;

                            //--------- Check Process another locations for backend ----------------------

                            if (IsBackEnd)
                            {
                                if (_bProLocations && szCurrentLocation.ToUpper() != szLocation.ToUpper())
                                {
                                    bLocvar = xmlLoc.ReadToFollowing("ProcessLocations");
                                    if (bLocvar)
                                    {
                                        szReturnPath = ReadProcessLocation(xmlLoc, szLocation);
                                        if (xmlLoc != null)
                                            xmlLoc.Close();
                                        xmlLoc = null;
                                    }
                                }
                            }
                            //-----------------------------------------------------------------

                            if (szReturnPath == "")
                            {
                                //------------- Get Physical path --------------------------------
                                if (_bIsfs && _bIssecfs == false)
                                {
                                    szReturnPath = szDrive + szRootDir;
                                    if (xmlLoc != null)
                                        xmlLoc.Close();
                                    xmlLoc = null;
                                }//if(bIsfs == true && bIssecfs == false)
                                // if department has sec fs then return that path else return physical path 
                                else if (_bIsfs && _bIssecfs)
                                {
                                    bLocvar = xmlLoc.ReadToFollowing("SecondaryFs");
                                    if (bLocvar)
                                    {
                                        szSecFsPath = ReadSecondaryFs(xmlLoc, szDepartment);
                                    }

                                    if (szSecFsPath != "")
                                        szReturnPath = szSecFsPath;
                                    else
                                        szReturnPath = szDrive + szRootDir;
                                    if (xmlLoc != null)
                                        xmlLoc.Close();
                                    xmlLoc = null;
                                } //if (bIsfs == true && bIssecfs == true)
                                // if department has no sec fs and current location is not file server then return default fs path 
                                else if (_bIsfs == false && _bIssecfs == false)
                                {
                                    bLocvar = xmlLoc.ReadToFollowing("DefaultFs");
                                    if (bLocvar)
                                    {
                                        szReturnPath = "\\\\" + xmlLoc.GetAttribute("ip") + "\\" + xmlLoc.GetAttribute("sharedfolder");
                                    }
                                    if (xmlLoc != null)
                                        xmlLoc.Close();
                                    xmlLoc = null;

                                }//if (bIsfs == false && bIssecfs == false)
                                else if (_bIsfs == false && _bIssecfs)
                                {
                                    bLocvar = xmlLoc.ReadToFollowing("DefaultFs");
                                    if (bLocvar)
                                    {
                                        szReturnPath = "\\\\" + xmlLoc.GetAttribute("ip") + "\\" + xmlLoc.GetAttribute("sharedfolder");
                                    }

                                    bLocvar = xmlLoc.ReadToFollowing("SecondaryFs");
                                    if (bLocvar)
                                    {
                                        szSecFsPath = ReadSecondaryFs(xmlLoc, szDepartment);
                                    }

                                    if (szSecFsPath != "")
                                    {
                                        szReturnPath = szSecFsPath;
                                    }
                                    if (xmlLoc != null)
                                        xmlLoc.Close();
                                    xmlLoc = null;
                                }// if (bIsfs == false && bIssecfs == true)
                            }

                            //----- Read Application Directory values from xml -----------------
                            bAppDir = xreader.ReadToFollowing("AppDirectory");
                            if (bAppDir)
                            {
                                xmlAppDir = xreader.ReadSubtree();
                                bFalg = xmlAppDir.ReadToFollowing(szReadValueOf);
                                if (bFalg)
                                {
                                    szReturnPath = szReturnPath + xmlAppDir.GetAttribute("value");
                                }
                                if (xmlAppDir != null)
                                    xmlAppDir.Close();
                                xmlAppDir = null;
                                break;
                            }
                            //-------------------------------------------------------------------

                        } //if (xmlLoc.GetAttribute("name").ToUpper() == szCurrentLocation.ToUpper())
                    }// if (xmlLoc.Read())
                } // while (xreader.Read())
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
            }
            finally
            {
                if (xmlAppDir != null)
                    xmlAppDir.Close();
                xmlAppDir = null;

                if (xreader != null)
                    xreader.Close();
                xreader = null;
                settings = null;
            }
            return szReturnPath;
        }

        #endregion

        private string ReadProcessLocation(XmlReader xRead, string szPocessLoacation)
        {
            string szReturn = "";
            try
            {
                while (xRead.Read())
                {
                    xRead.MoveToContent();

                    if (xRead.Name.ToUpper() == szPocessLoacation.ToUpper())
                    {
                        szReturn = xRead.GetAttribute("sharedfolder");
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
            }
            finally
            {
                if (xRead != null)
                    xRead.Close();
                xRead = null;
            }
            return szReturn;
        }

        private string ReadSecondaryFs(XmlReader xRead, string szDept)
        {
            string szSecFsPath = "";
            try
            {
                while (xRead.Read())
                {
                    xRead.MoveToContent();

                    if (xRead.Name.ToUpper() == szDept.ToUpper())
                    {
                        szSecFsPath = "\\\\" + xRead.GetAttribute("ip") + "\\" + xRead.GetAttribute("sharedfolder");
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
            }
            finally
            {
                if (xRead != null)
                    xRead.Close();
                xRead = null;
            }
            return szSecFsPath;
        }

        public bool IsDefaultFs(string szLocation)
        {
            ErrorMsg = "";
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.IgnoreWhitespace = true;
            XmlReader xreader = null;
            XmlReader xmlLoc = null;

            bool bLocvar = false;
            string szCurrentLocation = "";
            try
            {
                szCurrentLocation = GetCurrentLocation();
                xreader = XmlReader.Create(_szAppXmlPath, settings);
                while (xreader.Read())
                {
                    xreader.MoveToContent();
                    bLocvar = xreader.ReadToFollowing("Location");
                    xmlLoc = xreader.ReadSubtree();
                    if (xmlLoc.Read())
                    {
                        if (xmlLoc.GetAttribute("name").ToUpper() == szCurrentLocation.ToUpper())
                        {
                            _bIsfs = false;
                            if (xmlLoc.GetAttribute("isfs").ToUpper() == "N")
                                _bIsfs = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
            }
            finally
            {
                if (xmlLoc != null)
                    xmlLoc.Close();
                xmlLoc = null;

                if (xreader != null)
                    xreader.Close();
                xreader = null;
                settings = null;
            }
            return _bIsfs;
        }

        public string GetPhysicalPath(string szReadValueOf)
        {
            ErrorMsg = "";
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.IgnoreWhitespace = true;
            XmlReader xreader = null;// XmlReader.Create(szAppXmlPath, settings);
            XmlReader xmlLoc = null;
            XmlReader xmlAppDir = null;

            bool bAppDir = false, bFalg = false, bLocvar = false;
            string szDrive = "";
            string szCurrentLocation = "", szRootDir = "";
            string szReturnPath = "";
            try
            {
                szCurrentLocation = GetCurrentLocation();
                szRootDir = GetApplicationVariable("RootDir");

                //... code Changed By manav on 29/12/2011 for DRD-917953 ...
                szDrive = Get_DriveName();

                xreader = XmlReader.Create(_szAppXmlPath, settings);
                while (xreader.Read())
                {
                    xreader.MoveToContent();
                    bLocvar = xreader.ReadToFollowing("Location");
                    xmlLoc = xreader.ReadSubtree();
                    if (xmlLoc.Read())
                    {
                        if (xmlLoc.GetAttribute("name").ToUpper() == szCurrentLocation.ToUpper())
                        {
                            //----- Read Application Directory values from xml -----------------
                            bAppDir = xreader.ReadToFollowing("AppDirectory");
                            if (bAppDir)
                            {
                                xmlAppDir = xreader.ReadSubtree();
                                bFalg = xmlAppDir.ReadToFollowing(szReadValueOf);
                                if (bFalg)
                                {
                                    szReturnPath = szDrive + szRootDir + xmlAppDir.GetAttribute("value");
                                }
                                if (xmlAppDir != null)
                                    xmlAppDir.Close();
                                xmlAppDir = null;
                                break;
                            }
                            //-------------------------------------------------------------------
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
            }
            finally
            {
                if (xmlLoc != null)
                    xmlLoc.Close();
                xmlLoc = null;

                if (xmlAppDir != null)
                    xmlAppDir.Close();
                xmlAppDir = null;

                if (xreader != null)
                    xreader.Close();
                xreader = null;
                settings = null;
            }
            return szReturnPath;
        }

        public ArrayList GetAllLocation()
        {
            ErrorMsg = "";
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.IgnoreWhitespace = true;
            XmlReader xreader = null; // = XmlReader.Create(szAppXmlPath, settings);
            XmlReader xmlLoc = null;
            bool bLocvar = false;
            ArrayList arrReturn = new ArrayList();
            try
            {
                xreader = XmlReader.Create(_szAppXmlPath, settings);
                while (xreader.Read())
                {
                    xreader.MoveToContent();
                    if (xreader.NodeType == XmlNodeType.Element)
                    {
                        bLocvar = xreader.ReadToFollowing("Location");
                        if (bLocvar)
                        {
                            xmlLoc = xreader.ReadSubtree();
                            if (xmlLoc.Read())
                            {
                                if (xmlLoc.HasAttributes)
                                {
                                    arrReturn.Add(xmlLoc.GetAttribute("name"));
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
            }
            finally
            {
                if (xmlLoc != null)
                    xmlLoc.Close();
                xmlLoc = null;

                if (xreader != null)
                    xreader.Close();
                xreader = null;
                settings = null;
            }
            return arrReturn;
        }

        public string GetCurrentLocation(out bool bProcUnconfigLoc)
        {
            ErrorMsg = "";
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.IgnoreWhitespace = true;
            XmlReader xreader = null;
            bool bAppvar = false;
            string szReturn = "";
            bool bProcLoc = false;
            try
            {
                xreader = XmlReader.Create(_szAppXmlPath, settings);
                while (xreader.Read())
                {
                    xreader.MoveToContent();
                    if (xreader.NodeType == XmlNodeType.Element)
                    {
                        bAppvar = xreader.ReadToFollowing("CurrentLocation");
                        if (bAppvar)
                        {
                            szReturn = xreader.GetAttribute("name");

                            if (xreader.GetAttribute("procunconfigloc") == "Y")
                            {
                                bProcLoc = true;
                            }
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _bResult = false;
                ErrorMsg = ex.Message;
            }
            finally
            {
                if (xreader != null)
                    xreader.Close();
                xreader = null;
                settings = null;
            }
            bProcUnconfigLoc = bProcLoc;
            return szReturn;
        }

        private string Get_DriveName()
        {
            string szDriveName = "";
            try
            {
                if (IsBackEnd)
                {
                    //szDrive = Directory.GetDirectoryRoot(Environment.CurrentDirectory);
                    szDriveName = Directory.GetDirectoryRoot(AppDomain.CurrentDomain.BaseDirectory);
                }
                else
                {
                    szDriveName = Directory.GetDirectoryRoot(HttpContext.Current.Request.PhysicalApplicationPath);
                }
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
            }
            return szDriveName;
        }

        #endregion

        #region ..... Functions for IDisposable Interface .....

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

        ~clsReadAppXml()
        {
            Dispose(false);
        }

        private void Dispose(bool bDisposing)
        {
            if (!this.bDisposed)
            {
                if (bDisposing)
                {
                    if (_xReader != null)
                        _xReader.Close();

                    if (_dicAppVariables != null)
                        _dicAppVariables.Clear();

                    if (_lstMapExt != null)
                        _lstMapExt.Clear();
                    //component.Dispose();
                }
                _szError = null;
                _dicAppVariables = null;
                _lstMapExt = null;
                _xReader = null;
                _szAppXmlPath = null;
                ErrorMsg = null;
                //CloseHandle(handle);
                //handle = IntPtr.Zero;
                bDisposed = true;
            }
        }

        #endregion
    }
}