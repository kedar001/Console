/* Last code Changed by: manavya
 * Last Code changed on: 13-12-2012
 * Last Code changed for: DRT-4136
 */

using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
//using System.ComponentModel;

namespace GetIniInfo
{
    public class ClsGetIniInfo : IDisposable
    {
        #region .... Variable Declaration ....

        XmlTextReader _xReader;
        string _szAppXmlPath;
        string _szError;

        #endregion

        #region .... Constructor ....

        public ClsGetIniInfo(string szAppXmlPath)
        {
            //IntPtr handle
            //this.handle = handle;
            msgError = "";
            _szAppXmlPath = szAppXmlPath;
        }

        #endregion

        #region .... Property ....

        public string msgError
        {
            get { return _szError; }
            set { _szError = value; }
        }

        #endregion

        #region .... Function Definition ....

        public string GetIniInfo(string szType, string szKey)
        {
            string szContents = "";
            szType = szType.Trim();
            szKey = szKey.Trim();
            try
            {
                bool bKeyFound = false;
                bool bTagFound = false;
                bool bValueFound = false;
                bool bKeyTag = false;
                string szTagtype = "";
                _xReader = new XmlTextReader(_szAppXmlPath);

                while (_xReader.Read())
                {
                    if (bValueFound)
                        break;

                    _xReader.MoveToContent();
                    switch (_xReader.NodeType)
                    {
                        case XmlNodeType.Element:
                            {
                                if (_xReader.HasAttributes)
                                {
                                    szTagtype = _xReader.GetAttribute(0).Trim();

                                    if (szType != "")
                                        if (szTagtype.Equals(szType, StringComparison.CurrentCultureIgnoreCase))
                                        {
                                            bTagFound = true;
                                        }
                                }

                                if (_xReader.Name.Equals("key", StringComparison.CurrentCultureIgnoreCase))
                                    bKeyTag = true;
                            }
                            break;
                        case XmlNodeType.Text:
                            {
                                if (bTagFound && bKeyFound)
                                {
                                    szContents = _xReader.Value;
                                    bValueFound = true;
                                    bTagFound = false;
                                    bKeyFound = false;
                                }

                                if (bTagFound && !bKeyFound && bKeyTag)
                                {
                                    if (szKey != "")
                                        if (_xReader.Value.Trim().Equals(szKey, StringComparison.InvariantCultureIgnoreCase))
                                            bKeyFound = true;
                                }
                            }
                            break;
                        case XmlNodeType.EndElement:
                            {
                                if (_xReader.Name.Equals("key", StringComparison.CurrentCultureIgnoreCase))
                                    bKeyTag = false;

                                if (bTagFound && bKeyFound)
                                {
                                    if (_xReader.Name.Equals("value", StringComparison.CurrentCultureIgnoreCase))
                                        bValueFound = true;
                                }
                            }
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            finally
            {
                if (_xReader != null)
                    _xReader.Close();
                _xReader = null;
            }
            return szContents;
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

        ~ClsGetIniInfo()
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
                    //component.Dispose();
                }

                _xReader = null;
                _szAppXmlPath = null;
                msgError = null;
                //CloseHandle(handle);
                //handle = IntPtr.Zero;
                bDisposed = true;
            }
        }

        #endregion
    }
}