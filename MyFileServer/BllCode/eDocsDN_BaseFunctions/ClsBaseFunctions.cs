using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
using System.Data;
using DDLLCS;
using System.Security.Cryptography;
using System.IO;

namespace eDocsDN_BaseFunctions
{
    public class ClsBaseFunctions
    {
        #region .... Variable Declaration ....

        ClsBuildQuery objDal = null;

        string _szDBName;
        string _szAppXmlPath;
        string _szQuery;
        string _szError;

        #endregion

        #region .... Property ....

        public string ErrorMsg
        {
            get { return _szError; }
            set { _szError = value; }
        }

        #endregion

        #region .... Constructor ....

        public ClsBaseFunctions(string szAppXmlPath)
        {
            ErrorMsg = "";
            _szDBName = "";
            this._szAppXmlPath = szAppXmlPath;
            objDal = new ClsBuildQuery(_szAppXmlPath);
            objDal.OpenConnection();
        }

        public ClsBaseFunctions(string szDBName, string szAppXmlPath)
        {
            ErrorMsg = "";
            this._szDBName = szDBName;
            this._szAppXmlPath = szAppXmlPath;
            objDal = new ClsBuildQuery(_szDBName, _szAppXmlPath);
            objDal.OpenConnection();
        }

        public ClsBaseFunctions(ClsBuildQuery objDal, string szAppXmlPath)
        {
            ErrorMsg = "";
            _szDBName = "";
            this._szAppXmlPath = szAppXmlPath;
            this.objDal = objDal;
        }

        #endregion

        #region .... Function ....

        public string GetCodeDescription(string szCode, string szType)
        {
            ErrorMsg = "";
            object objReturnVal = null;
            string szReturn = "";
            try
            {
                _szQuery = "Select csed from zespl_sedoc_nommoc where edoc='" + szCode + "' and epyt='" + szType + "'";
                objReturnVal = objDal.GetFirstColumnValue(_szQuery);
                if (objDal.msgError != "")
                    throw new Exception(objDal.msgError);

                if (objReturnVal != null)
                {
                    szReturn = objReturnVal.ToString();
                }
                objReturnVal = null;
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
            }
            finally
            {
                objReturnVal = null;
            }
            return szReturn;
        }

        public string ConvertDate(string szDate, string szDateFormat)
        {
            ErrorMsg = "";
            string szConvert = "MM/dd/yyyy";
            char[] arrFormat = szDateFormat.ToCharArray();
            char[] arrDate = szDate.ToCharArray();
            char[] arrConvert = szConvert.ToCharArray();
            string szConvertedDate = "";
            string szTempDate = "";
            ArrayList arrdate = new ArrayList();
            try
            {
                if (szDate == "")
                    return "";
                //throw new Exception("Date shoild be empty");

                for (int i = 0; i < arrConvert.Length; i++)
                {
                    szTempDate = "";
                    for (int j = 0; j < arrFormat.Length; j++)
                    {
                        if (arrConvert[i] == arrFormat[j] && arrConvert[i] != '/')
                        {
                            szTempDate = szTempDate + arrDate[j];
                        }
                    }
                    if (szTempDate != szConvertedDate)
                    {
                        szConvertedDate = szTempDate;
                        arrdate.Add(szTempDate);
                    }
                }

                szConvertedDate = "";
                for (int k = 0; k < arrdate.Count; k++)
                {
                    if (arrdate[k].ToString() != "")
                    {
                        if (szConvertedDate == "")
                        {
                            szConvertedDate = arrdate[k].ToString();
                        }
                        else
                        {
                            szConvertedDate = szConvertedDate + "/" + arrdate[k].ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
                szConvertedDate = ex.Message;
            }

            DateTime dt = new DateTime();
            try
            {
                dt = Convert.ToDateTime(szConvertedDate);
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
            }
            return szConvertedDate;
        }

        #endregion

        # region .... Encrypt Decrypt ....

        public string Encrypt(string szString)
        {
            string szEncString = "";
            for (int i = 0; i < szString.Length; i++)
            {
                szEncString = szEncString + (char)((int)szString[i] + 1);
            }
            szEncString = ReverseString(szEncString);
            return szEncString;
        }

        public string ReverseString(string szString)
        {
            string szReverseString = "";
            for (int i = szString.Length - 1; i >= 0; i--)
            {
                szReverseString = szReverseString + szString[i];
            }
            return szReverseString;
        }

        public string Decrypt(string szString)
        {
            szString = ReverseString(szString);
            string szTemp = "";
            for (int i = 0; i < szString.Length; i++)
            {
                szTemp = szTemp + (char)((int)szString[i] - 1);
            }
            return szTemp;
        }

        public string getCurrentDate()
        {
            return string.Format("{0:yyyy/MM/dd}", DateTime.Now);
        }

        public string getCurrentTime()
        {
            return string.Format("{0:hh:mm:ss}", DateTime.Now);
        }

        public string GetMd5Hash(string szFilePath)
        {
            // Create a new instance of the MD5CryptoServiceProvider object.
            MD5 md5Hasher = MD5.Create();
            StreamReader sr = new StreamReader(szFilePath);

            // compute the hash for given file.
            byte[] data = md5Hasher.ComputeHash(sr.BaseStream);

            sr.Close();
            sr.Dispose();
            sr = null;
            md5Hasher.Clear();
            md5Hasher = null;

            StringBuilder sBuilder = new StringBuilder();
            // Loop through each byte of the hashed data 
            // and format each one as a hexadecimal string.
            for (int i = 0; i < data.Length; i++)
            {
                sBuilder.Append(data[i].ToString("x2"));
            }
            return sBuilder.ToString();
        }

        # endregion
    }
}