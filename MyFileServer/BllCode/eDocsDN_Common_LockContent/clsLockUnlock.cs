using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.IO;
using DDLLCS;
using eDocsDN_ReadAppXml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using System.Security;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using Op = DocumentFormat.OpenXml.CustomProperties;
using V = DocumentFormat.OpenXml.Vml;
using System.Security.Cryptography;
using System.IO;
using DocumentFormat.OpenXml.CustomProperties;

using Microsoft.Win32;

namespace eDocsDN_Common_LockContent
{
    public class clsLockUnlock : IDisposable
    {
        #region ...Class Variables...

        private string szError = "";
        private string szDatabaseName;
        private string szAppXMLPath;

        clsWord objWord = new clsWord();
        clsExcel objExcel = new clsExcel();
        ClsOpenXmlOperations objOpenXmlDocsx = new ClsOpenXmlOperations();
        clsReadAppXml oReadXml;
        DDLLCS.ClsBuildQuery objDAL;

        private string szDBName;

        #endregion

        #region ...Error Message...

        public string ErrorMsg
        {
            get
            {
                return szError;
            }
            set
            {
                szError = value;
            }
        }

        #endregion

        #region ... Constructor ...

        //public clsLockUnlock()
        //{
        //    //this.szDatabaseName = szDatabaseName;
        //    //this.ApplicationXmlFilePath = ApplicationXmlFilePath;
        //}//public clsLockUnlock()

        public clsLockUnlock(string AppXMLPath)
        {
            this.szAppXMLPath = AppXMLPath;
            szError = "";
        }

        public clsLockUnlock(string szDbName, string AppXMLPath)
        {
            this.szDBName = szDbName;
            this.szAppXMLPath = AppXMLPath;
            szError = "";
        }
        #endregion

        #region ....Public Functions.....
        /// <summary>
        /// DRT:4314
        /// Code Chaged by KeDaR 
        /// Lock Unlock for Docx using openXml SDK
        /// Fist Unlock the docuemnt for next lock 
        /// </summary>
        /// <param name="szFullFilePath"></param>
        /// <param name="szFileType"></param>
        /// <param name="szLockUnlockFlag"></param>
        /// <param name="szLockType"></param>
        /// <param name="szDocPwd"></param>
        /// <param name="szDocStatus"></param>
        /// <param name="bPrint"></param>
        /// <returns></returns>
        public string LockUnlock(string szFullFilePath, string szFileType, string szLockUnlockFlag, string szLockType, string szDocPwd, string szDocStatus, bool bPrint)
        {
            ErrorMsg = "";
            oReadXml = new clsReadAppXml(szAppXMLPath);
            try
            {
                if (szLockType == "")
                    szLockType = "COMM";

                if (oReadXml.IsWordDocument.Contains(szFileType.ToUpper()))
                {
                    if (Path.GetExtension(szFullFilePath).Equals(".docx", StringComparison.InvariantCultureIgnoreCase))
                    {
                        szDocPwd = "Espl123&;";
                        switch (szLockUnlockFlag.ToUpper())
                        {
                            case "L":
                                //Code Changed by KeDaR on 09/22/2014 for DRT-4314
                                switch (szLockType)
                                {
                                    case "COMM":
                                        if (!objOpenXmlDocsx.UnlockDocument(szFullFilePath, szDocPwd))
                                            throw new Exception(objOpenXmlDocsx.msgError);
                                        if (!objOpenXmlDocsx.LockDocument(szFullFilePath, ClsOpenXmlOperations.LockType.Comments, szDocPwd, szDocStatus, bPrint))
                                            throw new Exception(objOpenXmlDocsx.msgError);
                                        break;
                                    case "FORM":
                                        if (!objOpenXmlDocsx.UnlockDocument(szFullFilePath, szDocPwd))
                                            throw new Exception(objOpenXmlDocsx.msgError);
                                        if (!objOpenXmlDocsx.LockDocument(szFullFilePath, ClsOpenXmlOperations.LockType.Forms, szDocPwd, szDocStatus, bPrint))
                                            throw new Exception(objOpenXmlDocsx.msgError);
                                        break;
                                    default:
                                        if (!objOpenXmlDocsx.UnlockDocument(szFullFilePath, szDocPwd))
                                            throw new Exception(objOpenXmlDocsx.msgError);
                                        if (!objOpenXmlDocsx.LockDocument(szFullFilePath, ClsOpenXmlOperations.LockType.ReadOnly, szDocPwd, szDocStatus, bPrint))
                                            throw new Exception(objOpenXmlDocsx.msgError);
                                        break;
                                }
                                break;
                            case "U":
                                if (!objOpenXmlDocsx.UnlockDocument(szFullFilePath, szDocPwd))
                                    throw new Exception(objOpenXmlDocsx.msgError);
                                break;
                        }
                    }
                    else
                    {
                        if (szLockUnlockFlag.ToUpper().Equals("L", StringComparison.CurrentCultureIgnoreCase))
                        {
                            szError = objWord.UnLockContentDoc(szFullFilePath, szDocPwd);
                            if (szError.Trim() != "")
                                throw new Exception(szError);

                            szError = objWord.LockContentDoc(szFullFilePath, szLockType, szDocPwd, szDocStatus, bPrint);
                            if (szError.Trim() != "")
                                throw new Exception(szError);
                        }
                        else if (szLockUnlockFlag.ToUpper().Equals("U", StringComparison.CurrentCultureIgnoreCase))
                        {
                            szError = objWord.UnLockContentDoc(szFullFilePath, szDocPwd);
                            if (szError.Trim() != "")
                                throw new Exception(szError);
                        }
                    }
                }
                else if (oReadXml.IsExcelDocument.Contains(szFileType.ToUpper()))
                {
                    if (szLockUnlockFlag.ToUpper().Equals("L", StringComparison.CurrentCultureIgnoreCase))
                    {
                        szError = objExcel.UnLockContentXls(szFullFilePath, szDocPwd);
                        if (szError.Trim() != "")
                            throw new Exception(szError);

                        szError = objExcel.LockExcelContent(szFullFilePath, szLockType, szDocPwd, szDocStatus, bPrint);
                        if (szError.Trim() != "")
                            throw new Exception(szError);
                    }
                    else if (szLockUnlockFlag.ToUpper().Equals("U", StringComparison.CurrentCultureIgnoreCase))
                    {
                        szError = objExcel.UnLockContentXls(szFullFilePath, szDocPwd);
                        if (szError.Trim() != "")
                            throw new Exception(szError);
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
            }
            finally
            {
                if (oReadXml != null)
                    oReadXml.Dispose();
                oReadXml = null;
            }
            return (ErrorMsg);
        }

        public void ConvertDOCToHTML(string szSourceFilePath, string szDestFilePath, bool bQuitWord)
        {
            szError = objWord.ConvertWordToHtml(szSourceFilePath, szDestFilePath);
            if (!szError.Equals(""))
                ErrorMsg = szError;
        }

        //public void ConvertWordToPdf(string szWordSourcePath, string szPDFDestPath)
        //{
        //    ErrorMsg = objWord.ConvertWordToPdf(szWordSourcePath, szPDFDestPath);
        //}

        public void ConvertExcelToHTML(string szSourceExcelFilePath, string szTargetHtmlFilePath)
        {
            szError = objExcel.Convert_Excel_To_Html(szSourceExcelFilePath, szTargetHtmlFilePath);
            if (!szError.Equals(""))
                ErrorMsg = szError;
        }

        public void ConvertExcelToPDF(string szExcelPath, string szPDFPath)
        {
            objExcel.Convert_Excel_To_Pdf(szExcelPath, szPDFPath);
            if (!szError.Equals(""))
                ErrorMsg = szError;
        }

        public string Lock_Document_To_Author(string szComp, string szLoc, string szDept, string szDocType, string szFileName, string szFileExt, ClsBuildQuery objDAL)
        {

            szError = "";

            try
            {

                string szSQL;
                szSQL = "select * from zespl_setalpmet_cod  ";
                szSQL += "where upper(ynapmoc)='" + szComp.ToUpper() + "'";
                szSQL += "And upper(noitacol)='" + szLoc.ToUpper() + "'";
                szSQL += "And upper(tnemtraped)='" + szDept.ToUpper() + "'";
                szSQL += "And upper(epyt_cod)='" + szDocType.ToUpper() + "'";

                IDataReader drAdmin = objDAL.DecideDatabaseQDR(szSQL);
                if (objDAL.msgError != null)
                {
                    ErrorMsg = objDAL.msgError;
                    return (objDAL.msgError.ToString());
                }
                else
                {
                    if (drAdmin != null)
                    {
                        if (drAdmin.Read())
                        {
                            //if ((drAdmin["etalpmet_cod"].ToString().Equals("1")) || (drAdmin["tide_tamrof"].ToString().Equals("0")))
                            if (Convert.ToBoolean(drAdmin["etalpmet_cod"]) == true)
                            {
                                if (Convert.ToBoolean(drAdmin["tide_tamrof"]) == true)
                                {
                                    //no need to lock
                                }
                                else
                                {
                                    LockUnlock(szFileName, szFileExt, "L", "FORM", "", "", false);
                                    if (ErrorMsg != "")
                                    {
                                        return (ErrorMsg.ToString());
                                    }
                                }
                            }
                            //else if ((drAdmin["etalpmet_cod"].ToString().Equals("0")) || (drAdmin["tide_tamrof"].ToString().Equals("0")))
                            else if (Convert.ToBoolean(drAdmin["etalpmet_cod"]) == false)
                            {
                                if (Convert.ToBoolean(drAdmin["tide_tamrof"]) == true)
                                {
                                    //no need to lock
                                }
                                else
                                {
                                    LockUnlock(szFileName, szFileExt, "L", "FORM", "", "", false);
                                    if (ErrorMsg != "")
                                    {
                                        return (ErrorMsg.ToString());
                                    }
                                }
                            }

                        }//if(drAdmin.Read())
                        drAdmin.Close();
                        drAdmin.Dispose();
                        drAdmin = null;
                    }//if (drAdmin!=null)

                }
            }//try
            catch (Exception e)
            {
                szError = e.Message;
            }//catch

            return (szError);

        }//public void Lock_Document_To_Author(string szComp,string szLoc,string szDept,string szDocType,string szFileName,string szFileExt)

        public bool IsFileLocked(string szFilePath)
        {
            try
            {
                if (File.Exists(szFilePath))
                {

                    FileStream fs = new FileStream(szFilePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                    fs.Close();
                    return false;
                }
                else
                {
                    return false;
                }

            }
            catch (Exception e)
            {
                szError = e.Message;
                return true;
            }
        }//Public bool IsFileLocked()

        public bool Convert_To_Htm_And_Add_Security_Code(string szFullFilePath, string szFileType)
        {
            bool bREsult = true;
            oReadXml = new clsReadAppXml(szAppXMLPath);

            try
            {
                if (oReadXml.IsWordDocument.Contains(szFileType.ToUpper()))
                {
                    ConvertDOCToHTML(szFullFilePath, szFullFilePath.Replace("." + szFileType, ".htm"), true);
                }
                else if (oReadXml.IsExcelDocument.Contains(szFileType.ToUpper()))
                {
                    ConvertDOCToHTML(szFullFilePath, szFullFilePath.Replace("." + szFileType, ".htm"), true);
                }
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
                bREsult = false;
            }
            return bREsult;
        }

        public void ConvertWordFileToPdf(string szSourceFile, string szDestFilePath, bool bLockFlag)
        {
            bool isAdobeAcrobatInstalled;
            try
            {
                //.... Check adobe available or not ....
                isAdobeAcrobatInstalled = IsAdobeInstalled();
                if (!isAdobeAcrobatInstalled)
                    throw new Exception("Acrobat Distiller not installed.");

                objWord.ConvertWordFileToPdf(szSourceFile, szDestFilePath, bLockFlag);
                if (objWord.msgError != "")
                    throw new Exception(objWord.msgError);
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
            }
        }

        private bool IsAdobeInstalled()
        {
            bool isAdobeAcrobatInstalled = false;
            string szAdobeAcrobatVersion = "";
            //RegistryKey objRegAdobeKey = null;
            string RegPath = @"SOFTWARE\Adobe";

            try
            {
                using (RegistryKey RegKey = Registry.LocalMachine.OpenSubKey(RegPath))
                {
                    if (RegKey != null)
                    {
                        foreach (string SW_Name in RegKey.GetSubKeyNames())
                        {
                            if (SW_Name == "Acrobat Distiller" || SW_Name == "Adobe Acrobat")
                            {
                                isAdobeAcrobatInstalled = true;
                                //objRegAdobeKey = Registry.LocalMachine.OpenSubKey("Software").OpenSubKey("Adobe");
                                //if (null == objRegAdobeKey)
                                //{
                                //    var policies = Registry.LocalMachine.OpenSubKey("Software").OpenSubKey("Policies");
                                //    objRegAdobeKey = policies.OpenSubKey("Adobe");
                                //}
                                //break;                            
                            }
                        }
                    }
                }

                if (!isAdobeAcrobatInstalled)
                {
                    RegPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
                    using (RegistryKey RegKey = Registry.LocalMachine.OpenSubKey(RegPath))
                    {
                        if (RegKey != null)
                        {
                            foreach (string SW_Name in RegKey.GetSubKeyNames())
                            {
                                if (SW_Name == "Acrobat Distiller" || SW_Name == "Adobe Acrobat")
                                {
                                    isAdobeAcrobatInstalled = true;
                                    //objRegAdobeKey = Registry.LocalMachine.OpenSubKey("Software").OpenSubKey("Adobe");
                                    //if (null == objRegAdobeKey)
                                    //{
                                    //    var policies = Registry.LocalMachine.OpenSubKey("Software").OpenSubKey("Policies");
                                    //    objRegAdobeKey = policies.OpenSubKey("Adobe");
                                    //}
                                    //break;                            
                                }
                            }
                        }
                    }
                }

                //.... Code lines added for DRT- on 30/7/2014 by Harshad ....
                if (!isAdobeAcrobatInstalled)
                {
                    RegPath = @"SOFTWARE\Adobe";
                    using (RegistryKey RegKey = Registry.CurrentUser.OpenSubKey(RegPath))
                    {
                        if (RegKey != null)
                        {
                            foreach (string SW_Name in RegKey.GetSubKeyNames())
                            {
                                if (SW_Name == "Acrobat Distiller" || SW_Name == "Adobe Acrobat")
                                {
                                    isAdobeAcrobatInstalled = true;
                                    //objRegAdobeKey = Registry.LocalMachine.OpenSubKey("Software").OpenSubKey("Adobe");
                                    //if (null == objRegAdobeKey)
                                    //{
                                    //    var policies = Registry.LocalMachine.OpenSubKey("Software").OpenSubKey("Policies");
                                    //    objRegAdobeKey = policies.OpenSubKey("Adobe");
                                    //}
                                    //break;                            
                                }
                            }
                        }
                    }
                }
                //....End of Code lines added for DRT- on 30/7/2014 by Harshad ....
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
                //throw ex;
            }
            finally
            {
                szAdobeAcrobatVersion = RegPath = null;
                //objRegAdobeKey = null;
            }
            return isAdobeAcrobatInstalled;
        }

        #endregion


        #region .... disposable ...
        public new void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        protected new virtual void Dispose(bool disposing)
        {
            if (disposing)
            {


            }
            else
            {

            }
        }

        ~clsLockUnlock()
        {
            Dispose(false);
        }


        #endregion
    }
}