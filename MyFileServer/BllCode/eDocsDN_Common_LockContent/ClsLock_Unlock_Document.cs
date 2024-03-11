using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using eDocsDN_ReadAppXml;

namespace eDocsDN_Common_LockContent
{
    public class ClsLock_Unlock_Document
    {
        #region ... Variable Declaration ...
        string _szAppXmlPath = string.Empty;
        string _szDbName = string.Empty;
        clsReadAppXml _objINI = null;
        #endregion

        #region .... Property ....
        public string msgError { get; set; }
        #endregion

        #region ... Constructor ...

        public ClsLock_Unlock_Document(string AppXMLPath)
        {
            msgError = string.Empty;
            _szAppXmlPath = AppXMLPath;
        }

        public ClsLock_Unlock_Document(string szDbName, string AppXMLPath)
        {
            msgError = string.Empty;
            _szDbName = szDbName;
            _szAppXmlPath = AppXMLPath;
        }
        #endregion


        #region ....Public Functions.....
        public string LockUnlock(Stream strmDocument, string szFileType, string szLockUnlockFlag, ClsOpenXmlOperations.LockType eLockType, string szDocPwd, string szDocStatus, bool bPrint)
        {
            msgError = string.Empty;
            _objINI = new clsReadAppXml(_szAppXmlPath);
            try
            {
                if (_objINI.IsWordDocument.Contains(szFileType.ToUpper()))
                {
                    szDocPwd = "Espl123&;";
                    switch (szLockUnlockFlag.ToUpper())
                    {
                        case "L":
                            //Code Changed by KeDaR on 09/22/2014 for DRT-4314
                            switch (eLockType)
                            {
                                case ClsOpenXmlOperations.LockType.ReadOnly:
                                    break;
                                case ClsOpenXmlOperations.LockType.None:
                                    break;
                                case ClsOpenXmlOperations.LockType.Comments:
                                    if (!objOpenXmlDocsx.UnlockDocument(szFullFilePath, szDocPwd))
                                        throw new Exception(objOpenXmlDocsx.msgError);
                                    if (!objOpenXmlDocsx.LockDocument(szFullFilePath, ClsOpenXmlOperations.LockType.Comments, szDocPwd, szDocStatus, bPrint))
                                        throw new Exception(objOpenXmlDocsx.msgError);

                                    break;
                                case ClsOpenXmlOperations.LockType.TrackedChanges:
                                    break;
                                case ClsOpenXmlOperations.LockType.Forms:
                                    break;
                                default:
                                    break;
                            }


                            switch (eLockType)
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

        #endregion


    }
}
