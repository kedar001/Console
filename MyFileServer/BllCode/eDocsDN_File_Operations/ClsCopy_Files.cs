﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using eDocsDN_ReadAppXml;
using System.IO;
using DDLLCS;
using System.Collections;
using eDocsDN_Get_Directory_Info;
using eDocsDN_Save_File_in_Database;
using System.Data;
using System.Data.OleDb;
using eDocsDN_File_Server_Operations;

namespace eDocsDN_File_Operations
{
    public class ClsCopy_Files : ClsCheck_Configuration_For_File_Storage, IDisposable
    {
        #region .... Variable Declaration ...
        ClsDocumentDirPath objDir = null;
        Directory_Attributes objSource_Dir = null;
        Directory_Attributes objDestination_Dir = null;
        ClsSave_File_in_Database _objSave_as_Blob = null;
        File_Data _oDestination = null;
        ClsFile_Server_Operations objFile_server = null;

        string _szAppXmlPath = string.Empty;
        string _szDBName = string.Empty;
        string _szLogFileName = string.Empty;
        bool _bisDebug = false;

        #endregion

        #region ..... Property ...
        public string Location { get; set; }
        public bool bIsBackEndOperation { get; set; }
        public string Department { get; set; }
        public string Source_Path { get; set; }
        public string Destination_Path { get; set; }


        #endregion

        #region .... Constructor ....
        public ClsCopy_Files(string szAppXmlPath, string szDBName, string szLocation)
            : base(szAppXmlPath, szDBName, szLocation)
        {
            msgError = string.Empty;
            _szAppXmlPath = szAppXmlPath;
            Location = szLocation;
            _szDBName = szDBName;
            File_Storage_in_Database();
            _szLogFileName = AppDomain.CurrentDomain.BaseDirectory + @"\Log\DetailLog.txt";
            if (!Directory.Exists(AppDomain.CurrentDomain.BaseDirectory + @"\Log\"))
                Directory.CreateDirectory(AppDomain.CurrentDomain.BaseDirectory + @"\Log\");

        }
        #endregion

        #region .... ENUM ......
        public enum File_Type
        {
            Template = 0,
            Document
        }
        #endregion

        #region .... Public Functions ....

        public List<File_Data> Copy_File(string szSource_Dir, string szDestination_Dir, List<File_Data> lstFiles)
        {
            _oDestination = new File_Data();
            List<File_Data> lstUpdatedFile_Data = new List<File_Data>();
            File_Data oFileData = null;

            objDir = new ClsDocumentDirPath(_szAppXmlPath, _szDBName, Location, Department);
            _objSave_as_Blob = new ClsSave_File_in_Database(_szAppXmlPath, _szDBName, Location);
            try
            {
                objDestination_Dir = objDir.GetDirPath(szDestination_Dir, bStoreFilesinBlob, bPhysicalFileStorage, bIsEncryptionEnabled);
                objSource_Dir = objDir.GetDirPath(szSource_Dir, bStoreFilesinBlob, bPhysicalFileStorage, bIsEncryptionEnabled);


                #region .... Copy File on File System .....
                foreach (File_Data oSource in lstFiles)
                {
                    oSource.Destination_File_Name = string.IsNullOrEmpty(oSource.Destination_File_Name) ? oSource.File_Name : oSource.Destination_File_Name;
                    oFileData = Copy_File(objSource_Dir, oSource, objDestination_Dir, oSource);
                    lstUpdatedFile_Data.Add(oFileData);
                    oFileData.Dispose();
                    oFileData = null;
                }

                foreach (File_Data item in lstFiles)
                {
                    lstFiles.Remove(item);
                    item.Dispose();
                }


                #endregion


            }
            catch (Exception ex)
            {
                string message = "Exception type " + ex.GetType() + Environment.NewLine + "Exception message: " + ex.Message + Environment.NewLine + "Stack trace: " + ex.StackTrace + Environment.NewLine;
                if (ex.InnerException != null)
                {
                    message += "---BEGIN InnerException--- " + Environment.NewLine +
                               "Exception type " + ex.InnerException.GetType() + Environment.NewLine +
                               "Exception message: " + ex.InnerException.Message + Environment.NewLine +
                               "Stack trace: " + ex.InnerException.StackTrace + Environment.NewLine +
                               "---END Inner Exception";
                }
                msgError = message;
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + msgError + Environment.NewLine);
            }
            finally
            {
            }

            return lstUpdatedFile_Data;
        }

        public File_Data Copy_File(string szSource_Dir, string szDestination_Dir, File_Data oFiles)
        {
            objDir = new ClsDocumentDirPath(_szAppXmlPath, _szDBName, Location, Department);
            _objSave_as_Blob = new ClsSave_File_in_Database(_szAppXmlPath, _szDBName, Location);
            _oDestination = new File_Data();
            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " ======START(Copy_File(3))=====" + Environment.NewLine);
            try
            {
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Copy_File commented " + Environment.NewLine);
                objDestination_Dir = objDir.GetDirPath(szDestination_Dir, bStoreFilesinBlob, bPhysicalFileStorage, bIsEncryptionEnabled);
                objSource_Dir = objDir.GetDirPath(szSource_Dir, bStoreFilesinBlob, bPhysicalFileStorage, bIsEncryptionEnabled);
                #region .... Copy File on File System .....
                oFiles.Destination_File_Name = string.IsNullOrEmpty(oFiles.Destination_File_Name) ? oFiles.File_Name : oFiles.Destination_File_Name;

                if (bPhysicalFileStorage && oFiles.Destination_Directory.Equals("DV"))
                {
                    objDestination_Dir.Directory_Path = objDestination_Dir.Directory_Path + "\\" + Convert.ToString(oFiles.Draft_Version) + "\\";
                }
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : oFiles.Destination_File_Name : " + Convert.ToString(oFiles.Destination_File_Name) + Environment.NewLine);
                oFiles = Copy_File(objSource_Dir, oFiles, objDestination_Dir, oFiles);
                _oDestination = oFiles;

                #endregion

                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Copy_File end " + Environment.NewLine);
            }
            catch (Exception ex)
            {
                string message = "Exception type " + ex.GetType() + Environment.NewLine + "Exception message: " + ex.Message + Environment.NewLine + "Stack trace: " + ex.StackTrace + Environment.NewLine;
                if (ex.InnerException != null)
                {
                    message += "---BEGIN InnerException--- " + Environment.NewLine +
                               "Exception type " + ex.InnerException.GetType() + Environment.NewLine +
                               "Exception message: " + ex.InnerException.Message + Environment.NewLine +
                               "Stack trace: " + ex.InnerException.StackTrace + Environment.NewLine +
                               "---END Inner Exception";
                }
                msgError = message;
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + message + Environment.NewLine);
            }
            finally
            {
            }
            return oFiles;
        }

        public File_Data Copy_File(File_Data oFiles)
        {

            _oDestination = new File_Data();
            objDir = new ClsDocumentDirPath(_szAppXmlPath, _szDBName, Location, Department);

            try
            {
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " ======START(Copy_File(1))=====" + Environment.NewLine);
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Copy File Called : " + msgError + Environment.NewLine);
                if (!string.IsNullOrEmpty(oFiles.Referance_Source_Location))
                {
                    if (!oFiles.Referance_Source_Location.Equals(Location))
                    {
                        if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Referance_Source_Location : " + oFiles.Referance_Source_Location + Environment.NewLine);
                        if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Location : " + Location + Environment.NewLine);
                        if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : _szAppXmlPath : " + _szAppXmlPath + Environment.NewLine);
                        eDocsDN_File_Server_Operations.File_Server.File_Data objFileData = new eDocsDN_File_Server_Operations.File_Server.File_Data();
                        objFile_server = new ClsFile_Server_Operations(_szAppXmlPath, oFiles.Referance_Source_Location);
                        objFileData.Source_Directory = oFiles.Referance_Source_Dir;
                        objFileData.SurrKey = oFiles.Referance_Source_SurrKey;
                        objFileData.Serial_Number = oFiles.Referance_Source_SerialNumber;
                        objFileData.File_Name = oFiles.Referance_Source_FileName;
                        oFiles.Data = objFile_server.Get_File_Information(objFileData).Data;
                        if (objFile_server.msgError != "")
                        {
                            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + objFile_server.msgError + Environment.NewLine);
                            throw new Exception(objFile_server.msgError);
                        }
                        objFileData = null;
                    }
                    else
                    {
                        Directory_Attributes objTempSource_Dir = null;
                        objTempSource_Dir = objDir.GetDirPath(oFiles.Referance_Source_Dir, bStoreFilesinBlob, bPhysicalFileStorage, bIsEncryptionEnabled);
                        File_Data oGet_source = new File_Data();
                        oGet_source.Source_Directory = oFiles.Referance_Source_Dir;
                        oGet_source.SurrKey = oFiles.Referance_Source_SurrKey;
                        oGet_source.Serial_Number = oFiles.Referance_Source_SerialNumber;
                        oGet_source.File_Name = oFiles.Referance_Source_FileName;
                        oFiles.Data = Get_File_Infomation(objTempSource_Dir, oGet_source).Data;
                        objTempSource_Dir = null;
                    }
                }



                if (!string.IsNullOrEmpty(oFiles.Destination_Directory))
                    objDestination_Dir = objDir.GetDirPath(oFiles.Destination_Directory, bStoreFilesinBlob, bPhysicalFileStorage, bIsEncryptionEnabled);
                if (!string.IsNullOrEmpty(oFiles.Source_Directory))
                    objSource_Dir = objDir.GetDirPath(oFiles.Source_Directory, bStoreFilesinBlob, bPhysicalFileStorage, bIsEncryptionEnabled);


                #region .... Copy File on File System .....
                if (bPhysicalFileStorage && oFiles.Destination_Directory.Equals("DV"))
                {
                    objDestination_Dir.Directory_Path = objDestination_Dir.Directory_Path + "\\" + Convert.ToString(oFiles.Draft_Version) + "\\";
                }
                oFiles.Destination_File_Name = string.IsNullOrEmpty(oFiles.Destination_File_Name) ? oFiles.File_Name : oFiles.Destination_File_Name;

                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Before Copy_File : " + msgError + Environment.NewLine);
                oFiles = Copy_File(objSource_Dir, oFiles, objDestination_Dir, _oDestination);
                _oDestination = oFiles;

                #endregion
            }
            catch (Exception ex)
            {
                string message = "Exception type " + ex.GetType() + Environment.NewLine + "Exception message: " + ex.Message + Environment.NewLine + "Stack trace: " + ex.StackTrace + Environment.NewLine;
                if (ex.InnerException != null)
                {
                    message += "---BEGIN InnerException--- " + Environment.NewLine +
                               "Exception type " + ex.InnerException.GetType() + Environment.NewLine +
                               "Exception message: " + ex.InnerException.Message + Environment.NewLine +
                               "Stack trace: " + ex.InnerException.StackTrace + Environment.NewLine +
                               "---END Inner Exception";
                }
                msgError = message;
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + message + Environment.NewLine);
            }
            finally
            {
                ////if (oFiles != null)
                ////    oFiles.Dispose();
                ////oFiles = null;
                //if (objDir != null)
                //    objDir.Dispose();
                //objDir = null;
                //if (objSource_Dir != null)
                //    objSource_Dir.Dispose();
                //objSource_Dir = null;
                //if (objDestination_Dir != null)
                //    objDestination_Dir.Dispose();
                //objDestination_Dir = null;
                //GC.Collect();
            }

            return oFiles;
        }

        public bool Check_File_Exist_In_Source(File_Data oFile)
        {
            bool bIsFile_Exist = false;
            try
            {
                objDir = new ClsDocumentDirPath(_szAppXmlPath, _szDBName, Location, Department);
                objSource_Dir = objDir.GetDirPath(oFile.Source_Directory, bStoreFilesinBlob, bPhysicalFileStorage, bIsEncryptionEnabled);
                if (bPhysicalFileStorage && oFile.Source_Directory.Equals("DV"))
                {
                    objSource_Dir.Directory_Path = objSource_Dir.Directory_Path + "\\" + Convert.ToString(oFile.Draft_Version) + "\\";
                }
                bIsFile_Exist = Check_File_Exist(objSource_Dir, oFile);
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + msgError + Environment.NewLine);
                bIsFile_Exist = false;
            }
            finally
            {
                if (oFile != null)
                    oFile.Dispose();
                oFile = null;
                if (objDir != null)
                    objDir.Dispose();
                objDir = null;
                if (objSource_Dir != null)
                    objSource_Dir.Dispose();
                objSource_Dir = null;
                if (objDestination_Dir != null)
                    objDestination_Dir.Dispose();
                objDestination_Dir = null;
            }
            return bIsFile_Exist;
        }

        public bool Check_File_Exist_In_Destination(File_Data oFile)
        {
            bool bIsFile_Exist = false;
            try
            {
                objDir = new ClsDocumentDirPath(_szAppXmlPath, _szDBName, Location, Department);
                objSource_Dir = objDir.GetDirPath(oFile.Destination_Directory, bStoreFilesinBlob, bPhysicalFileStorage, bIsEncryptionEnabled);
                if (bPhysicalFileStorage && oFile.Destination_Directory.Equals("DV"))
                {
                    objSource_Dir.Directory_Path = objSource_Dir.Directory_Path + "\\" + Convert.ToString(oFile.Draft_Version) + "\\";
                }
                bIsFile_Exist = Check_File_Exist(objSource_Dir, oFile);
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + msgError + Environment.NewLine);
                bIsFile_Exist = false;
            }
            finally
            {
                if (oFile != null)
                    oFile.Dispose();
                oFile = null;
                if (objDir != null)
                    objDir.Dispose();
                objDir = null;
                if (objSource_Dir != null)
                    objSource_Dir.Dispose();
                objSource_Dir = null;
                if (objDestination_Dir != null)
                    objDestination_Dir.Dispose();
                objDestination_Dir = null;
            }
            return bIsFile_Exist;
        }

        public bool Check_File_is_Locked(string szFileName)
        {
            bool bIsFile_Locked = true;
            try
            {
                objDir = new ClsDocumentDirPath(_szAppXmlPath, _szDBName, Location, Department);
                objSource_Dir = objDir.GetDirPath("TS", bStoreFilesinBlob, bPhysicalFileStorage, bIsEncryptionEnabled);
                bIsFile_Locked = Check_File_is_Locked(objSource_Dir, szFileName);
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + msgError + Environment.NewLine);
            }
            finally
            {
                if (objDir != null)
                    objDir.Dispose();
                objDir = null;
                if (objSource_Dir != null)
                    objSource_Dir.Dispose();
                objSource_Dir = null;
                if (objDestination_Dir != null)
                    objDestination_Dir.Dispose();
                objDestination_Dir = null;
            }
            return bIsFile_Locked;
        }

        public bool Check_File_is_Locked(string szFileName, string szOfficeVersion)
        {
            bool bIsFile_Locked = true;
            try
            {
                objDir = new ClsDocumentDirPath(_szAppXmlPath, _szDBName, Location, Department);
                objSource_Dir = objDir.GetDirPath("TS", bStoreFilesinBlob, bPhysicalFileStorage, bIsEncryptionEnabled);
                bIsFile_Locked = Check_File_is_Locked(objSource_Dir, szFileName, szOfficeVersion);
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + msgError + Environment.NewLine);
            }
            finally
            {
                if (objDir != null)
                    objDir.Dispose();
                objDir = null;
                if (objSource_Dir != null)
                    objSource_Dir.Dispose();
                objSource_Dir = null;
                if (objDestination_Dir != null)
                    objDestination_Dir.Dispose();
                objDestination_Dir = null;
            }
            return bIsFile_Locked;
        }


        public bool Pre_Check_File(string szFileName)
        {
            bool bIsFile_Locked = true;
            try
            {
                objDir = new ClsDocumentDirPath(_szAppXmlPath, _szDBName, Location, Department);
                objSource_Dir = objDir.GetDirPath("TS", bStoreFilesinBlob, bPhysicalFileStorage, bIsEncryptionEnabled);
                Pre_Check_Document(objSource_Dir, szFileName);
            }
            catch (Exception ex)
            {
                bIsFile_Locked = false;
                msgError = ex.Message;
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + msgError + Environment.NewLine);
            }
            finally
            {
                if (objDir != null)
                    objDir.Dispose();
                objDir = null;
                if (objSource_Dir != null)
                    objSource_Dir.Dispose();
                objSource_Dir = null;
                if (objDestination_Dir != null)
                    objDestination_Dir.Dispose();
                objDestination_Dir = null;
            }
            return bIsFile_Locked;
        }
        public bool Pre_Check_File(byte[] arrFile)
        {
            bool bIsFile_Locked = true;
            try
            {
                Pre_Check_Document(arrFile);
            }
            catch (Exception ex)
            {
                bIsFile_Locked = false;
                msgError = ex.Message;
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + msgError + Environment.NewLine);
            }
            finally
            {
                if (objDir != null)
                    objDir.Dispose();
                objDir = null;
                if (objSource_Dir != null)
                    objSource_Dir.Dispose();
                objSource_Dir = null;
                if (objDestination_Dir != null)
                    objDestination_Dir.Dispose();
                objDestination_Dir = null;
            }
            return bIsFile_Locked;
        }
        public bool Pre_Check_File(File_Data oFile)
        {
            bool bIsFile_Locked = true;
            try
            {
                objDir = new ClsDocumentDirPath(_szAppXmlPath, _szDBName, Location, Department);
                objSource_Dir = objDir.GetDirPath(oFile.Source_Directory, bStoreFilesinBlob, bPhysicalFileStorage, bIsEncryptionEnabled);
                Pre_Check_Document(objSource_Dir, oFile.File_Name);
            }
            catch (Exception ex)
            {
                bIsFile_Locked = false;
                msgError = ex.Message;
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + msgError + Environment.NewLine);
            }
            finally
            {
                if (objDir != null)
                    objDir.Dispose();
                objDir = null;
                if (objSource_Dir != null)
                    objSource_Dir.Dispose();
                objSource_Dir = null;
                if (objDestination_Dir != null)
                    objDestination_Dir.Dispose();
                objDestination_Dir = null;
            }
            return bIsFile_Locked;
        }



        public bool Delete_File(File_Data oFile)
        {
            bool bResult = true;
            try
            {
                objDir = new ClsDocumentDirPath(_szAppXmlPath, _szDBName, Location, Department);
                objSource_Dir = objDir.GetDirPath(oFile.Source_Directory, bStoreFilesinBlob, bPhysicalFileStorage, bIsEncryptionEnabled);
                bResult = Delete_File(objSource_Dir, oFile);
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
                bResult = false;
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + msgError + Environment.NewLine);
            }
            finally
            {
                if (objDir != null)
                    objDir.Dispose();
                objDir = null;
                if (objSource_Dir != null)
                    objSource_Dir.Dispose();
                objSource_Dir = null;
                if (objDestination_Dir != null)
                    objDestination_Dir.Dispose();
                objDestination_Dir = null;
            }
            return bResult;
        }

        public File_Data Get_File_Information(File_Data oFile)
        {
            try
            {

                objDir = new ClsDocumentDirPath(_szAppXmlPath, _szDBName, Location, Department);
                objSource_Dir = objDir.GetDirPath(oFile.Source_Directory, bStoreFilesinBlob, bPhysicalFileStorage, bIsEncryptionEnabled);
                if (bPhysicalFileStorage && oFile.Source_Directory.Equals("DV"))
                {
                    objSource_Dir.Directory_Path = objSource_Dir.Directory_Path + "\\" + Convert.ToString(oFile.Draft_Version) + "\\";
                }

                oFile = Get_File_Infomation(objSource_Dir, oFile);

            }
            catch (Exception ex)
            {
                msgError = ex.Message;
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + msgError + Environment.NewLine);
            }
            finally
            {
                if (objDir != null)
                    objDir.Dispose();
                objDir = null;
                if (objSource_Dir != null)
                    objSource_Dir.Dispose();
                objSource_Dir = null;
                if (objDestination_Dir != null)
                    objDestination_Dir.Dispose();
                objDestination_Dir = null;
            }
            return oFile;
        }

        public List<File_Data> Get_Documents(File_Data oFile)
        {
            List<File_Data> lstFile_Data = new List<File_Data>();
            try
            {
                objDir = new ClsDocumentDirPath(_szAppXmlPath, _szDBName, Location, Department);
                objSource_Dir = objDir.GetDirPath(oFile.Source_Directory, bStoreFilesinBlob, bPhysicalFileStorage, bIsEncryptionEnabled);
                lstFile_Data = Get_Documents(objSource_Dir, oFile);
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + msgError + Environment.NewLine);
            }
            finally
            {
                if (objDir != null)
                    objDir.Dispose();
                objDir = null;

                if (objSource_Dir != null)
                    objSource_Dir.Dispose();
                objSource_Dir = null;

                if (objDestination_Dir != null)
                    objDestination_Dir.Dispose();
                objDestination_Dir = null;
                if (oFile != null)
                    oFile.Dispose();
                oFile = null;
            }
            return lstFile_Data;
        }

        public string Get_File_Checksum(File_Data oFile)
        {
            string szDocumentCheckSum = string.Empty;
            try
            {
                objDir = new ClsDocumentDirPath(_szAppXmlPath, _szDBName, Location, Department);
                objSource_Dir = objDir.GetDirPath(oFile.Source_Directory, bStoreFilesinBlob, bPhysicalFileStorage, bIsEncryptionEnabled);
                szDocumentCheckSum = Get_Document_CheckSum(objSource_Dir, oFile);
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + msgError + Environment.NewLine);
            }
            finally
            {
                if (oFile != null)
                    oFile.Data = null;
                oFile = null;
                if (objDir != null)
                    objDir.Dispose();
                objDir = null;
                if (objSource_Dir != null)
                    objSource_Dir.Dispose();
                objSource_Dir = null;
                if (objDestination_Dir != null)
                    objDestination_Dir.Dispose();
                objDestination_Dir = null;
            }
            return szDocumentCheckSum;
        }

        public bool Convert_Document_to_Pdf(File_Data oFile)
        {
            bool bResult = true;
            try
            {
                objDir = new ClsDocumentDirPath(_szAppXmlPath, _szDBName, Location, Department);
                objSource_Dir = objDir.GetDirPath(oFile.Source_Directory, bStoreFilesinBlob, bPhysicalFileStorage, bIsEncryptionEnabled);
                if (!string.IsNullOrEmpty(oFile.Destination_Directory))
                    objDestination_Dir = objDir.GetDirPath(oFile.Destination_Directory, bStoreFilesinBlob, bPhysicalFileStorage, bIsEncryptionEnabled);

                Convert_To_PDF(objSource_Dir, oFile, objDestination_Dir, oFile);
            }
            catch (Exception ex)
            {
                bResult = false;
                msgError = ex.Message;
                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + msgError + Environment.NewLine);
            }
            finally
            {
                if (oFile != null)
                    oFile.Data = null;
                oFile = null;
                if (objDir != null)
                    objDir.Dispose();
                objDir = null;
                if (objSource_Dir != null)
                    objSource_Dir.Dispose();
                objSource_Dir = null;
                if (objDestination_Dir != null)
                    objDestination_Dir.Dispose();
                objDestination_Dir = null;
            }
            return bResult;
        }





        #endregion

        #region .... Read Excel File ...

        public DataSet Read_Excel_File(string szFilePath)
        {
            string szCommenction_String = Get_Connection_String(szFilePath);
            DataSet objDs = new DataSet();
            try
            {
                foreach (var sheetName in GetExcelSheetNames(szCommenction_String))
                {
                    using (OleDbConnection con = new OleDbConnection(szCommenction_String))
                    {
                        var dataTable = new DataTable();
                        string query = string.Format("SELECT * FROM [{0}]", sheetName);
                        con.Open();
                        OleDbDataAdapter adapter = new OleDbDataAdapter(query, con);
                        adapter.Fill(dataTable);
                        objDs.Tables.Add(dataTable);
                    }
                }
            }
            catch (Exception ex)
            {
                objDs = null;
                msgError = ex.Message;
            }
            return objDs;

        }
        private string Get_Connection_String(string szPath)
        {
            string connectionString = string.Empty;

            switch (Path.GetExtension(szPath).ToUpper())
            {
                case ".XLS":
                    connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + szPath + ";Extended Properties=\"Excel 8.0;HDR=YES;\"";
                    break;
                case ".XLSX":
                    connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + szPath + ";Extended Properties=\"Excel 12.0;HDR=YES;\"";
                    break;
                default:
                    break;
            }

            return connectionString;

        }
        static string[] GetExcelSheetNames(string connectionString)
        {
            String[] excelSheetNames = null;
            OleDbConnection con = null;
            DataTable dt = null;
            try
            {
                using (con = new OleDbConnection(connectionString))
                {
                    con.Open();
                    dt = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                    if (dt == null)
                        return null;

                    excelSheetNames = new String[dt.Rows.Count];
                    int i = 0;

                    foreach (DataRow row in dt.Rows)
                    {
                        excelSheetNames[i] = row["TABLE_NAME"].ToString();
                        i++;
                    }
                }
            }
            finally
            {
                if (dt != null)
                {
                    dt.Dispose();
                    dt = null;
                }
                con = null;
            }
            return excelSheetNames;
        }


        #endregion

        #region .... Private Methods....

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
            //GC.Collect();
            //GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (objDir != null)
                    objDir.Dispose();
                objDir = null;

                if (objSource_Dir != null)
                    objSource_Dir.Dispose();
                objSource_Dir = null;

                if (objDestination_Dir != null)
                    objDestination_Dir.Dispose();
                objDestination_Dir = null;

                if (_oDestination != null)
                    _oDestination.Dispose();
                _oDestination = null;
                _objSave_as_Blob = null;

            }
            else
            {

            }
            _szAppXmlPath = string.Empty;
            _szDBName = string.Empty;
        }

        ~ClsCopy_Files()
        {
            Dispose(false);
        }


        #endregion
    }
}
