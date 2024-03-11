//using System;
//using System.Collections.Generic;
//using System.IO;
//using System.Linq;
//using System.Security.Cryptography;
//using System.Text;
//using System.Threading.Tasks;
//using eDocsDN_Get_Directory_Info;
//using eDocsDN_File_Encryption;
//using eDocsDN_Save_File_in_Database;
//using eDocDN_Get_Custom_Properies;
//using eDocsDN_OpenXml_Operations;
//using eDocsDN_ReadAppXml;
//using System.Transactions;
//using eDocDN_Document_Pre_Check;
//using System.Diagnostics;
//using eDocsDN_syncfusion_Operations;
//using WebDav;
//using System.Xml.Linq;
//using System.Security.Cryptography.X509Certificates;
//using System.Net;
//using System.Net.Security;

namespace eDocsDN_File_Operations
{
    //public class ClsCheck_Configuration_For_File_Storage : ClsFile_Encryption
    //{
    //    #region .... Variable Declaration ...
    //    //ClsBuildQuery _objDal = null;
    //    clsReadAppXml _objINI = null;
    //    string _szSqlQuery = string.Empty;
    //    string _szAppXmlPath = string.Empty;
    //    string _szDBName = string.Empty;
    //    string _szLocation = string.Empty;
    //    ClsSave_File_in_Database _objSave_as_Blob = null;
    //    ClsDocumentPre_Check _objDocPreCheck = null;
    //    string _szLogFileName = string.Empty;
    //    bool _bisDebug = false;

    //    #endregion

    //    #region .... Property .....
    //    public bool bStoreFilesinBlob { get; set; }
    //    public bool bPhysicalFileStorage { get; set; }
    //    public bool bIsEncryptionEnabled { get; set; }
    //    #endregion

    //    #region ..... Constructor ....
    //    public ClsCheck_Configuration_For_File_Storage(string szAppXmlPath, string szDBName, string szLocation)
    //    {
    //        msgError = string.Empty;
    //        _szAppXmlPath = szAppXmlPath;
    //        _szDBName = szDBName;
    //        _szLocation = szLocation;
    //        _szLogFileName = AppDomain.CurrentDomain.BaseDirectory + @"\Log\DetailLog.txt";
    //        if (!Directory.Exists(AppDomain.CurrentDomain.BaseDirectory + @"\Log\"))
    //            Directory.CreateDirectory(AppDomain.CurrentDomain.BaseDirectory + @"\Log\");
    //    }

    //    #endregion

    //    #region .... Enum ...
    //    public enum eType
    //    {
    //        encrypt = 0,
    //        decrypt
    //    }
    //    #endregion

    //    #region .... Public Method ....

    //    public bool Delete_File(Directory_Attributes oSource_Dir, File_Data oSourceFile)
    //    {
    //        bool bResult = true;
    //        _objSave_as_Blob = new ClsSave_File_in_Database(_szAppXmlPath, _szDBName, _szLocation);
    //        if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Delete_File " + Environment.NewLine);
    //        try
    //        {
    //            if (oSource_Dir.Physical_Directory && oSource_Dir.Files_To_Be_Encrypted)
    //                bResult = Check_File_Exist_in_Encrypted_Or_Decrypted_Format(Action.Encrypt, oSource_Dir, oSourceFile);
    //            if (oSource_Dir.Physical_Directory && !oSource_Dir.Files_To_Be_Encrypted)
    //            {
    //                if (File.Exists(oSource_Dir.Directory_Path + oSourceFile.File_Name))
    //                    File.Delete(oSource_Dir.Directory_Path + oSourceFile.File_Name);
    //            }

    //            #region .... Save Files in Database ....

    //            if (oSource_Dir.Database_Storage)
    //            {
    //                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Delete_File_From_Database " + Environment.NewLine);
    //                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " oSource_Dir " + oSource_Dir.Directory_Path + Environment.NewLine);
    //                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " oSourceFile " + oSourceFile.File_Name + Environment.NewLine);
    //                bResult = _objSave_as_Blob.Delete_File_From_Database(oSource_Dir, oSourceFile);
    //                if (_objSave_as_Blob.msgError != "")
    //                    throw new Exception("Error occured on delete operations in Database : " + _objSave_as_Blob.msgError);
    //            }
    //            #endregion

    //        }
    //        catch (Exception ex)
    //        {
    //            msgError = ex.ToString();
    //            bResult = false;
    //        }
    //        finally
    //        { }
    //        return bResult;
    //    }

    //    public bool Check_File_is_Locked(Directory_Attributes oSource_Dir, string szFileName)
    //    {
    //        bool bResult = false;
    //        //FileStream fs = null;
    //        string _szDocServer = string.Empty;
    //        try
    //        {
    //            if (oSource_Dir.Physical_Directory)
    //            {
    //                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Check File Exist " + Environment.NewLine);
    //                if (File.Exists(oSource_Dir.Directory_Path + szFileName))
    //                {
    //                    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " File Exist " + oSource_Dir.Directory_Path + szFileName + Environment.NewLine);
    //                    using (_objINI = new clsReadAppXml(_szAppXmlPath))
    //                    {
    //                        _szDocServer = _objINI.GetLocationVariable(_objINI.GetCurrentLocation(), "", "DocServer").Trim();
    //                    }
    //                    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " _szDocServer " + _szDocServer + Environment.NewLine);
    //                    ServicePointManager.ServerCertificateValidationCallback += delegate (Object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
    //                    {
    //                        return true;
    //                    };

    //                    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Server URL  " + _szDocServer + Environment.NewLine);


    //                    //HttpWebRequest request;
    //                    //NetworkCredential cred;
    //                    //request = (HttpWebRequest)WebRequest.Create(Path.Combine(_szDocServer, szFileName));
    //                    //request.Method = "PROPFIND";
    //                    //request.ContentType = "text/xml";
    //                    ////request.Credentials = CredentialCache.DefaultCredentials;
    //                    //request.Credentials = new System.Net.NetworkCredential("Administrator", "Espl123&");
    //                    //request.Headers.Add("Translate", "f");
    //                    //request.Headers.Add("Depth", "1");
    //                    //request.SendChunked = true;
    //                    //ServicePointManager.ServerCertificateValidationCallback += delegate (Object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
    //                    //{
    //                    //    return true;
    //                    //};
    //                    //using (Stream stream = request.GetRequestStream())
    //                    //{
    //                    //    //stream.Write(buffer, 0, buffer.Length);
    //                    //}
    //                    //using (WebResponse response = request.GetResponse())
    //                    //{
    //                    //    string content = new StreamReader(response.GetResponseStream()).ReadToEnd();
    //                    //    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Tested  HttpWebRequest" + content + Environment.NewLine);
    //                    //}

    //                    var clientParams = new WebDavClientParams { BaseAddress = new Uri(_szDocServer) };
    //                    using (var client = new WebDavClient(clientParams))
    //                    {
    //                        //var resuly = await client.Propfind("Sample.docx").Result;

    //                        var result = client.Propfind(szFileName).Result;
    //                        var iActiveLock = ((List<WebDavResource>)result.Resources)[0].ActiveLocks.Count;
    //                        var iActiveLock1 = (List<WebDavProperty>)(((List<WebDavResource>)result.Resources)[0].Properties);
    //                        if (!string.IsNullOrEmpty(iActiveLock1[2].Value))
    //                        {
    //                            //..MessageBox.Show(iActiveLock1[2].Value);
    //                            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " iActiveLock1[2].Value " + iActiveLock1[2].Value + Environment.NewLine);
    //                            if (!string.IsNullOrEmpty(iActiveLock1[2].Value))
    //                            {
    //                                XDocument xmlDoc = XDocument.Parse(iActiveLock1[2].Value);
    //                                XNamespace ns = "DAV:";

    //                                XElement peopleCounting = xmlDoc.Root.Element(ns + "owner");
    //                                string enter = peopleCounting.Element(ns + "href").Value.Trim();
    //                                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " enter " + enter + Environment.NewLine);
    //                                bResult = true;
    //                                //MessageBox.Show("Locked With " + enter);
    //                            }
    //                            else
    //                            {
    //                                bResult = false;
    //                                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " NO Lock " + Environment.NewLine);
    //                                //MessageBox.Show("NO Lock");
    //                            }

    //                        }
    //                    }
    //                }

    //                //if (File.Exists(oSource_Dir.Directory_Path + szFileName))
    //                //{
    //                //    fs = new FileStream(oSource_Dir.Directory_Path + szFileName, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
    //                //    fs.Close();
    //                //}
    //                //else
    //                //{
    //                //    bResult = false;
    //                //}
    //            }
    //        }
    //        catch (Exception ex)
    //        {
    //            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Error  " + ex.Message + Environment.NewLine);
    //            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Error  " + ex.InnerException + Environment.NewLine);
    //            msgError = ex.Message;
    //            bResult = false;
    //        }
    //        finally
    //        {
    //            //if (fs != null)
    //            //    fs.Close();
    //            //fs = null;
    //        }
    //        return bResult;
    //    }

    //    public bool Check_File_Exist(Directory_Attributes oSource_Dir, File_Data oSourceFile)
    //    {
    //        bool bResult = false;

    //        try
    //        {
    //            switch (oSource_Dir.Database_Storage)
    //            {
    //                case true:
    //                    _objSave_as_Blob = new ClsSave_File_in_Database(_szAppXmlPath, _szDBName, _szLocation);
    //                    bResult = _objSave_as_Blob.isRecord_Exist_In_Database(oSource_Dir, oSourceFile);
    //                    if (_objSave_as_Blob.msgError != "")
    //                        throw new Exception("Error occured on Save operations in Database : " + _objSave_as_Blob.msgError);
    //                    break;
    //                case false:
    //                    switch (oSource_Dir.Files_To_Be_Encrypted)
    //                    {
    //                        case true:
    //                            bResult = Check_File_Exist_in_Encrypted_Or_Decrypted_Format(Action.Encrypt, oSource_Dir, oSourceFile);
    //                            break;
    //                        case false:
    //                            bResult = File.Exists(oSource_Dir.Directory_Path + oSourceFile.File_Name);

    //                            break;
    //                    }
    //                    break;
    //            }
    //        }
    //        catch (Exception ex)
    //        {
    //            msgError = ex.ToString();
    //            bResult = false;
    //        }
    //        finally
    //        { }
    //        return bResult;
    //    }



    //    public File_Data Get_File_Infomation(Directory_Attributes oSource_Dir, File_Data oSourceFile)
    //    {
    //        bool bResult = false;
    //        try
    //        {
    //            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Get_File_Infomation " + Environment.NewLine);
    //            switch (oSource_Dir.Database_Storage)
    //            {
    //                case true:
    //                    _objSave_as_Blob = new ClsSave_File_in_Database(_szAppXmlPath, _szDBName, _szLocation);
    //                    oSourceFile = _objSave_as_Blob.Get_File_from_Database(oSource_Dir, oSourceFile);
    //                    if (_objSave_as_Blob.msgError != "")
    //                        throw new Exception("Error occured on Save operations in Database : " + _objSave_as_Blob.msgError);
    //                    break;
    //                case false:
    //                    switch (oSource_Dir.Files_To_Be_Encrypted)
    //                    {
    //                        case true:
    //                            bResult = Check_File_Exist_in_Encrypted_Or_Decrypted_Format(Action.Encrypt, oSource_Dir, oSourceFile);
    //                            break;
    //                        case false:
    //                            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + "before ReadAllBytes " + Environment.NewLine);
    //                            oSourceFile.Data = File.ReadAllBytes(oSource_Dir.Directory_Path + oSourceFile.File_Name);
    //                            oSourceFile.CheckSum = GetMd5_CheckSum(oSource_Dir.Directory_Path + oSourceFile.File_Name);
    //                            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + "end ReadAllBytes " + Environment.NewLine);
    //                            break;
    //                    }
    //                    break;
    //            }

    //        }
    //        catch (Exception ex)
    //        {
    //            msgError = ex.ToString();
    //            bResult = false;

    //            string message =
    //  "Exception type " + ex.GetType() + Environment.NewLine +
    //  "Exception message: " + ex.Message + Environment.NewLine +
    //  "Stack trace: " + ex.StackTrace + Environment.NewLine;
    //            if (ex.InnerException != null)
    //            {
    //                message += "---BEGIN InnerException--- " + Environment.NewLine +
    //                           "Exception type " + ex.InnerException.GetType() + Environment.NewLine +
    //                           "Exception message: " + ex.InnerException.Message + Environment.NewLine +
    //                           "Stack trace: " + ex.InnerException.StackTrace + Environment.NewLine +
    //                           "---END Inner Exception";
    //            }
    //            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + ex.Message + Environment.NewLine);
    //            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + ex.StackTrace + Environment.NewLine);
    //            //throw ex;
    //        }
    //        finally
    //        { }
    //        return oSourceFile;
    //    }
    //    public List<File_Data> Get_Documents(Directory_Attributes oSource_Dir, File_Data oSourceFile)
    //    {
    //        bool bResult = false;
    //        List<File_Data> lstFile_Data = new List<File_Data>();
    //        try
    //        {
    //            switch (oSource_Dir.Database_Storage)
    //            {
    //                case true:
    //                    _objSave_as_Blob = new ClsSave_File_in_Database(_szAppXmlPath, _szDBName, _szLocation);
    //                    lstFile_Data = _objSave_as_Blob.Get_Documents(oSource_Dir, oSourceFile);
    //                    if (_objSave_as_Blob.msgError != "")
    //                        throw new Exception("Error occured on Save operations in Database : " + _objSave_as_Blob.msgError);
    //                    break;

    //                case false:

    //                    switch (oSource_Dir.Files_To_Be_Encrypted)
    //                    {
    //                        case true:
    //                            bResult = Check_File_Exist_in_Encrypted_Or_Decrypted_Format(Action.Encrypt, oSource_Dir, oSourceFile);
    //                            break;
    //                        case false:
    //                            if (File.Exists(oSource_Dir.Directory_Path + oSourceFile.File_Name))
    //                            {
    //                                oSourceFile.Data = File.ReadAllBytes(oSource_Dir.Directory_Path + oSourceFile.File_Name);
    //                                oSourceFile.CheckSum = GetMd5_CheckSum(oSource_Dir.Directory_Path + oSourceFile.File_Name);
    //                                File_Data oFile = new File_Data();
    //                                oFile.Data = oSourceFile.Data;
    //                                oFile.Serial_Number = oSourceFile.Serial_Number;
    //                                oFile.CheckSum = oSourceFile.CheckSum;
    //                                oFile.Source_File_CheckSum = oSourceFile.CheckSum;
    //                                oFile.File_Name = oSourceFile.File_Name;
    //                                oFile.User_Id = "";
    //                                lstFile_Data.Add(oFile);

    //                            }
    //                            break;
    //                    }
    //                    break;

    //            }
    //        }
    //        catch (Exception ex)
    //        {
    //            msgError = ex.ToString();
    //            bResult = false;
    //        }
    //        finally
    //        { }
    //        return lstFile_Data;
    //    }
    //    public List<File_Data> Get_Document_List(Directory_Attributes oSource_Dir, File_Data oSourceFile)
    //    {
    //        bool bResult = false;
    //        List<File_Data> lstFile_Data = new List<File_Data>();
    //        try
    //        {
    //            switch (oSource_Dir.Database_Storage)
    //            {
    //                case true:
    //                    _objSave_as_Blob = new ClsSave_File_in_Database(_szAppXmlPath, _szDBName, _szLocation);
    //                    lstFile_Data = _objSave_as_Blob.Get_Documents(oSource_Dir, oSourceFile);
    //                    if (_objSave_as_Blob.msgError != "")
    //                        throw new Exception("Error occured on Save operations in Database : " + _objSave_as_Blob.msgError);
    //                    break;

    //                case false:

    //                    switch (oSource_Dir.Files_To_Be_Encrypted)
    //                    {
    //                        case true:
    //                            bResult = Check_File_Exist_in_Encrypted_Or_Decrypted_Format(Action.Encrypt, oSource_Dir, oSourceFile);
    //                            break;
    //                        case false:
    //                            oSourceFile.Data = File.ReadAllBytes(oSource_Dir.Directory_Path + oSourceFile.File_Name);
    //                            oSourceFile.CheckSum = GetMd5_CheckSum(oSource_Dir.Directory_Path + oSourceFile.File_Name);
    //                            lstFile_Data.Add(oSourceFile);
    //                            break;
    //                    }
    //                    break;

    //            }
    //        }
    //        catch (Exception ex)
    //        {
    //            msgError = ex.ToString();
    //            bResult = false;
    //        }
    //        finally
    //        { }
    //        return lstFile_Data;
    //    }


    //    public File_Data Copy_File(Directory_Attributes oSource_Dir, File_Data oSourceFile, Directory_Attributes oDestination_Dir, File_Data oDestination_File)
    //    {

    //        try
    //        {
    //            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " commented " + Environment.NewLine);

    //            #region ..... Get Source File Data .....
    //            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : oSource_Dir " + Environment.NewLine);
    //            if (oSource_Dir != null)
    //            {
    //                #region .... Get Source File ....
    //                switch (oSource_Dir.Physical_Directory)
    //                {
    //                    case true:
    //                        if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Physical_Directory: " + oSource_Dir.Physical_Directory.ToString() + Environment.NewLine);

    //                        if (oDestination_Dir.Physical_Directory)
    //                        {
    //                            if (!Directory.Exists(oDestination_Dir.Directory_Path))
    //                                Directory.CreateDirectory(oDestination_Dir.Directory_Path);
    //                        }

    //                        switch (oSource_Dir.Files_To_Be_Encrypted)
    //                        {
    //                            case true:
    //                                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Files_To_Be_Encrypted: " + oSource_Dir.Files_To_Be_Encrypted.ToString() + Environment.NewLine);
    //                                oSourceFile = Encrypt_Descypt_Files(Action.Decrypt, oSource_Dir, oSourceFile, oDestination_Dir, oDestination_File);
    //                                oDestination_File = oSourceFile;
    //                                break;
    //                            case false:
    //                                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Files_To_Be_Encrypted: " + oSource_Dir.Files_To_Be_Encrypted.ToString() + Environment.NewLine);
    //                                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " GetMd5_CheckSum: " + oSource_Dir.Directory_Path + oSourceFile.File_Name + Environment.NewLine);
    //                                oSourceFile.Source_File_CheckSum = GetMd5_CheckSum(oSource_Dir.Directory_Path + oSourceFile.File_Name);
    //                                oSourceFile.CheckSum = GetMd5_CheckSum(oSource_Dir.Directory_Path + oSourceFile.File_Name);
    //                                oSourceFile.Data = File.ReadAllBytes(oSource_Dir.Directory_Path + oSourceFile.File_Name);
    //                                oDestination_File = oSourceFile;


    //                                break;
    //                            default:
    //                                break;
    //                        }
    //                        break;
    //                    case false:
    //                        if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Physical_Directory: " + oSource_Dir.Physical_Directory.ToString() + Environment.NewLine);
    //                        #region .... get file from Database ....
    //                        _objSave_as_Blob = new ClsSave_File_in_Database(_szAppXmlPath, _szDBName, _szLocation);
    //                        if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Get_File_from_Database: " + Environment.NewLine);
    //                        oSourceFile = _objSave_as_Blob.Get_File_from_Database(oSource_Dir, oSourceFile);
    //                        if (_objSave_as_Blob.msgError != "")
    //                        {
    //                            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Error occured on Save operations in Database : " + _objSave_as_Blob.msgError + Environment.NewLine);
    //                            throw new Exception("Error occured on Save operations in Database : " + _objSave_as_Blob.msgError);
    //                        }
    //                        oDestination_File = oSourceFile;

    //                        #endregion

    //                        break;
    //                    default:
    //                        break;
    //                }
    //                #endregion
    //            }
    //            else
    //            {
    //                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " oDestination_File = oSourceFile" + Environment.NewLine);
    //                oDestination_File = oSourceFile;
    //            }
    //            #endregion

    //            #region .... Process File ....
    //            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + "Process_Document" + Environment.NewLine);

    //            if (oDestination_File.Data.Length < 42162774)
    //            {
    //                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + "Process_Document_Stream" + Environment.NewLine);
    //                oDestination_File = Process_Document_Stream(oDestination_File);
    //            }
    //            else
    //            {
    //                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + "Process_Document_Physically" + Environment.NewLine);
    //                oDestination_File = Process_Document_Physically(oDestination_File);
    //            }


    //            #endregion

    //            #region .... Write Destination File ...
    //            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Write Destination File " + Environment.NewLine);

    //            if (oDestination_Dir.Physical_Directory && oDestination_Dir.Files_To_Be_Encrypted)
    //                Encrypt_Descypt_Files(Action.Encrypt, oSource_Dir, oSourceFile, oDestination_Dir, oDestination_File);

    //            if (oDestination_Dir.Physical_Directory && !oDestination_Dir.Files_To_Be_Encrypted)
    //            {
    //                if (!Directory.Exists(oDestination_Dir.Directory_Path))
    //                    Directory.CreateDirectory(oDestination_Dir.Directory_Path);

    //                File.WriteAllBytes(oDestination_Dir.Directory_Path + oDestination_File.Destination_File_Name, oDestination_File.Data);
    //                oDestination_File.CheckSum = GetMd5_CheckSum(oDestination_Dir.Directory_Path + oDestination_File.Destination_File_Name);
    //            }
    //            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Writed Destination File " + Environment.NewLine);


    //            #region .... Save Files in Database ....

    //            if (oDestination_Dir.Database_Storage)
    //            {
    //                _objSave_as_Blob = new ClsSave_File_in_Database(_szAppXmlPath, _szDBName, _szLocation);
    //                oDestination_File = _objSave_as_Blob.Save_File_In_Database(oDestination_Dir, oDestination_File);
    //                if (_objSave_as_Blob.msgError != "")
    //                    throw new Exception("Error occured on Save operations in Database : " + _objSave_as_Blob.msgError);
    //            }

    //            #endregion

    //            #endregion
    //        }
    //        catch (Exception ex)
    //        {
    //            string message =
    //  "Exception type " + ex.GetType() + Environment.NewLine +
    //  "Exception message: " + ex.Message + Environment.NewLine +
    //  "Stack trace: " + ex.StackTrace + Environment.NewLine;
    //            if (ex.InnerException != null)
    //            {
    //                message += "---BEGIN InnerException--- " + Environment.NewLine +
    //                           "Exception type " + ex.InnerException.GetType() + Environment.NewLine +
    //                           "Exception message: " + ex.InnerException.Message + Environment.NewLine +
    //                           "Stack trace: " + ex.InnerException.StackTrace + Environment.NewLine +
    //                           "---END Inner Exception";
    //            }
    //            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + ex.Message + Environment.NewLine);
    //            throw ex;

    //        }
    //        finally
    //        {
    //            _objSave_as_Blob = null;
    //        }
    //        return oDestination_File;
    //    }

    //    public bool View_File(Directory_Attributes oSource_Dir, File_Data oSourceFile, Directory_Attributes oDestination_Dir, File_Data oDestination_File)
    //    {
    //        bool bResult = true;
    //        try
    //        {
    //            if (!Directory.Exists(oDestination_Dir.Directory_Path))
    //                Directory.CreateDirectory(oDestination_Dir.Directory_Path);

    //            if (oDestination_Dir.Physical_Directory && oDestination_Dir.Files_To_Be_Encrypted)
    //                Encrypt_Descypt_Files(Action.Decrypt, oSource_Dir, oSourceFile, oDestination_Dir, oDestination_File);

    //            if (oDestination_Dir.Physical_Directory && !oDestination_Dir.Files_To_Be_Encrypted)
    //            {
    //                File.WriteAllBytes(oDestination_Dir.Directory_Path + oDestination_File.Destination_File_Name, oSourceFile.Data);
    //                oDestination_File.CheckSum = GetMd5_CheckSum(oDestination_Dir.Directory_Path + oDestination_File.Destination_File_Name);
    //            }
    //        }
    //        catch (Exception ex)
    //        {
    //            msgError = ex.ToString();
    //            bResult = false;
    //        }
    //        finally
    //        { }
    //        return bResult;
    //    }

    //    public string Get_Document_CheckSum(Directory_Attributes oSource_Dir, File_Data oSourceFile)
    //    {
    //        bool bResult = false;

    //        try
    //        {
    //            if (oSource_Dir.Physical_Directory && oSource_Dir.Files_To_Be_Encrypted)
    //                bResult = Check_File_Exist_in_Encrypted_Or_Decrypted_Format(Action.Encrypt, oSource_Dir, oSourceFile);
    //            if (oSource_Dir.Physical_Directory && !oSource_Dir.Files_To_Be_Encrypted)
    //                oSourceFile.CheckSum = GetMd5_CheckSum(oSource_Dir.Directory_Path + oSourceFile.File_Name);

    //            #region .... Save Files in Database ....

    //            if (oSource_Dir.Database_Storage)
    //            {
    //                _objSave_as_Blob = new ClsSave_File_in_Database(_szAppXmlPath, _szDBName, _szLocation);
    //                oSourceFile = _objSave_as_Blob.Get_File_Checksum_From_Database(oSource_Dir, oSourceFile);
    //                if (_objSave_as_Blob.msgError != "")
    //                    throw new Exception("Error occured on Get File Checksum From Database : " + _objSave_as_Blob.msgError);
    //            }
    //            #endregion

    //        }
    //        catch (Exception ex)
    //        {
    //            msgError = ex.ToString();
    //            bResult = false;
    //        }
    //        finally
    //        { }
    //        return oSourceFile.CheckSum;
    //    }

    //    internal void Pre_Check_Document(Directory_Attributes oSource_Dir, string szFileName)
    //    {
    //        bool bAllowEmbedeImages = false;
    //        try
    //        {
    //            #region .... Check Document ...

    //            if (oSource_Dir.Physical_Directory)
    //            {
    //                using (_objINI = new clsReadAppXml(_szAppXmlPath))
    //                {
    //                    if (_objINI.IsWordDocument.Contains(Path.GetExtension(szFileName).Replace(".", "").ToUpper()))
    //                    {
    //                        bAllowEmbedeImages = _objINI.GetLocationVariable(_objINI.GetCurrentLocation(), "", "AllowEmbededImages") == "1" ? true : false;
    //                        if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " bAllowEmbedeImages  :" + bAllowEmbedeImages.ToString() + Environment.NewLine);
    //                        if (File.Exists(oSource_Dir.Directory_Path + szFileName))
    //                        {
    //                            _objDocPreCheck = new ClsDocumentPre_Check(oSource_Dir.Directory_Path + szFileName, bAllowEmbedeImages);
    //                            if (!_objDocPreCheck.PreCheck_Document())
    //                                throw new Exception(_objDocPreCheck.msgError);
    //                        }
    //                    }
    //                }
    //            }

    //            #endregion
    //        }
    //        finally { }
    //    }
    //    internal void Pre_Check_Document(byte[] arrFile)
    //    {
    //        bool bAllowEmbedeImages = false;
    //        try
    //        {
    //            #region .... Check Document ...
    //            using (_objINI = new clsReadAppXml(_szAppXmlPath))
    //            {
    //                bAllowEmbedeImages = _objINI.GetLocationVariable(_objINI.GetCurrentLocation(), "", "AllowEmbededImages") == "1" ? true : false;
    //            }
    //            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " bAllowEmbedeImages  :" + bAllowEmbedeImages.ToString() + Environment.NewLine);
    //            _objDocPreCheck = new ClsDocumentPre_Check(Convert_Document_To_Stream(arrFile), bAllowEmbedeImages);
    //            if (!_objDocPreCheck.PreCheck_Document())
    //                throw new Exception(_objDocPreCheck.msgError);

    //            #endregion
    //        }
    //        finally { }
    //    }



    //    #endregion

    //    #region .... Private functions ....

    //    private File_Data Process_Document_Stream(File_Data oSourceFile)
    //    {
    //        Stream strmDocument = null;
    //        ClsDocument_Operations objDocOperations = null;
    //        LockUnlockFile objLockUnlock = null;
    //        Update_Document_Custom_Variables objUpdate_Document_Properties = null;
    //        Update_Users_Comments objUpdate_Comments = null;
    //        File_Operations objFileOperations = oSourceFile.File_Operations;
    //        if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Process_Document_Stream " + Environment.NewLine);
    //        try
    //        {
    //            _objINI = new clsReadAppXml(_szAppXmlPath);
    //            if (_objINI.IsWordDocument.Contains(Path.GetExtension(oSourceFile.File_Name).Replace(".", "").ToUpper()))
    //            {
    //                #region .... Process File ....

    //                if (oSourceFile.File_Operations != null)
    //                {
    //                    objLockUnlock = null;
    //                    objUpdate_Document_Properties = null;
    //                    objUpdate_Comments = null;
    //                    objFileOperations = oSourceFile.File_Operations;

    //                    #region .... convert to Stream ....
    //                    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Convert_Document_To_Stream " + Environment.NewLine);
    //                    strmDocument = Convert_Document_To_Stream(oSourceFile.Data);
    //                    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Converted Document To Stream " + Environment.NewLine);

    //                    #endregion

    //                    switch (objFileOperations.ConvertToPdf)
    //                    {
    //                        case true:
    //                            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Convert To Pdf : " + objFileOperations.ConvertToPdf.ToString() + Environment.NewLine);
    //                            Process_Physical_Document(oSourceFile);

    //                            break;
    //                        default:

    //                            #region .... Set Track Changes to Off ...
    //                            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : TrackRevisions set to false" + Environment.NewLine);
    //                            strmDocument = clsOpenXml_Operations.TrackRevisions(strmDocument, false);
    //                            if (!string.IsNullOrEmpty(clsOpenXml_Operations.msgError))
    //                                throw new Exception(clsOpenXml_Operations.msgError);
    //                            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : TrackRevisions set to false Complited" + Environment.NewLine);

    //                            #endregion

    //                            #region .... Clear Existing Comments from Documents ...

    //                            if (objFileOperations.Update_Properties != null)
    //                            {
    //                                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " DeleteAllCommentsFromDocument " + Environment.NewLine);
    //                                if (objFileOperations.Update_Properties.eDocument_Process == Documents_Process.Controller_Live)
    //                                {
    //                                    if (!clsOpenXml_Operations.DeleteAllCommentsFromDocument(strmDocument))
    //                                        throw new Exception(clsOpenXml_Operations.msgError);
    //                                }
    //                            }
    //                            #endregion

    //                            #region ... Update Scan Signatues ...

    //                            if (objFileOperations.ScanSignature != null)
    //                            {
    //                                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Chck for Scan Sign" + Environment.NewLine);
    //                                Scan_Signature oScanSign = objFileOperations.ScanSignature;
    //                                eDocDN_Update_ScanSign.ClsUpdate_ScanSign objScanSign = new eDocDN_Update_ScanSign.ClsUpdate_ScanSign(strmDocument);
    //                                if (oScanSign.Remove_Scan_Sign)
    //                                    strmDocument = objScanSign.RemoveScanSign();
    //                                if (oScanSign.Users_Scan_Sign != null)
    //                                {
    //                                    if (oScanSign.Users_Scan_Sign.Count > 0)
    //                                        strmDocument = objScanSign.UpdateScanSign(oScanSign.Users_Scan_Sign, oScanSign.Remove_Scan_Sign);
    //                                }
    //                                if (objScanSign.msgError != "")
    //                                    throw new Exception("Error Occured while Updatinf ScanSign. Error : " + objScanSign.msgError);
    //                                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Chck for Scan Sign Complited" + Environment.NewLine);
    //                            }

    //                            #endregion

    //                            #region .... Update Document Properties ...


    //                            if (objFileOperations.Update_Properties != null)
    //                            {
    //                                objUpdate_Document_Properties = objFileOperations.Update_Properties;
    //                                objDocOperations = new ClsDocument_Operations(_szDBName, _szAppXmlPath);
    //                                strmDocument = objDocOperations.Update_Custom_Variables(oSourceFile, strmDocument, objUpdate_Document_Properties.eDocument_Status, objUpdate_Document_Properties.eDocument_Process);
    //                                if (!string.IsNullOrEmpty(objDocOperations.msgError))
    //                                    throw new Exception("Error occured while Updating Document Properties. Error:" + objDocOperations.msgError);
    //                                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Update_Properties Complited" + Environment.NewLine);
    //                            }
    //                            #endregion

    //                            #region .... Process File Physically ....

    //                            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Unlock Document" + Environment.NewLine);
    //                            objDocOperations = new ClsDocument_Operations();
    //                            strmDocument = objDocOperations.Lock_Unlock_Document(strmDocument, ClsDocument_Operations.Lock_Unlock.Unlock, LockType.None, true);
    //                            if (objDocOperations.msgError != "")
    //                                throw new Exception("Error occured while Lock Unlock. Error:" + objDocOperations.msgError);

    //                            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Unlock Document Complited" + Environment.NewLine);

    //                            oSourceFile.Data = strmDocument.ReadAllBytes();
    //                            if (strmDocument != null)
    //                            {
    //                                strmDocument.Flush();
    //                                strmDocument.Dispose();
    //                            }
    //                            strmDocument = null;
    //                            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Process_Physical_Document" + Environment.NewLine);
    //                            oSourceFile = Process_Physical_Document(oSourceFile);
    //                            strmDocument = Convert_Document_To_Stream(oSourceFile.Data);
    //                            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Processed physical File Complited" + Environment.NewLine);

    //                            #endregion

    //                            #region .... Update Comments ....
    //                            if (objFileOperations.UpdateComments != null)
    //                            {
    //                                objUpdate_Comments = objFileOperations.UpdateComments;
    //                                ClsUpdate_Comment_Author oUpdate_Comments = new ClsUpdate_Comment_Author();
    //                                strmDocument = oUpdate_Comments.Set_Author_To_Comments(strmDocument, objUpdate_Comments.UserID, objUpdate_Comments.dtDateTime);
    //                                if (string.IsNullOrEmpty(oUpdate_Comments.msgError))
    //                                    throw new Exception("Error Occured while Updating User Comments in Document.  Error :" + oUpdate_Comments.msgError);
    //                                oUpdate_Comments = null;

    //                            }
    //                            #endregion

    //                            #region .... Print Form Data ...
    //                            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Set Print Form Data" + Environment.NewLine);
    //                            strmDocument = clsOpenXml_Operations.PrintFormsData(strmDocument, oSourceFile.PrintFormData);
    //                            if (!string.IsNullOrEmpty(clsOpenXml_Operations.msgError))
    //                                throw new Exception(clsOpenXml_Operations.msgError);

    //                            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Set Print Form Data Complited" + Environment.NewLine);
    //                            #endregion

    //                            #region .... Lock Unlock ....
    //                            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Locak Unlick Operation" + Environment.NewLine);
    //                            if (objFileOperations.LockUnlock != null)
    //                            {
    //                                objLockUnlock = objFileOperations.LockUnlock;
    //                                switch (objLockUnlock.LockFile)
    //                                {
    //                                    case true:

    //                                        objDocOperations = new ClsDocument_Operations();
    //                                        strmDocument = objDocOperations.Lock_Unlock_Document(strmDocument, ClsDocument_Operations.Lock_Unlock.Lock, (LockType)objLockUnlock.Lock_Type, true);
    //                                        if (objDocOperations.msgError != "")
    //                                            throw new Exception(objDocOperations.msgError);
    //                                        oSourceFile.Data = strmDocument.ReadAllBytes();
    //                                        break;
    //                                    case false:
    //                                        objDocOperations = new ClsDocument_Operations();
    //                                        strmDocument = objDocOperations.Lock_Unlock_Document(strmDocument, ClsDocument_Operations.Lock_Unlock.Unlock, LockType.None, true);
    //                                        if (objDocOperations.msgError != "")
    //                                            throw new Exception(objDocOperations.msgError);

    //                                        break;
    //                                }
    //                            }
    //                            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Locak Unlick Operation Complited" + Environment.NewLine);
    //                            #endregion

    //                            #region .... Track Changes ...
    //                            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Set TrackRevisions" + Environment.NewLine);
    //                            strmDocument = clsOpenXml_Operations.TrackRevisions(strmDocument, oSourceFile.TrackChanges);
    //                            if (!string.IsNullOrEmpty(clsOpenXml_Operations.msgError))
    //                                throw new Exception(clsOpenXml_Operations.msgError);
    //                            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Set TrackRevisions Complited" + Environment.NewLine);
    //                            #endregion

    //                            #region .... Prepare Document For Printting ...

    //                            if (objFileOperations.Print_Documents != null)
    //                            {
    //                                if (objFileOperations.Print_Documents.Clear_comments)
    //                                    if (!clsOpenXml_Operations.DeleteAllCommentsFromDocument(strmDocument))
    //                                        throw new Exception(clsOpenXml_Operations.msgError);

    //                                strmDocument = clsOpenXml_Operations.PrintFormsData(strmDocument, false);
    //                                if (!string.IsNullOrEmpty(clsOpenXml_Operations.msgError))
    //                                    throw new Exception(clsOpenXml_Operations.msgError);
    //                            }
    //                            #endregion

    //                            break;
    //                    }



    //                    #region ... Convert to Byte[] ...
    //                    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Convert stream To byte" + Environment.NewLine);
    //                    oSourceFile.Data = strmDocument.ReadAllBytes();
    //                    oSourceFile.CheckSum = GetMd5_CheckSum(oSourceFile.Data);
    //                    oSourceFile.Source_File_CheckSum = oSourceFile.CheckSum;
    //                    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Converted stream To byte" + Environment.NewLine);
    //                    #endregion
    //                }

    //                #endregion
    //            }
    //            else if (_objINI.IsExcelDocument.Contains(Path.GetExtension(oSourceFile.File_Name).Replace(".", "").ToUpper()))
    //            {
    //                #region ..... Process Excel Document ....
    //                if (oSourceFile.File_Operations != null)
    //                {
    //                    objLockUnlock = null;
    //                    objUpdate_Document_Properties = null;
    //                    objUpdate_Comments = null;
    //                    objFileOperations = oSourceFile.File_Operations;


    //                    #region .... convert to Stream ....
    //                    strmDocument = Convert_Document_To_Stream(oSourceFile.Data);
    //                    #endregion

    //                    #region ...... Unlock Document ...
    //                    objDocOperations = new ClsDocument_Operations();
    //                    strmDocument = objDocOperations.Lock_Unlock_Excel_Document(strmDocument, ClsDocument_Operations.Lock_Unlock.Unlock, LockType.None, true);
    //                    if (objDocOperations.msgError != "")
    //                        throw new Exception("Error occured while Lock Unlock. Error:" + objDocOperations.msgError);
    //                    #endregion

    //                    #region .... Lock Unlock ....

    //                    if (objFileOperations.LockUnlock != null)
    //                    {
    //                        objLockUnlock = objFileOperations.LockUnlock;
    //                        switch (objLockUnlock.LockFile)
    //                        {
    //                            case true:

    //                                objDocOperations = new ClsDocument_Operations();
    //                                strmDocument = objDocOperations.Lock_Unlock_Excel_Document(strmDocument, ClsDocument_Operations.Lock_Unlock.Lock, (LockType)objLockUnlock.Lock_Type, true);
    //                                if (objDocOperations.msgError != "")
    //                                    throw new Exception(objDocOperations.msgError);
    //                                oSourceFile.Data = strmDocument.ReadAllBytes();
    //                                break;
    //                            case false:
    //                                objDocOperations = new ClsDocument_Operations();
    //                                strmDocument = objDocOperations.Lock_Unlock_Excel_Document(strmDocument, ClsDocument_Operations.Lock_Unlock.Unlock, LockType.None, true);
    //                                if (objDocOperations.msgError != "")
    //                                    throw new Exception(objDocOperations.msgError);

    //                                break;
    //                        }
    //                    }
    //                    #endregion

    //                    #region ... Convert to Byte[] ...

    //                    oSourceFile.Data = strmDocument.ReadAllBytes();
    //                    oSourceFile.CheckSum = GetMd5_CheckSum(oSourceFile.Data);
    //                    oSourceFile.Source_File_CheckSum = oSourceFile.CheckSum;
    //                    #endregion

    //                }
    //                #endregion
    //            }
    //            else
    //            {
    //                #region ..... Process Excel Document ....
    //                if (oSourceFile.Data != null)
    //                {
    //                    oSourceFile.CheckSum = GetMd5_CheckSum(oSourceFile.Data);
    //                    oSourceFile.Source_File_CheckSum = oSourceFile.CheckSum;
    //                }
    //                #endregion
    //            }
    //        }
    //        catch (Exception ex)
    //        {
    //            string message =
    // "Exception type " + ex.GetType() + Environment.NewLine +
    // "Exception message: " + ex.Message + Environment.NewLine +
    // "Stack trace: " + ex.StackTrace + Environment.NewLine;
    //            if (ex.InnerException != null)
    //            {
    //                message += "---BEGIN InnerException--- " + Environment.NewLine +
    //                           "Exception type " + ex.InnerException.GetType() + Environment.NewLine +
    //                           "Exception message: " + ex.InnerException.Message + Environment.NewLine +
    //                           "Stack trace: " + ex.InnerException.StackTrace + Environment.NewLine +
    //                           "---END Inner Exception";
    //            }
    //            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + ex.Message + Environment.NewLine);
    //            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + ex.StackTrace + Environment.NewLine);
    //            throw ex;
    //        }
    //        finally
    //        {
    //            if (_objINI != null)
    //                _objINI.Dispose();
    //            _objINI = null;
    //            if (strmDocument != null)
    //            {
    //                strmDocument.Close();
    //                strmDocument.Flush();
    //                strmDocument.Dispose();
    //            }
    //            strmDocument = null;
    //            objDocOperations = null;
    //        }
    //        return oSourceFile;
    //    }

    //    private File_Data Process_Document_Physically(File_Data oSourceFile)
    //    {
    //        Stream strmDocument = null;
    //        ClsDocument_Operations objDocOperations = null;
    //        LockUnlockFile objLockUnlock = null;
    //        Update_Document_Custom_Variables objUpdate_Document_Properties = null;
    //        Update_Users_Comments objUpdate_Comments = null;
    //        File_Operations objFileOperations = oSourceFile.File_Operations;
    //        ClsDocumentDirPath objDir = null;
    //        Directory_Attributes objTempDir = null;
    //        if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + "Process_Document_Physically" + Environment.NewLine);
    //        objDir = new ClsDocumentDirPath(_szAppXmlPath, _szDBName, _szLocation, "");
    //        objTempDir = new Directory_Attributes();
    //        objTempDir = objDir.GetDirPath("TS", false, true, false);
    //        if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + "TS : " + objTempDir.Directory_Path + Environment.NewLine);
    //        try
    //        {
    //            _objINI = new clsReadAppXml(_szAppXmlPath);
    //            if (_objINI.IsWordDocument.Contains(Path.GetExtension(oSourceFile.File_Name).Replace(".", "").ToUpper()))
    //            {
    //                #region .... Process File ....

    //                if (oSourceFile.File_Operations != null)
    //                {
    //                    objLockUnlock = null;
    //                    objUpdate_Document_Properties = null;
    //                    objUpdate_Comments = null;
    //                    objFileOperations = oSourceFile.File_Operations;
    //                    clsOpenXml_Operations.msgError = string.Empty;


    //                    #region .... convert to Stream ....

    //                    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + "WriteAllBytes" + Environment.NewLine);
    //                    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + "objTempDir.Directory_Path :" + objTempDir.Directory_Path + Environment.NewLine);
    //                    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " oSourceFile.Destination_File_Name :" + oSourceFile.Destination_File_Name + Environment.NewLine);
    //                    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " oSourceFile.SourceFilePath :" + oSourceFile.SourceFilePath + Environment.NewLine);
    //                    File.WriteAllBytes(objTempDir.Directory_Path + oSourceFile.Destination_File_Name, oSourceFile.Data);

    //                    #endregion

    //                    #region .... Set Track Changes to Off ...
    //                    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : TrackRevisions set to false" + Environment.NewLine);
    //                    clsOpenXml_Operations.TrackRevisions(objTempDir.Directory_Path + oSourceFile.Destination_File_Name, false);
    //                    if (!string.IsNullOrEmpty(clsOpenXml_Operations.msgError))
    //                        throw new Exception(clsOpenXml_Operations.msgError);
    //                    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : TrackRevisions set to false Complited" + Environment.NewLine);

    //                    #endregion

    //                    #region .... Print Form Data ...
    //                    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Set Print Form Data" + Environment.NewLine);
    //                    clsOpenXml_Operations.PrintFormsData(objTempDir.Directory_Path + oSourceFile.Destination_File_Name, oSourceFile.PrintFormData);
    //                    if (!string.IsNullOrEmpty(clsOpenXml_Operations.msgError))
    //                        throw new Exception(clsOpenXml_Operations.msgError);

    //                    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Set Print Form Data Complited" + Environment.NewLine);
    //                    #endregion

    //                    #region ... Update Scan Signatues ...

    //                    if (objFileOperations.ScanSignature != null)
    //                    {
    //                        if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Chck for Scan Sign" + Environment.NewLine);
    //                        Scan_Signature oScanSign = objFileOperations.ScanSignature;
    //                        eDocDN_Update_ScanSign.ClsUpdate_ScanSign objScanSign = new eDocDN_Update_ScanSign.ClsUpdate_ScanSign(objTempDir.Directory_Path + oSourceFile.Destination_File_Name);
    //                        if (oScanSign.Remove_Scan_Sign)
    //                            strmDocument = objScanSign.RemoveScanSign();
    //                        if (oScanSign.Users_Scan_Sign != null)
    //                        {
    //                            if (oScanSign.Users_Scan_Sign.Count > 0)
    //                                strmDocument = objScanSign.UpdateScanSign(oScanSign.Users_Scan_Sign, oScanSign.Remove_Scan_Sign);
    //                        }
    //                        if (!string.IsNullOrEmpty(objScanSign.msgError))
    //                            throw new Exception(objScanSign.msgError);
    //                        if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Chck for Scan Sign Complited" + Environment.NewLine);
    //                    }

    //                    #endregion

    //                    #region .... Update Document Properties ...

    //                    if (objFileOperations.Update_Properties != null)
    //                    {
    //                        if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Update_Properties" + Environment.NewLine);
    //                        objUpdate_Document_Properties = objFileOperations.Update_Properties;
    //                        objDocOperations = new ClsDocument_Operations(_szDBName, _szAppXmlPath);
    //                        objDocOperations.Update_Custom_Variables(oSourceFile, objTempDir.Directory_Path + oSourceFile.Destination_File_Name, objUpdate_Document_Properties.eDocument_Status, objUpdate_Document_Properties.eDocument_Process);
    //                        if (!string.IsNullOrEmpty(objDocOperations.msgError))
    //                            throw new Exception(objDocOperations.msgError);
    //                        if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Update_Properties Complited" + Environment.NewLine);
    //                    }
    //                    #endregion

    //                    #region .... Update Comments ....
    //                    if (objFileOperations.UpdateComments != null)
    //                    {
    //                        objUpdate_Comments = objFileOperations.UpdateComments;
    //                        ClsUpdate_Comment_Author oUpdate_Comments = new ClsUpdate_Comment_Author();
    //                        oUpdate_Comments.Set_Author_To_Comments(objTempDir.Directory_Path + oSourceFile.Destination_File_Name, objUpdate_Comments.UserID, objUpdate_Comments.dtDateTime);
    //                        if (string.IsNullOrEmpty(oUpdate_Comments.msgError))
    //                            throw new Exception("Error Occured while Updating User Comments in Document.  Error :" + oUpdate_Comments.msgError);
    //                        oUpdate_Comments = null;

    //                    }
    //                    #endregion

    //                    #region .... Process File Physically ....

    //                    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Unlock Document" + Environment.NewLine);
    //                    objDocOperations = new ClsDocument_Operations();
    //                    objDocOperations.Lock_Unlock_Document(objTempDir.Directory_Path + oSourceFile.Destination_File_Name, ClsDocument_Operations.Lock_Unlock.Unlock, LockType.None, true);
    //                    if (objDocOperations.msgError != "")
    //                        throw new Exception(objDocOperations.msgError);

    //                    Process_Physical_Document(oSourceFile, objTempDir.Directory_Path + oSourceFile.Destination_File_Name);
    //                    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Processed physical File Complited" + Environment.NewLine);

    //                    #endregion

    //                    #region .... Lock Unlock ....
    //                    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Locak Unlick Operation" + Environment.NewLine);
    //                    if (objFileOperations.LockUnlock != null)
    //                    {
    //                        objLockUnlock = objFileOperations.LockUnlock;
    //                        switch (objLockUnlock.LockFile)
    //                        {
    //                            case true:

    //                                objDocOperations = new ClsDocument_Operations();
    //                                objDocOperations.Lock_Unlock_Document(objTempDir.Directory_Path + oSourceFile.Destination_File_Name, ClsDocument_Operations.Lock_Unlock.Lock, (LockType)objLockUnlock.Lock_Type, true);
    //                                if (objDocOperations.msgError != "")
    //                                    throw new Exception(objDocOperations.msgError);
    //                                break;
    //                            case false:
    //                                objDocOperations = new ClsDocument_Operations();
    //                                objDocOperations.Lock_Unlock_Document(objTempDir.Directory_Path + oSourceFile.Destination_File_Name, ClsDocument_Operations.Lock_Unlock.Unlock, LockType.None, true);
    //                                if (objDocOperations.msgError != "")
    //                                    throw new Exception(objDocOperations.msgError);

    //                                break;
    //                        }
    //                    }
    //                    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Locak Unlick Operation Complited" + Environment.NewLine);
    //                    #endregion

    //                    #region .... Track Changes ...
    //                    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Set TrackRevisions" + Environment.NewLine);
    //                    clsOpenXml_Operations.TrackRevisions(objTempDir.Directory_Path + oSourceFile.Destination_File_Name, oSourceFile.TrackChanges);
    //                    if (!string.IsNullOrEmpty(clsOpenXml_Operations.msgError))
    //                        throw new Exception(clsOpenXml_Operations.msgError);
    //                    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Set TrackRevisions Complited" + Environment.NewLine);
    //                    #endregion

    //                    #region .... Prepare Document For Printting ...

    //                    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : PrintFormsData" + Environment.NewLine);
    //                    if (objFileOperations.Print_Documents != null)
    //                    {
    //                        if (objFileOperations.Print_Documents.Clear_comments)
    //                            if (!clsOpenXml_Operations.DeleteAllCommentsFromDocument(objTempDir.Directory_Path + oSourceFile.Destination_File_Name))
    //                                throw new Exception(clsOpenXml_Operations.msgError);

    //                        clsOpenXml_Operations.PrintFormsData(objTempDir.Directory_Path + oSourceFile.Destination_File_Name, false);
    //                        if (!string.IsNullOrEmpty(clsOpenXml_Operations.msgError))
    //                            throw new Exception(clsOpenXml_Operations.msgError);
    //                    }
    //                    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : PrintFormsData C" + Environment.NewLine);

    //                    #endregion

    //                    #region ... Convert physical File ...

    //                    #region .... Return Process File ....
    //                    if (File.Exists(objTempDir.Directory_Path + oSourceFile.Destination_File_Name))
    //                    {
    //                        WaitForFile(objTempDir.Directory_Path + oSourceFile.Destination_File_Name);
    //                        oSourceFile.Data = File.ReadAllBytes(objTempDir.Directory_Path + oSourceFile.Destination_File_Name);
    //                        oSourceFile.CheckSum = GetMd5_CheckSum(objTempDir.Directory_Path + oSourceFile.Destination_File_Name);
    //                        oSourceFile.Source_File_CheckSum = oSourceFile.CheckSum;
    //                    }

    //                    #endregion

    //                    if (!File.Exists(objTempDir.Directory_Path + oSourceFile.Destination_File_Name))
    //                        throw new Exception("File Not found " + objTempDir.Directory_Path + oSourceFile.Destination_File_Name);

    //                    //WaitForFile(objTempDir.Directory_Path + oSourceFile.Destination_File_Name);
    //                    //File.Delete(objTempDir.Directory_Path + oSourceFile.Destination_File_Name);
    //                    //oSourceFile.Destination_File_Name = oSourceFile.File_Name;
    //                    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Complete" + Environment.NewLine);


    //                    #endregion
    //                }

    //                #endregion
    //            }
    //            else if (_objINI.IsExcelDocument.Contains(Path.GetExtension(oSourceFile.File_Name).Replace(".", "").ToUpper()))
    //            {
    //                #region ..... Process Excel Document .....

    //                if (oSourceFile.File_Operations != null)
    //                {
    //                    objLockUnlock = null;
    //                    objUpdate_Document_Properties = null;
    //                    objUpdate_Comments = null;
    //                    objFileOperations = oSourceFile.File_Operations;

    //                    #region .... convert to Stream ....
    //                    strmDocument = Convert_Document_To_Stream(oSourceFile.Data);
    //                    #endregion

    //                    #region .... Lock Unlock ....

    //                    if (objFileOperations.LockUnlock != null)
    //                    {
    //                        objLockUnlock = objFileOperations.LockUnlock;
    //                        switch (objLockUnlock.LockFile)
    //                        {
    //                            case true:

    //                                objDocOperations = new ClsDocument_Operations();
    //                                strmDocument = objDocOperations.Lock_Unlock_Excel_Document(strmDocument, ClsDocument_Operations.Lock_Unlock.Lock, (LockType)objLockUnlock.Lock_Type, true);
    //                                if (objDocOperations.msgError != "")
    //                                    throw new Exception(objDocOperations.msgError);
    //                                oSourceFile.Data = strmDocument.ReadAllBytes();
    //                                break;
    //                            case false:
    //                                objDocOperations = new ClsDocument_Operations();
    //                                strmDocument = objDocOperations.Lock_Unlock_Excel_Document(strmDocument, ClsDocument_Operations.Lock_Unlock.Unlock, LockType.None, true);
    //                                if (objDocOperations.msgError != "")
    //                                    throw new Exception(objDocOperations.msgError);

    //                                break;
    //                        }
    //                    }
    //                    #endregion

    //                    #region ... Convert to Byte[] ...

    //                    oSourceFile.Data = strmDocument.ReadAllBytes();
    //                    oSourceFile.CheckSum = GetMd5_CheckSum(oSourceFile.Data);
    //                    oSourceFile.Source_File_CheckSum = oSourceFile.CheckSum;
    //                    #endregion
    //                }
    //                #endregion
    //            }
    //            else
    //            {
    //                #region ..... Other Document ....
    //                if (oSourceFile.Data != null)
    //                {
    //                    oSourceFile.CheckSum = GetMd5_CheckSum(oSourceFile.Data);
    //                    oSourceFile.Source_File_CheckSum = oSourceFile.CheckSum;
    //                }
    //                #endregion
    //            }
    //        }
    //        catch (Exception ex)
    //        {
    //            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + ex.Message + Environment.NewLine);
    //            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + ex.StackTrace + Environment.NewLine);
    //            throw ex;
    //        }
    //        finally
    //        {
    //            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Delete physical File" + Environment.NewLine);
    //            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Source File " + oSourceFile.Source_Directory + Environment.NewLine);
    //            if (objTempDir != null)
    //                if (File.Exists(objTempDir.Directory_Path + oSourceFile.Destination_File_Name))
    //                {
    //                    WaitForFile(objTempDir.Directory_Path + oSourceFile.Destination_File_Name);
    //                    File.Delete(objTempDir.Directory_Path + oSourceFile.Destination_File_Name);
    //                    if (objFileOperations != null)
    //                    {
    //                        if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " :" + oSourceFile.Source_Directory + Environment.NewLine);
    //                        switch (objFileOperations.ConvertToPdf)
    //                        {
    //                            case true:
    //                                break;
    //                            default:
    //                                if (objFileOperations.Update_Properties != null)
    //                                {
    //                                    switch (objFileOperations.Update_Properties.eDocument_Process)
    //                                    {
    //                                        case Documents_Process.Document_Issuance:
    //                                            break;
    //                                        default:
    //                                            oSourceFile.Destination_File_Name = oSourceFile.File_Name;
    //                                            break;
    //                                    }
    //                                }
    //                                else if (objFileOperations.Print_Documents != null)
    //                                {
    //                                    //oSourceFile.Destination_File_Name = oSourceFile.File_Name;
    //                                }
    //                                else
    //                                {
    //                                    oSourceFile.Destination_File_Name = oSourceFile.File_Name;
    //                                }
    //                                break;
    //                        }
    //                    }

    //                }
    //            if (_objINI != null)
    //                _objINI.Dispose();
    //            _objINI = null;
    //            if (strmDocument != null)
    //            {
    //                strmDocument.Flush();
    //                strmDocument.Dispose();
    //            }
    //            strmDocument = null;
    //            objDocOperations = null;
    //        }
    //        return oSourceFile;
    //    }

    //    private File_Data Process_Physical_Document(File_Data oDestinationFile)
    //    {
    //        ClsDocumentDirPath objDir = new ClsDocumentDirPath(_szAppXmlPath, _szDBName, _szLocation, "");
    //        ClsXml_Operations objWord_Operations;
    //        Update_Document_Custom_Variables objUpdate_Document_Properties = null;
    //        File_Operations objFileOperations = null;
    //        Directory_Attributes objTempDir = new Directory_Attributes();
    //        objTempDir = objDir.GetDirPath("TS", false, true, false);
    //        try
    //        {
    //            _objINI = new clsReadAppXml(_szAppXmlPath);
    //            if (_objINI.IsWordDocument.Contains(Path.GetExtension(oDestinationFile.Destination_File_Name).Replace(".", "").ToUpper()))
    //            {
    //                if (File.Exists(objTempDir.Directory_Path + oDestinationFile.Destination_File_Name))
    //                    File.Delete(objTempDir.Directory_Path + oDestinationFile.Destination_File_Name);

    //                if (oDestinationFile.File_Operations != null)
    //                {
    //                    objFileOperations = oDestinationFile.File_Operations;
    //                    switch (objFileOperations.ConvertToPdf)
    //                    {
    //                        case true:
    //                            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Convert To Pdf : " + objFileOperations.ConvertToPdf.ToString() + Environment.NewLine);

    //                            #region .... Convert To PDF ...
    //                            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Destination_File_Name : " + objTempDir.Directory_Path + oDestinationFile.Destination_File_Name + Environment.NewLine);
    //                            File.WriteAllBytes(objTempDir.Directory_Path + oDestinationFile.Destination_File_Name, oDestinationFile.Data);
    //                            objWord_Operations = new ClsXml_Operations(_szAppXmlPath);
    //                            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Convert_DOcument_To_PDF " + Environment.NewLine);

    //                            if (!objWord_Operations.Convert_Document_To_PDF(objTempDir.Directory_Path + oDestinationFile.Destination_File_Name))
    //                            {
    //                                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Error Occured while Converting File To PDF : " + objWord_Operations.msgError + Environment.NewLine);
    //                                throw new Exception("Error Occured while Converting Document to PDF. Error : " + objWord_Operations.msgError);
    //                            }
    //                            #endregion

    //                            break;
    //                        default:

    //                            #region .... BackEnd Process ...
    //                            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Process_Physical_Document" + Environment.NewLine);

    //                            if (objFileOperations.Update_Properties != null)
    //                            {
    //                                objUpdate_Document_Properties = objFileOperations.Update_Properties;
    //                                switch (objUpdate_Document_Properties.eDocument_Process)
    //                                {
    //                                    case Documents_Process.Controller_Live:
    //                                    case Documents_Process.Transfer_Document:
    //                                    case Documents_Process.Document_Recall:
    //                                    case Documents_Process.obsolete_Document:
    //                                    case Documents_Process.Expired_Document:
    //                                    case Documents_Process.Document_Issuance:
    //                                    case Documents_Process.TR4:
    //                                    case Documents_Process.Controller_Publish:
    //                                        File.WriteAllBytes(objTempDir.Directory_Path + oDestinationFile.Destination_File_Name, oDestinationFile.Data);
    //                                        objWord_Operations = new ClsXml_Operations(_szAppXmlPath);
    //                                        if (!objWord_Operations.Update_Document_Properties(objTempDir.Directory_Path + oDestinationFile.Destination_File_Name))
    //                                        {
    //                                            throw new Exception("Error Occured while Processing Document(Physical). Error : " + objWord_Operations.msgError + Environment.NewLine);
    //                                        }
    //                                        if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : objUpdate_Document_Properties.eDocument_Process" + objUpdate_Document_Properties.eDocument_Process.ToString() + Environment.NewLine);


    //                                        break;
    //                                    case Documents_Process.Preview:
    //                                        if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : WriteAllBytes " + objUpdate_Document_Properties.eDocument_Process.ToString() + Environment.NewLine);

    //                                        File.WriteAllBytes(objTempDir.Directory_Path + oDestinationFile.Destination_File_Name, oDestinationFile.Data);
    //                                        if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : WriteAllBytes Completed " + objUpdate_Document_Properties.eDocument_Process.ToString() + Environment.NewLine);

    //                                        objWord_Operations = new ClsXml_Operations(_szAppXmlPath);
    //                                        if (!objWord_Operations.Process_Word_Document(objTempDir.Directory_Path + oDestinationFile.Destination_File_Name))
    //                                            throw new Exception("Error Occured while Processing Document(Physical). Error : " + objWord_Operations.msgError);
    //                                        if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : objUpdate_Document_Properties.eDocument_Process " + objUpdate_Document_Properties.eDocument_Process.ToString() + Environment.NewLine);

    //                                        break;
    //                                    default:
    //                                        break;
    //                                }
    //                            }
    //                            #endregion

    //                            break;
    //                    }
    //                }

    //                #region .... Return Process File ....
    //                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Convert File To Stream : " + objTempDir.Directory_Path + oDestinationFile.Destination_File_Name + Environment.NewLine);
    //                if (File.Exists(objTempDir.Directory_Path + oDestinationFile.Destination_File_Name))
    //                {
    //                    WaitForFile(objTempDir.Directory_Path + oDestinationFile.Destination_File_Name);
    //                    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Read File : " + objTempDir.Directory_Path + oDestinationFile.Destination_File_Name + Environment.NewLine);
    //                    oDestinationFile.Data = File.ReadAllBytes(objTempDir.Directory_Path + oDestinationFile.Destination_File_Name);
    //                    oDestinationFile.CheckSum = GetMd5_CheckSum(objTempDir.Directory_Path + oDestinationFile.Destination_File_Name);
    //                    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " CheckSum : " + oDestinationFile.CheckSum + Environment.NewLine);
    //                    oDestinationFile.Source_File_CheckSum = oDestinationFile.CheckSum;
    //                }
    //                if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Converted Physical File to Stream" + Environment.NewLine);

    //                #endregion

    //            }
    //            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : END : " + _szAppXmlPath + Environment.NewLine);


    //        }
    //        catch (Exception ex)
    //        {
    //            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + ex.Message + Environment.NewLine);
    //            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + ex.StackTrace + Environment.NewLine);
    //            throw ex;
    //        }
    //        finally
    //        {
    //            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Delete physical File" + Environment.NewLine);
    //            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Source File " + oDestinationFile.Source_Directory + Environment.NewLine);
    //            if (File.Exists(objTempDir.Directory_Path + oDestinationFile.Destination_File_Name))
    //            {
    //                WaitForFile(objTempDir.Directory_Path + oDestinationFile.Destination_File_Name);
    //                File.Delete(objTempDir.Directory_Path + oDestinationFile.Destination_File_Name);
    //                if (oDestinationFile.Source_Directory.Equals("MI", StringComparison.InvariantCultureIgnoreCase))
    //                {
    //                    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " oDestinationFile.File_Name " + oDestinationFile.File_Name + Environment.NewLine);
    //                    if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " oDestinationFile.Destination_File_Name " + oDestinationFile.Destination_File_Name + Environment.NewLine);
    //                    oDestinationFile.File_Name = oDestinationFile.Destination_File_Name;

    //                };
    //                if (objFileOperations != null)
    //                {
    //                    switch (objFileOperations.ConvertToPdf)
    //                    {
    //                        case true:
    //                            break;
    //                        default:
    //                            if (objFileOperations.Update_Properties != null)
    //                            {
    //                                switch (objFileOperations.Update_Properties.eDocument_Process)
    //                                {
    //                                    case Documents_Process.Document_Issuance:
    //                                        break;
    //                                    default:
    //                                        oDestinationFile.Destination_File_Name = oDestinationFile.File_Name;
    //                                        break;
    //                                }
    //                            }
    //                            else
    //                            {
    //                                oDestinationFile.Destination_File_Name = oDestinationFile.File_Name;
    //                            }
    //                            break;
    //                    }
    //                }

    //            }
    //            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Deleted physical File" + Environment.NewLine);
    //            objWord_Operations = null;
    //            if (_objINI != null)
    //                _objINI.Dispose();
    //            _objINI = null;
    //        }
    //        return oDestinationFile;

    //    }


    //    private File_Data Process_Physical_Document(File_Data oDestinationFile, string szFilePath)
    //    {
    //        ClsXml_Operations objWord_Operations = null;
    //        try
    //        {
    //            _objINI = new clsReadAppXml(_szAppXmlPath);
    //            if (_objINI.IsWordDocument.Contains(Path.GetExtension(szFilePath).Replace(".", "").ToUpper()))
    //            {
    //                if (oDestinationFile.File_Operations != null)
    //                {
    //                    Update_Document_Custom_Variables objUpdate_Document_Properties = null;
    //                    File_Operations objFileOperations = oDestinationFile.File_Operations;
    //                    switch (objFileOperations.ConvertToPdf)
    //                    {
    //                        case true:

    //                            #region .... Convert To PDF ...
    //                            objWord_Operations = new ClsXml_Operations(_szAppXmlPath);
    //                            if (!objWord_Operations.Convert_Document_To_PDF(szFilePath))
    //                                throw new Exception("Error Occured while Converting Document to PDF. Error : " + objWord_Operations.msgError);
    //                            #endregion

    //                            break;
    //                        default:

    //                            #region .... BackEnd Process ....
    //                            if (objFileOperations.Update_Properties != null)
    //                            {
    //                                objUpdate_Document_Properties = objFileOperations.Update_Properties;
    //                                switch (objUpdate_Document_Properties.eDocument_Process)
    //                                {
    //                                    case Documents_Process.Controller_Live:
    //                                    case Documents_Process.Transfer_Document:
    //                                    case Documents_Process.Document_Recall:
    //                                    case Documents_Process.obsolete_Document:
    //                                    case Documents_Process.Expired_Document:
    //                                    case Documents_Process.Document_Issuance:
    //                                    case Documents_Process.TR4:
    //                                    case Documents_Process.Controller_Publish:
    //                                        if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : objUpdate_Document_Properties.eDocument_Process  : " + objUpdate_Document_Properties.eDocument_Process.ToString() + Environment.NewLine);

    //                                        objWord_Operations = new ClsXml_Operations(_szAppXmlPath);
    //                                        if (!objWord_Operations.Update_Document_Properties(szFilePath))
    //                                        {
    //                                            throw new Exception("Error Occured while Processing Document(Physical). Error : " + objWord_Operations.msgError + Environment.NewLine);
    //                                        }
    //                                        if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : objUpdated Document_Properties.eDocument_Process  " + Environment.NewLine);

    //                                        break;
    //                                    case Documents_Process.Preview:
    //                                        if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Preview Process " + Environment.NewLine);

    //                                        objWord_Operations = new ClsXml_Operations(_szAppXmlPath);
    //                                        if (!objWord_Operations.Process_Word_Document(szFilePath))
    //                                            throw new Exception("Error Occured while Processing Document(Physical). Error : " + objWord_Operations.msgError);
    //                                        if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " Previewed Process " + Environment.NewLine);
    //                                        break;
    //                                    default:
    //                                        break;
    //                                }
    //                            }
    //                            #endregion

    //                            break;
    //                    }

    //                }
    //            }
    //        }
    //        catch (Exception ex)
    //        {
    //            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + ex.Message + Environment.NewLine);
    //            if (_bisDebug) File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + ex.StackTrace + Environment.NewLine);
    //            throw ex;
    //        }
    //        finally
    //        {

    //            objWord_Operations = null;
    //            if (_objINI != null)
    //                _objINI.Dispose();
    //            _objINI = null;
    //        }
    //        return oDestinationFile;

    //    }


    //    //public bool isDebugEnable(string szAppXmlPath)
    //    //{
    //    //    bool bIsDebugEnable = false;
    //    //    string szDebugStatus = string.Empty;
    //    //    try
    //    //    {
    //    //        _objINI = new clsReadAppXml(szAppXmlPath);
    //    //        szDebugStatus = _objINI.GetLocationVariable(_szLocation, "", "DebugStatus");
    //    //        if (_objINI.ErrorMsg != "")
    //    //            throw new Exception(_objINI.ErrorMsg);
    //    //        if (szDebugStatus.Equals("TRUE", StringComparison.InvariantCultureIgnoreCase))
    //    //            bIsDebugEnable = true;

    //    //    }
    //    //    finally { }
    //    //    return bIsDebugEnable;
    //    //}



    //    public void WaitForFile(string filename)
    //    {
    //        while (!IsFileReady(filename)) { }
    //    }
    //    public bool IsFileReady(string filename)
    //    {
    //        try
    //        {
    //            using (FileStream inputStream = File.Open(filename, FileMode.Open, FileAccess.Read, FileShare.None))
    //                return inputStream.Length > 0;
    //        }
    //        catch (Exception)
    //        {
    //            return false;
    //        }
    //    }


    //    internal static Stream Convert_Document_To_Stream(byte[] arrDocument)
    //    {
    //        MemoryStream strmDocument = new MemoryStream();
    //        strmDocument.Write(arrDocument, 0, (int)arrDocument.Length);
    //        return strmDocument;
    //    }

    //    //internal bool File_Encryption_Enabled()
    //    //{
    //    //    bool bEncrypt_Files = false;
    //    //    try
    //    //    {
    //    //        _objDal = new ClsBuildQuery(_szDBName, _szAppXmlPath);
    //    //        _objDal.OpenConnection(ProductName.DocsExecutive, "NA");

    //    //        if (_objDal.msgError != "")
    //    //            throw new Exception("Error Occured While Opening Database Connection");

    //    //        _szSqlQuery = "select eno_yek from zespl_sys where epyt='FE'";
    //    //        bEncrypt_Files = _objDal.GetSingleValue(_szSqlQuery) == 1 ? true : false;
    //    //        if (_objDal.msgError != "")
    //    //            throw new Exception("Error occured While reading File Encryption Flag : " + _objDal.msgError);
    //    //    }
    //    //    finally
    //    //    {
    //    //        if (_objDal != null)
    //    //        {
    //    //            _objDal.CloseConnection();
    //    //            _objDal.Dispose();
    //    //        }
    //    //        _objDal = null;
    //    //    }
    //    //    return bEncrypt_Files;
    //    //}

    //    //internal void File_Storage_in_Database()
    //    //{
    //    //    var option = new TransactionOptions();
    //    //    try
    //    //    {
    //    //        option.IsolationLevel = System.Transactions.IsolationLevel.ReadUncommitted;
    //    //        option.Timeout = TimeSpan.FromMinutes(5);
    //    //        //File.AppendAllText(_szLogFileName, DateTime.Now + " : TransactionScope Start " + Environment.NewLine);
    //    //        using (var scope = new TransactionScope(TransactionScopeOption.RequiresNew, option))
    //    //        //using (var scope = new TransactionScope())
    //    //        {
    //    //            //File.AppendAllText(_szLogFileName, DateTime.Now + " : _szDBName " + _szDBName + Environment.NewLine);
    //    //            //File.AppendAllText(_szLogFileName, DateTime.Now + " : _szAppXmlPath " + _szAppXmlPath + Environment.NewLine);
    //    //            _objDal = new ClsBuildQuery(_szDBName, _szAppXmlPath);
    //    //            _objDal.OpenConnection(ConnectionFor.DocsExecutive, ConnectionType.NewConnection);

    //    //            if (_objDal.msgError != "")
    //    //            {
    //    //                throw new Exception("Error Occured While Opening Database Connection. Error:" + _objDal.msgError);
    //    //            }

    //    //            _szSqlQuery = "select eno_yek from zespl_sys where epyt='FS'";
    //    //            bStoreFilesinBlob = _objDal.GetSingleValue(_szSqlQuery) == 1 ? true : false;
    //    //            if (_objDal.msgError != "")
    //    //                throw new Exception("Error occured While reading File Encryption Flag : " + _objDal.msgError);

    //    //            _szSqlQuery = "select eno_yek from zespl_sys where epyt='FE'";
    //    //            bIsEncryptionEnabled = _objDal.GetSingleValue(_szSqlQuery) == 1 ? true : false;
    //    //            if (_objDal.msgError != "")
    //    //                throw new Exception("Error occured While reading File Encryption Flag : " + _objDal.msgError);

    //    //            scope.Complete();
    //    //            //File.AppendAllText(_szLogFileName, DateTime.Now + " : TransactionScope Complete " + Environment.NewLine);
    //    //        }

    //    //    }
    //    //    catch (Exception ex)
    //    //    {
    //    //        //File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + ex.Message + Environment.NewLine);
    //    //        //File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + ex.StackTrace + Environment.NewLine);
    //    //        throw ex;
    //    //    }
    //    //    finally
    //    //    {
    //    //        if (_objDal != null)
    //    //        {
    //    //            _objDal.CloseConnection();
    //    //            _objDal.Dispose();
    //    //        }
    //    //        _objDal = null;
    //    //    }
    //    //}
    //    internal void File_Storage_in_Database()
    //    {
    //        string szFileSystem = string.Empty;
    //        string szEncryptFiles = string.Empty;
    //        try
    //        {
    //            _objINI = new clsReadAppXml(_szAppXmlPath);
    //            szFileSystem = _objINI.GetLocationVariable(_objINI.GetCurrentLocation(), "", "FileSystem");
    //            szEncryptFiles = _objINI.GetLocationVariable(_objINI.GetCurrentLocation(), "", "EncryptFile");
    //            _bisDebug = _objINI.GetLocationVariable(_objINI.GetCurrentLocation(), "", "DebugStatus").ToUpper() == "TRUE" ? true : false;
    //            switch (szFileSystem)
    //            {
    //                case "0": //..Physical
    //                    bStoreFilesinBlob = false;
    //                    bPhysicalFileStorage = true;
    //                    bIsEncryptionEnabled = szEncryptFiles == "0" ? false : true;
    //                    break;
    //                case "1":
    //                    bStoreFilesinBlob = true;
    //                    bPhysicalFileStorage = false;
    //                    bIsEncryptionEnabled = false;
    //                    break;
    //                case "2":
    //                    bStoreFilesinBlob = true;
    //                    bPhysicalFileStorage = true;
    //                    bIsEncryptionEnabled = szEncryptFiles == "0" ? false : true;
    //                    break;

    //                default:
    //                    break;
    //            }


    //        }
    //        catch (Exception ex)
    //        {
    //            //File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + ex.Message + Environment.NewLine);
    //            //File.AppendAllText(_szLogFileName, DateTime.Now + " : Error : " + ex.StackTrace + Environment.NewLine);
    //            throw ex;
    //        }
    //        finally
    //        {
    //            if (_objINI != null)
    //                _objINI.Dispose();
    //            _objINI = null;
    //        }
    //    }



    //    #endregion
    //}


    //public static class StreamExtensions
    //{
    //    public static byte[] ReadAllBytes(this Stream instream)
    //    {
    //        if (instream is MemoryStream)
    //            return ((MemoryStream)instream).ToArray();

    //        using (var memoryStream = new MemoryStream())
    //        {
    //            instream.CopyTo(memoryStream);
    //            return memoryStream.ToArray();
    //        }
    //    }
    //}

}
