using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DDLLCS;
using eDocsDN_Get_Directory_Info;
using System.Collections;
using System.Data;
using System.IO;
using System.Transactions;

namespace eDocsDN_Save_File_in_Database
{
    public class ClsSave_File_in_Database : IDisposable
    {
        #region .... Variable Declaration ....
        ClsBuildQuery _objDal = null;
        string _szAppXmlPath = string.Empty;
        string _szDBName = string.Empty;
        string _Location = string.Empty;
        string _szSqlQuery = string.Empty;
        string _szLogFileName = string.Empty;


        #endregion

        #region .... Property ...
        public string msgError { get; set; }
        #endregion

        #region .... Constroctor ...
        public ClsSave_File_in_Database(string szAppXmlPath, string szDBName, string szLocation)
        {
            msgError = string.Empty;
            _szAppXmlPath = szAppXmlPath;
            _szDBName = szDBName;
            _Location = szLocation;
            _szLogFileName = AppDomain.CurrentDomain.BaseDirectory + @"\Log\DB.txt";
        }
        #endregion

        #region ..... Public Method ...
        public List<File_Data> Save_File_In_Database(Directory_Attributes oDestination_Dir, List<File_Data> lstFile_Data)
        {
            ArrayList arrWpara = null;
            ArrayList arrpara = null;
            //var option = new TransactionOptions();
            try
            {
                if (oDestination_Dir.Database_Storage)
                {
                    GC.Collect();
                    //option.IsolationLevel = System.Transactions.IsolationLevel.ReadUncommitted;
                    //option.Timeout = TimeSpan.FromMinutes(5);
                    //using (var scope = new TransactionScope(TransactionScopeOption.RequiresNew, option))
                    //{
                    using (_objDal = new ClsBuildQuery(_szAppXmlPath))
                    {
                        if (!_objDal.OpenConnection(ProductName.DocsExecutive, _Location))
                            throw new Exception("Error occured while opening File System DB Connection. Error : " + _objDal.msgError);

                        foreach (File_Data oDoc in lstFile_Data)
                        {
                            #region ..... Delete Existing Data ....

                            arrWpara = new ArrayList();
                            arrpara = new ArrayList();



                            if (oDestination_Dir.Directory_Path.Contains("Template"))
                                arrWpara.Add("on_rs = " + oDoc.Serial_Number + " and yek_rrus = " + oDoc.SurrKey);
                            else
                                arrWpara.Add("on_rs = " + oDoc.Serial_Number + " and on_frc = " + oDoc.SurrKey);
                            arrWpara.Add("");
                            arrWpara.Add("");

                            if (!_objDal.Delete(oDestination_Dir.Table_Name, arrWpara))
                                throw new Exception("Error occured while Deleting Previous blob Data : " + _objDal.msgError);

                            #endregion

                            #region ..... Insert new File Data ...

                            arrWpara.Clear();
                            arrpara.Clear();



                            if (oDestination_Dir.Directory_Path.Contains("Template"))
                            {
                                arrpara.Add("yek_rrus");
                                arrpara.Add("numeric");
                                arrpara.Add(oDoc.SurrKey);
                            }
                            else
                            {
                                arrpara.Add("on_frc");
                                arrpara.Add("numeric");
                                arrpara.Add(oDoc.SurrKey);
                            }

                            arrpara.Add("on_rs");
                            arrpara.Add("numeric");
                            arrpara.Add(oDoc.Serial_Number);

                            arrpara.Add("epyt_rid");
                            arrpara.Add("varchar");
                            arrpara.Add("PV");

                            arrpara.Add("eman_elif");
                            arrpara.Add("varchar");
                            arrpara.Add(oDoc.File_Name);


                            arrpara.Add("eno_5dm");
                            arrpara.Add("varchar");
                            arrpara.Add(oDoc.CheckSum);


                            arrpara.Add("owt_5dm");
                            arrpara.Add("varchar");
                            arrpara.Add(oDoc.Source_File_CheckSum);


                            arrpara.Add("atad_cod");
                            arrpara.Add("Blob");
                            arrpara.Add(oDoc.Data);

                            arrpara.Add("yb_detaerc");
                            arrpara.Add("varchar");
                            arrpara.Add(oDoc.User_Id);

                            arrpara.Add("no_detaerc");
                            arrpara.Add("varchar");
                            arrpara.Add(DateTime.Now.ToString("MM/dd/yyyy"));

                            if (!_objDal.Insert(oDestination_Dir.Table_Name, arrpara))
                                throw new Exception("Error occured While inserting Blob Data :" + _objDal.msgError);

                            #endregion

                        }

                        _objDal.CloseConnection();
                    }
                    //    scope.Complete();
                    //}
                }
            }
            finally
            {
                //if (_bIsBeginTransaction)
                //    _objDal.RollBackTransaction();
                if (arrWpara != null)
                    arrWpara.Clear();
                arrWpara = null;
                if (arrpara != null)
                    arrpara.Clear();
                arrpara = null;

                if (_objDal != null)
                {
                    _objDal.CloseConnection();
                    _objDal.Dispose();
                }
                _objDal = null;
                GC.Collect();
            }
            return lstFile_Data;
        }

        public File_Data Save_File_In_Database(Directory_Attributes oDestination_Dir, File_Data oFile_Data)
        {
            ArrayList arrWpara = null;
            ArrayList arrpara = null;
            var option = new TransactionOptions();
            string szWCondition = string.Empty;
            try
            {
                if (oDestination_Dir.Database_Storage)
                {
                    File.AppendAllText(_szLogFileName, "Start " + oFile_Data.SurrKey + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);
                    GC.Collect();
                    option.IsolationLevel = System.Transactions.IsolationLevel.ReadUncommitted;
                    option.Timeout = TimeSpan.FromMinutes(5);
                    using (var scope = new TransactionScope(TransactionScopeOption.RequiresNew, option))
                    {
                        File.AppendAllText(_szLogFileName, " DAL Object Creation " + oFile_Data.SurrKey + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);
                        using (_objDal = new ClsBuildQuery(_szAppXmlPath))
                        {
                            File.AppendAllText(_szLogFileName, " Open Connection " + oFile_Data.SurrKey + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);
                            if (!_objDal.OpenConnection(ProductName.DocsExecutive, _Location))
                                throw new Exception("Error occured while opening File System DB Connection. Error : " + _objDal.msgError);

                            File.AppendAllText(_szLogFileName, " Opened Connection " + oFile_Data.SurrKey + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);

                            arrWpara = new ArrayList();
                            arrpara = new ArrayList();
                            File.AppendAllText(_szLogFileName, " Prepare SQL Query " + oFile_Data.SurrKey + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);
                            switch (oFile_Data.Destination_Directory)
                            {
                                case "TW":
                                case "TE":
                                case "TO":
                                    szWCondition = "on_rs = " + oFile_Data.Serial_Number + " and yek_rrus = " + oFile_Data.SurrKey;
                                    break;
                                case "DV":
                                    szWCondition = "on_rs = " + oFile_Data.Serial_Number + " and on_frc = " + oFile_Data.SurrKey + " AND no_noisrev=" + oFile_Data.Draft_Version;
                                    break;
                                case "CH":
                                case "CHP":
                                    szWCondition = "on_rs = " + oFile_Data.Serial_Number + " and on_frc = " + oFile_Data.SurrKey + " AND yb_detaerc='" + oFile_Data.User_Id + "' And epyt_resu='" + oFile_Data.Type_of_User + "'";
                                    break;
                                default:
                                    szWCondition = "on_rs = " + oFile_Data.Serial_Number + " and on_frc = " + oFile_Data.SurrKey;
                                    break;
                            }

                            #region ..... Update Delete Indicator ....
                            //..New Logic for Performance improvement (Update/Insert)
                            File.AppendAllText(_szLogFileName, " Update Delete " + oFile_Data.SurrKey + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);
                            //if (!_objDal.Delete(oDestination_Dir.Table_Name, arrWpara))
                            //    throw new Exception("Error occured while Deleting Previous blob Data : " + _objDal.msgError);

                            _szSqlQuery = "Update " + oDestination_Dir.Table_Name + " SET eteled=1 where " + szWCondition;
                            if (!_objDal.ExecuteQuery(_szSqlQuery))
                                throw new Exception("Error occured while Updating Record. Error : " + _objDal.msgError);


                            File.AppendAllText(_szLogFileName, " Update Completed " + oFile_Data.SurrKey + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);

                            #endregion

                            #region ..... Insert new File Data ...

                            arrWpara.Clear();
                            arrpara.Clear();
                            File.AppendAllText(_szLogFileName, " Prepare insert Query " + oFile_Data.SurrKey + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);
                            switch (oFile_Data.Destination_Directory)
                            {
                                case "TW":
                                case "TE":
                                case "TO":
                                    arrpara.Add("yek_rrus");
                                    arrpara.Add("numeric");
                                    arrpara.Add(oFile_Data.SurrKey);
                                    File.AppendAllText(_szLogFileName, " oFile_Data.SurrKey : " + oFile_Data.SurrKey + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);
                                    break;
                                case "DV":
                                    arrpara.Add("on_frc");
                                    arrpara.Add("numeric");
                                    arrpara.Add(oFile_Data.SurrKey);


                                    arrpara.Add("no_noisrev");
                                    arrpara.Add("numeric");
                                    arrpara.Add(oFile_Data.Draft_Version);


                                    break;
                                case "CH":
                                case "CHP":
                                    arrpara.Add("on_frc");
                                    arrpara.Add("numeric");
                                    arrpara.Add(oFile_Data.SurrKey);

                                    arrpara.Add("epyt_resu");
                                    arrpara.Add("varchar");
                                    arrpara.Add(oFile_Data.Type_of_User);

                                    break;
                                default:
                                    arrpara.Add("on_frc");
                                    arrpara.Add("numeric");
                                    arrpara.Add(oFile_Data.SurrKey);
                                    break;

                            }

                            arrpara.Add("on_rs");
                            arrpara.Add("numeric");
                            arrpara.Add(oFile_Data.Serial_Number);
                            File.AppendAllText(_szLogFileName, " oFile_Data.Serial_Number : " + oFile_Data.Serial_Number + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);



                            arrpara.Add("eman_elif");
                            arrpara.Add("varchar");
                            arrpara.Add(oFile_Data.Destination_File_Name);
                            File.AppendAllText(_szLogFileName, " oFile_Data.Destination_File_Name : " + oFile_Data.Destination_File_Name + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);


                            arrpara.Add("eno_5dm");
                            arrpara.Add("varchar");
                            arrpara.Add(oFile_Data.CheckSum);
                            File.AppendAllText(_szLogFileName, " oFile_Data.CheckSum : " + oFile_Data.CheckSum + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);


                            arrpara.Add("owt_5dm");
                            arrpara.Add("varchar");
                            arrpara.Add(oFile_Data.Source_File_CheckSum);
                            File.AppendAllText(_szLogFileName, " oFile_Data.Source_File_CheckSum : " + oFile_Data.Source_File_CheckSum + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);


                            arrpara.Add("atad_cod");
                            arrpara.Add("Blob");
                            arrpara.Add(oFile_Data.Data);
                            File.AppendAllText(_szLogFileName, " oFile_Data.Data : " + oFile_Data.Data.ToString() + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);

                            arrpara.Add("yb_detaerc");
                            arrpara.Add("varchar");
                            arrpara.Add(oFile_Data.User_Id);
                            File.AppendAllText(_szLogFileName, " oFile_Data.User_Id : " + oFile_Data.User_Id.ToString() + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);

                            arrpara.Add("no_detaerc");
                            arrpara.Add("varchar");
                            arrpara.Add(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss"));

                            File.AppendAllText(_szLogFileName, " Prepared insert Query " + oFile_Data.SurrKey + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);

                            File.AppendAllText(_szLogFileName, " Start insert Query " + oFile_Data.SurrKey + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);
                            File.AppendAllText(_szLogFileName, " Table Name : " + oDestination_Dir.Table_Name + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);
                            if (!_objDal.Insert(oDestination_Dir.Table_Name, arrpara))
                                throw new Exception("Error occured While inserting Blob Data :" + _objDal.msgError);

                            File.AppendAllText(_szLogFileName, " END insert Query " + oFile_Data.SurrKey + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);


                            //_objDal.CommitTransaction();
                            //_bIsBeginTransaction = false;
                            #endregion

                            _objDal.CloseConnection();
                        }
                        scope.Complete();
                    }
                }
            }
            catch (Exception ex)
            {
                File.AppendAllText(_szLogFileName, " Error :  " + ex.Message + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);
            }
            finally
            {
                //if (_bIsBeginTransaction)
                //    _objDal.RollBackTransaction();

                if (arrWpara != null)
                    arrWpara.Clear();
                arrWpara = null;
                if (arrpara != null)
                    arrpara.Clear();
                arrpara = null;

                if (_objDal != null)
                {
                    _objDal.CloseConnection();
                    _objDal.Dispose();
                }
                _objDal = null;
                GC.Collect();
            }
            return oFile_Data;
        }

        public File_Data Get_File_from_Database(Directory_Attributes oSource_Dir, File_Data oFile_Data)
        {
            IDataReader _objDrReader = null;
            string szWCondition = string.Empty;
            var option = new TransactionOptions();
            try
            {
                if (oSource_Dir.Database_Storage)
                {
                    GC.Collect();
                    option.IsolationLevel = System.Transactions.IsolationLevel.ReadUncommitted;
                    option.Timeout = TimeSpan.FromMinutes(5);
                    using (var scope = new TransactionScope(TransactionScopeOption.RequiresNew, option))
                    {
                        try
                        {
                            File.AppendAllText(_szLogFileName, "Get_File_from_Database " + oFile_Data.SurrKey + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);

                            using (_objDal = new ClsBuildQuery(_szAppXmlPath))
                            {
                                File.AppendAllText(_szLogFileName, "OpenConnection " + oFile_Data.SurrKey + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);

                                if (!_objDal.OpenConnection(ProductName.DocsExecutive, _Location))
                                    throw new Exception("Error occured while opening File System DB Connection. Error : " + _objDal.msgError);

                                File.AppendAllText(_szLogFileName, "Opened Connection " + oFile_Data.SurrKey + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);

                                #region ..... new File Data ...


                                File.AppendAllText(_szLogFileName, "Select Record " + oFile_Data.SurrKey + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);

                                switch (oFile_Data.Source_Directory)
                                {
                                    case "TW":
                                    case "TE":
                                    case "TO":
                                        szWCondition = "on_rs = " + oFile_Data.Serial_Number + " and yek_rrus = " + oFile_Data.SurrKey + " AND eteled=0";
                                        break;
                                    case "DV":
                                        szWCondition = "on_rs = " + oFile_Data.Serial_Number + " and on_frc = " + oFile_Data.SurrKey + " AND no_noisrev=" + oFile_Data.Draft_Version + " AND eteled=0";
                                        break;
                                    case "CH":
                                    case "CHP":
                                        szWCondition = "on_rs = " + oFile_Data.Serial_Number + " and on_frc = " + oFile_Data.SurrKey + " AND yb_detaerc='" + oFile_Data.User_Id + "' And epyt_resu='" + oFile_Data.Type_of_User + "' AND eteled=0";
                                        break;
                                    default:
                                        szWCondition = "on_rs = " + oFile_Data.Serial_Number + " and on_frc = " + oFile_Data.SurrKey + " AND eteled=0";
                                        break;
                                }




                                switch (oFile_Data.Source_Directory)
                                {
                                    case "TW":
                                    case "TE":
                                    case "TO":
                                        _szSqlQuery = "select atad_cod,eno_5dm,eman_elif,owt_5dm from " + oSource_Dir.Table_Name + " where " + szWCondition;
                                        break;
                                    default:
                                        _szSqlQuery = "select atad_cod,eno_5dm,eman_elif,owt_5dm from " + oSource_Dir.Table_Name + " where " + szWCondition;
                                        break;
                                }
                                File.AppendAllText(_szLogFileName, " IsRecordExist " + oFile_Data.SurrKey + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);

                                if (!_objDal.IsRecordExist(_szSqlQuery))
                                    throw new Exception("File Not Found in Source Directory : " + oSource_Dir.Directory_Path);

                                File.AppendAllText(_szLogFileName, " Get Data " + oFile_Data.SurrKey + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);

                                _objDrReader = _objDal.DecideDatabaseQDR(_szSqlQuery);
                                if (_objDal.msgError != "")
                                    throw new Exception(_objDal.msgError);

                                if (_objDrReader.Read())
                                {
                                    oFile_Data.Data = (byte[])_objDrReader["atad_cod"];
                                    oFile_Data.CheckSum = Convert.ToString(_objDrReader["eno_5dm"]);
                                    oFile_Data.Source_File_CheckSum = Convert.ToString(_objDrReader["owt_5dm"]);
                                    oFile_Data.File_Name = Convert.ToString(_objDrReader["eman_elif"]);
                                }

                                _objDrReader.Close();
                                _objDrReader.Dispose();
                                _objDrReader = null;
                                File.AppendAllText(_szLogFileName, " Data Retrived Successfully " + oFile_Data.SurrKey + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);



                                #endregion

                                _objDal.CloseConnection();
                            }
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                        finally
                        {
                            if (_objDrReader != null)
                            {
                                _objDrReader.Close();
                                _objDrReader.Dispose();
                            }
                            _objDrReader = null;
                            if (_objDal != null)
                            {
                                _objDal.CloseConnection();
                                _objDal.Dispose();
                            }
                            _objDal = null;
                            GC.Collect();
                        }
                        scope.Complete();
                    }
                }
            }
            finally
            {
                if (_objDrReader != null)
                {
                    _objDrReader.Close();
                    _objDrReader.Dispose();
                }
                _objDrReader = null;
                if (_objDal != null)
                {
                    _objDal.CloseConnection();
                    _objDal.Dispose();
                }
                _objDal = null;
            }
            return oFile_Data;
        }

        public File_Data Get_File_Checksum_From_Database(Directory_Attributes oSource_Dir, File_Data oFile_Data)
        {
            object _objReturnVal = null;
            string szWCondition = string.Empty;
            try
            {
                if (oSource_Dir.Database_Storage)
                {
                    GC.Collect();
                    //option.IsolationLevel = System.Transactions.IsolationLevel.ReadUncommitted;
                    //option.Timeout = TimeSpan.FromMinutes(5);
                    //using (var scope = new TransactionScope(TransactionScopeOption.RequiresNew, option))
                    //{
                    try
                    {
                        using (_objDal = new ClsBuildQuery(_szAppXmlPath))
                        {
                            if (!_objDal.OpenConnection(ProductName.DocsExecutive, _Location))
                                throw new Exception("Error occured while opening File System DB Connection. Error : " + _objDal.msgError);

                            #region ..... new File Data ...


                            switch (oFile_Data.Source_Directory)
                            {
                                case "TW":
                                case "TE":
                                case "TO":
                                    szWCondition = "on_rs = " + oFile_Data.Serial_Number + " and yek_rrus = " + oFile_Data.SurrKey + " AND eteled=0";
                                    break;
                                case "DV":
                                    szWCondition = "on_rs = " + oFile_Data.Serial_Number + " and on_frc = " + oFile_Data.SurrKey + " AND no_noisrev=" + oFile_Data.Draft_Version + " AND eteled=0";
                                    break;
                                case "CH":
                                case "CHP":
                                    szWCondition = "on_rs = " + oFile_Data.Serial_Number + " and on_frc = " + oFile_Data.SurrKey + " AND yb_detaerc='" + oFile_Data.User_Id + "' And epyt_resu='" + oFile_Data.Type_of_User + "' AND eteled=0";
                                    break;
                                default:
                                    szWCondition = "on_rs = " + oFile_Data.Serial_Number + " and on_frc = " + oFile_Data.SurrKey + " AND eteled=0";
                                    break;
                            }


                            switch (oFile_Data.Source_Directory)
                            {
                                case "TW":
                                case "TE":
                                case "TO":
                                    _szSqlQuery = "select eno_5dm from " + oSource_Dir.Table_Name + " where " + szWCondition;
                                    break;
                                default:
                                    _szSqlQuery = "select eno_5dm from " + oSource_Dir.Table_Name + " where " + szWCondition;
                                    break;
                            }

                            _objReturnVal = _objDal.GetFirstColumnValue(_szSqlQuery);
                            if (_objDal.msgError != "")
                                throw new Exception(_objDal.msgError);

                            if (_objReturnVal != null)
                                oFile_Data.CheckSum = Convert.ToString(_objReturnVal);
                            _objReturnVal = null;


                            #endregion

                            _objDal.CloseConnection();
                        }
                    }
                    catch (Exception ex)
                    {
                        //File.AppendAllTextFile.AppendAllText(_szLogFileName, DateTime.Now + " Error :" + ex.Message + Environment.NewLine);
                        //File.AppendAllTextFile.AppendAllText(_szLogFileName, DateTime.Now + " Error :" + ex.StackTrace + Environment.NewLine);
                        throw ex;
                    }
                    finally
                    {
                        _objReturnVal = null;
                        if (_objDal != null)
                        {
                            _objDal.CloseConnection();
                            _objDal.Dispose();
                        }
                        _objDal = null;
                        GC.Collect();
                        //}
                        //scope.Complete();
                    }
                }
            }
            finally
            {
                _objReturnVal = null;
                if (_objDal != null)
                {
                    _objDal.CloseConnection();
                    _objDal.Dispose();
                }
                _objDal = null;
            }
            return oFile_Data;
        }

        public List<File_Data> Get_Documents(Directory_Attributes oSource_Dir, File_Data oFile_Data)
        {
            IDataReader _objDrReader = null;
            string szWCondition = string.Empty;
            List<File_Data> lstFile_Data = new List<File_Data>();
            try
            {
                if (oSource_Dir.Database_Storage)
                {

                    GC.Collect();
                    File.AppendAllText(_szLogFileName, " Get_Documents " + oFile_Data.SurrKey + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);
                    using (_objDal = new ClsBuildQuery(_szAppXmlPath))
                    {
                        File.AppendAllText(_szLogFileName, " OpenConnection " + oFile_Data.SurrKey + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);
                        if (!_objDal.OpenConnection(ProductName.DocsExecutive, _Location))
                            throw new Exception("Error occured while opening File System DB Connection. Error : " + _objDal.msgError);

                        File.AppendAllText(_szLogFileName, " Opened Connection " + oFile_Data.SurrKey + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);

                        #region ..... new File Data ...


                        File.AppendAllText(_szLogFileName, " Get Data " + oFile_Data.SurrKey + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);
                        switch (oFile_Data.Source_Directory)
                        {
                            case "TW":
                            case "TE":
                            case "TO":
                                szWCondition = "yek_rrus = " + oFile_Data.SurrKey + " AND eteled=0";
                                break;
                            case "DV":
                                if (oFile_Data.Draft_Version != 0)
                                    szWCondition = "on_frc = " + oFile_Data.SurrKey + " AND no_noisrev=" + oFile_Data.Draft_Version + " AND eteled=0";
                                else
                                    szWCondition = "on_frc = " + oFile_Data.SurrKey + " AND eteled=0";
                                break;
                            case "CH":
                            case "CHP":
                                szWCondition = " on_frc = " + oFile_Data.SurrKey;

                                if (!string.IsNullOrEmpty(oFile_Data.User_Id))
                                    szWCondition = szWCondition + " AND yb_detaerc='" + oFile_Data.User_Id + "' AND eteled=0";
                                if (!string.IsNullOrEmpty(oFile_Data.Type_of_User))
                                    szWCondition = szWCondition + " And epyt_resu='" + oFile_Data.Type_of_User + "' AND eteled=0";
                                break;
                            default:
                                szWCondition = " on_frc = " + oFile_Data.SurrKey + " AND eteled=0";
                                break;
                        }


                        switch (oFile_Data.Source_Directory)
                        {
                            case "TW":
                            case "TE":
                            case "TO":
                                _szSqlQuery = "select * from " + oSource_Dir.Table_Name + " where " + szWCondition;
                                break;
                            default:
                                _szSqlQuery = "select * from " + oSource_Dir.Table_Name + " where " + szWCondition;
                                break;
                        }

                        _objDrReader = _objDal.DecideDatabaseQDR(_szSqlQuery);
                        if (_objDal.msgError != "")
                            throw new Exception(_objDal.msgError);

                        while (_objDrReader.Read())
                        {
                            File_Data oFile = new File_Data();
                            oFile.Data = (byte[])_objDrReader["atad_cod"];
                            oFile.Serial_Number = Convert.ToInt32(_objDrReader["on_rs"]);
                            oFile.CheckSum = Convert.ToString(_objDrReader["eno_5dm"]);
                            oFile.Source_File_CheckSum = Convert.ToString(_objDrReader["owt_5dm"]);
                            oFile.File_Name = Convert.ToString(_objDrReader["eman_elif"]);
                            oFile.User_Id = Convert.ToString(_objDrReader["yb_detaerc"]);
                            if ((_objDrReader.GetSchemaTable().Select("ColumnName = 'epyt_resu'").Count() == 1))
                                oFile.Type_of_User = Convert.ToString(_objDrReader["epyt_resu"]);

                            lstFile_Data.Add(oFile);
                            oFile_Data = null;
                        }

                        _objDrReader.Close();
                        _objDrReader.Dispose();
                        _objDrReader = null;
                        File.AppendAllText(_szLogFileName, " Retrived Data " + oFile_Data.SurrKey + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + Environment.NewLine);


                        #endregion

                    }
                }
            }
            finally
            {
                if (_objDrReader != null)
                {
                    _objDrReader.Close();
                    _objDrReader.Dispose();
                }
                _objDrReader = null;
                if (_objDal != null)
                {
                    _objDal.CloseConnection();
                    _objDal.Dispose();
                }
                _objDal = null;
                GC.Collect();
            }
            return lstFile_Data;
        }


        public bool isRecord_Exist_In_Database(Directory_Attributes oSource_Dir, File_Data oFile_Data)
        {
            bool bResult = false;
            string szWCondition = string.Empty;
            try
            {
                if (oSource_Dir.Database_Storage)
                {
                    using (_objDal = new ClsBuildQuery(_szAppXmlPath))
                    {
                        if (!_objDal.OpenConnection(ProductName.DocsExecutive, _Location))
                            throw new Exception("Error occured while opening File System DB Connection. Error : " + _objDal.msgError);

                        #region ..... Check File ...

                        switch (oFile_Data.Source_Directory)
                        {
                            case "TW":
                            case "TE":
                            case "TO":
                                szWCondition = "on_rs = " + oFile_Data.Serial_Number + " and yek_rrus = " + oFile_Data.SurrKey + " AND eteled=0";
                                break;
                            case "DV":
                                szWCondition = "on_rs = " + oFile_Data.Serial_Number + " and on_frc = " + oFile_Data.SurrKey + " AND no_noisrev=" + oFile_Data.Draft_Version + " AND eteled=0";
                                break;
                            case "CH":
                            case "CHP":
                                szWCondition = "on_rs = " + oFile_Data.Serial_Number + " and on_frc = " + oFile_Data.SurrKey + " AND yb_detaerc='" + oFile_Data.User_Id + "' And epyt_resu='" + oFile_Data.Type_of_User + "' AND eteled=0";
                                break;
                            default:
                                szWCondition = "on_rs = " + oFile_Data.Serial_Number + " and on_frc = " + oFile_Data.SurrKey + " AND eteled=0";
                                break;
                        }

                        switch (oFile_Data.Source_Directory)
                        {
                            case "TW":
                            case "TE":
                            case "TO":
                                _szSqlQuery = "select eman_elif from " + oSource_Dir.Table_Name + " where " + szWCondition;
                                break;
                            default:
                                _szSqlQuery = "select eman_elif from " + oSource_Dir.Table_Name + " where " + szWCondition;
                                break;

                        }
                        bResult = _objDal.IsRecordExist(_szSqlQuery);
                        if (_objDal.msgError != "")
                            throw new Exception(_objDal.msgError);

                        #endregion
                    }
                }
            }
            finally
            {
                if (_objDal != null)
                {
                    _objDal.CloseConnection();
                    _objDal.Dispose();
                }
                _objDal = null;

            } return bResult;
        }

        public bool Delete_File_From_Database(Directory_Attributes oDestination_Dir, File_Data oFile_Data)
        {
            string szWCondition = string.Empty;
            try
            {
                if (oDestination_Dir.Database_Storage)
                {
                    using (_objDal = new ClsBuildQuery(_szAppXmlPath))
                    {
                        if (!_objDal.OpenConnection(ProductName.DocsExecutive, _Location))
                            throw new Exception("Error occured while opening File System DB Connection. Error : " + _objDal.msgError);


                        #region ..... Delete Existing Data ....

                        switch (oFile_Data.Source_Directory)
                        {
                            case "TW":
                            case "TE":
                            case "TO":
                                szWCondition = "on_rs = " + oFile_Data.Serial_Number + " and yek_rrus = " + oFile_Data.SurrKey;
                                break;
                            case "DV":
                                //szWCondition = "on_rs = " + oFile_Data.Serial_Number + " and on_frc = " + oFile_Data.SurrKey + " AND no_noisrev=" + oFile_Data.Draft_Version;
                                if (oFile_Data.Draft_Version != 0)
                                    szWCondition = "on_frc = " + oFile_Data.SurrKey + " AND no_noisrev=" + oFile_Data.Draft_Version;
                                else
                                    szWCondition = "on_frc = " + oFile_Data.SurrKey;
                                break;
                            case "CH":
                            case "DF":
                                szWCondition = "on_frc = " + oFile_Data.SurrKey;
                                break;
                            default:
                                szWCondition = "on_rs = " + oFile_Data.Serial_Number + " and on_frc = " + oFile_Data.SurrKey;
                                break;
                        }

                        switch (oFile_Data.Source_Directory)
                        {
                            case "TW":
                            case "TE":
                            case "TO":
                                _szSqlQuery = "Update " + oDestination_Dir.Table_Name + " SET eteled=1 where " + szWCondition;
                                break;
                            default:
                                _szSqlQuery = "Update " + oDestination_Dir.Table_Name + " SET eteled=1 where " + szWCondition;
                                break;

                        }
                        ;
                        if (!_objDal.ExecuteQuery(_szSqlQuery))
                            throw new Exception(_objDal.msgError);

                        #endregion

                    }
                }
            }
            finally
            {
                if (_objDal != null)
                {
                    _objDal.CloseConnection();
                    _objDal.Dispose();
                }
                _objDal = null;
            }
            return true;
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
                if (_objDal != null)
                {
                    _objDal.CloseConnection();
                    _objDal.Dispose();
                }
                _objDal = null;
                _szAppXmlPath = string.Empty;
                _szDBName = string.Empty;
                _Location = string.Empty;
                _szSqlQuery = string.Empty;

            }
            else
            {

            }
        }

        ~ClsSave_File_in_Database()
        {
            Dispose(false);
        }


        #endregion

    }
}
