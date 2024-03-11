//... Code Changed By Manavya on 22/07/2011 For DRT-2783 ...
//... Code Changed By Manavya on 04/12/2012 For DRT-4136 ...
using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.Web;
using System.IO;
using System.Collections;
using System.ComponentModel;

namespace DDLLCS
{
    # region .... Global Variable Declaration ....

    # region .... Enumeration ....

    public enum ProductName
    {
        /// <summary>
        /// Connection For Docs-Executive Database. (For: DocsExecutive Web Application)
        /// </summary>
        DocsExecutive = 0,
        /// <summary>
        /// Connection For Docs-Executive Database.
        /// </summary>
        DocsExecutive_Backend = 1,
        /// <summary>
        /// Connection For Docs-Dossier Bridge Database.
        /// </summary>
        DocsDossierBridge,
        DossierMgmt,
        DrugListing,
        RsrchExecutive
    }

    public enum ConnectionFor
    {
        /// <summary>
        /// Connection For Docs-Executive Database. (For: DocsExecutive Web Application)
        /// </summary>
        DocsExecutive = 0,
        /// <summary>
        /// Connection For Docs-Executive Database.
        /// </summary>
        DocsExecutive_Backend = 1,
        /// <summary>
        /// Connection For Docs-Dossier Bridge Database.
        /// </summary>
        DocsDossierBridge,
        DossierMgmt,
        DrugListing,
        RsrchExecutive
    }

    public enum ConnectionType
    {
        /// <summary>
        /// Default: Connection Open & use by all Users. (Connection is Always Open. Need not to call CloseConnection() method).
        /// </summary>
        Global = 0,
        /// <summary>
        /// Temporary: For Transaction handling or short period (Connection is Open for short period of time so Always call CloseConnection() method after Use).
        /// </summary>
        NewConnection
    }

    #endregion

    #endregion

    /// <summary> Interface used in Dal
    /// 1]IDbConnection interface:- The IDbConnection interface enables an inheriting
    /// class to implement a Connection class, which represents a unique session with a data source 
    /// 2]IDbTransaction interface:- The IDbTransaction interface allows an inheriting
    /// class to implement a Transaction class, which represents the transaction to be performed at a data source
    /// 3]IDbCommand interface:- The IDbCommand interface enables an inheriting class to
    /// implement a Command class, which represents an SQL statement that is executed at a data source.
    /// 4]IDbDataAdapter interface:- The IDbDataAdapter interface inherits from the 
    /// IDataAdapter interface and allows an object to create a DataAdapter designed
    /// for use with a relational database. The IDbDataAdapter interface and,optionally,
    /// the utility class, DbDataAdapter, allow an inheriting class to implement a DataAdapter 
    /// class, which represents the bridge between a data source and a DataSet.
    /// 5]IDataReader interface :- The IDataReader and IDataRecord interfaces allow
    /// an inheriting class to implement a DataReader class, which provides a means
    /// of reading one or more forward-only streams of result sets.
    /// </summary>

    public class ClsExecuteQuery : IDisposable
    {
        # region .... Variable Declaration ....

        // object for reading Application.xml file
        //private eDocsDN_ReadAppXml.clsReadAppXml objReadXml;
        private GetConInfo.ClsConInfo objGetConInfo = null;
        private IDbConnection idbConnection = null;
        private IDbTransaction idbTransaction = null;
        private IDbCommand idbCommand;
        private IDbDataAdapter idbDtAdapter;
        private IDataReader objDtReader;
        private DataSet objDataset;
        private DataTable objDtTable;
        private Uri _uriConnAppSite;

        private ConnectionFor _eConnectionFor;
        private ConnectionType _eConnectionType;
        private bool _bResult;
        int _iCommandTimeout;
        protected string _szDbName;
        private string _szAppXmlPath;
        private string _szProductName;
        private string _szRegion;
        private string _szConnectionName;
        private string _szError;

        # region .... Enumeration ....

        //public enum ConnectionFor
        //{
        //    /// <summary>
        //    /// Extra Connection For Transaction Handling. (For Docs-Executive Database)
        //    /// </summary>
        //    Transaction = 1,
        //    /// <summary>
        //    /// Connection For Docs-Executive Database. (For: DocsExecutive Web Application)
        //    /// </summary>
        //    DocsExecutive,
        //    /// <summary>
        //    /// Connection For Docs-Executive Database.
        //    /// </summary>
        //    DocsExecutive_Backend,
        //}

        #endregion

        #endregion

        #region .... Constructor ....

        public ClsExecuteQuery(IntPtr iPtrHandle)
        {
            msgError = "";
            this.handle = iPtrHandle;
            _iCommandTimeout = 30;
        }

        #endregion

        # region .... Database, msgError And AppXmlPath and DataBaseConnStr Property ....

        /// <summary> Properties Used In Class
        /// 1] Database :- This property used for getting Database name 
        /// 2] AppXmlPath:- This property uesd for getting Application.xml path
        /// 3] msgError :- This property used for setting database error.
        ///    This property used in Select_DataReader and Select_Dataset function 
        /// </summary>

        protected internal Uri ConnAppSiteUri
        {
            get { return _uriConnAppSite; }
            set { _uriConnAppSite = value; }
        }

        /// <summary>
        /// CommandTimeout in seconds. (Default Value is 30 sec.)
        /// </summary>
        public int CommandTimeout
        {
            get { return _iCommandTimeout; }
            set { _iCommandTimeout = value; }
        }

        protected internal string DataBase
        {
            get { return _szDbName; }
            set { _szDbName = value; }
        }

        protected internal string AppXmlPath
        {
            get { return _szAppXmlPath; }
            set { _szAppXmlPath = value; }
        }

        protected internal string Product_Name
        {
            get { return _szProductName; }
            set { _szProductName = value; }
        }

        protected internal string Region
        {
            get { return _szRegion; }
            set { _szRegion = value; }
        }

        protected internal string ConnectionName
        {
            get { return _szConnectionName; }
            set { _szConnectionName = value; }
        }

        public string msgError
        {
            get { return _szError; }
            set { _szError = value; }
        }

        #endregion

        # region .... GetCommand and GetDataAdapter ....

        /// <summary>
        /// This function return Command object depend upon database
        /// </summary>
        /// <returns>IDbCommand</returns>
        private IDbCommand GetCommand()
        {
            try
            {
                switch (DataBase)
                {
                    case "MSSQL":
                    case "ORACLE":
                        {
                            idbCommand = new OleDbCommand();
                            idbCommand.CommandTimeout = _iCommandTimeout;
                        }
                        break;
                    default:
                        {
                            idbCommand = new OleDbCommand();
                            idbCommand.CommandTimeout = _iCommandTimeout;
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            return idbCommand;
        }

        /// <summary>
        /// This function return DataAdapter depend upon database
        /// </summary>
        /// <returns>IDbDataAdapter</returns>
        private IDbDataAdapter GetDataAdapter()
        {
            try
            {
                switch (DataBase)
                {
                    case "MSSQL":
                    case "ORACLE":
                        idbDtAdapter = new OleDbDataAdapter();
                        break;
                    default:
                        idbDtAdapter = new OleDbDataAdapter();
                        break;
                }
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            return idbDtAdapter;
        }

        # endregion

        #region .... OpenConnection And CloseConnection ....

        private bool OpenDatabaseConnection()
        {
            msgError = "";
            _bResult = true;
            try
            {
                #region ... Open Connection ...

                if (_eConnectionType == ConnectionType.NewConnection)
                {
                    #region ... For Transaction OR OtherDatabse ...

                    switch (DataBase)
                    {
                        case "MSSQL":
                        case "ORACLE":
                            idbConnection = new OleDbConnection();
                            break;
                        default:
                            idbConnection = new OleDbConnection();
                            break;
                    }

                    if (idbConnection.State != ConnectionState.Open)
                    {
                        idbConnection.ConnectionString = GetConnectionString();
                        idbConnection.Open();
                    }

                    #endregion
                }
                else
                {
                    switch (_eConnectionFor)
                    {
                        case ConnectionFor.DocsExecutive:
                            {
                                #region ... DocsExecutive Common (Web) Connection ...

                                if (HttpContext.Current.Application["DOCS_CC"] != null)
                                {
                                    idbConnection = (OleDbConnection)Convert.ChangeType(HttpContext.Current.Application["DOCS_CC"], typeof(OleDbConnection));
                                    if (idbConnection.State != ConnectionState.Open)
                                    {
                                        idbConnection.ConnectionString = GetConnectionString();
                                        idbConnection.Open();
                                        HttpContext.Current.Application["DOCS_CC"] = idbConnection;
                                    }
                                }
                                else
                                {
                                    switch (DataBase)
                                    {
                                        case "MSSQL":
                                        case "ORACLE":
                                            idbConnection = new OleDbConnection();
                                            break;
                                        default:
                                            idbConnection = new OleDbConnection();
                                            break;
                                    }

                                    if (idbConnection.State != ConnectionState.Open)
                                    {
                                        idbConnection.ConnectionString = GetConnectionString();
                                        idbConnection.Open();
                                    }
                                    HttpContext.Current.Application["DOCS_CC"] = idbConnection;
                                }

                                #endregion
                            }
                            break;
                        case ConnectionFor.DocsDossierBridge:
                            {
                                #region ... Docs-Dossier Bridge Connection ...

                                //----------------------------------------------------------------------------------------------//
                                if (HttpContext.Current.Application["BRIDGE_CC"] != null)
                                {
                                    idbConnection = (OleDbConnection)Convert.ChangeType(HttpContext.Current.Application["BRIDGE_CC"], typeof(OleDbConnection));
                                    if (idbConnection.State != ConnectionState.Open)
                                    {
                                        idbConnection.ConnectionString = GetConnectionString();
                                        idbConnection.Open();
                                        HttpContext.Current.Application["BRIDGE_CC"] = idbConnection;
                                    }
                                }
                                else
                                {
                                    switch (DataBase)
                                    {
                                        case "MSSQL":
                                        case "ORACLE":
                                            idbConnection = new OleDbConnection();
                                            break;
                                        default:
                                            idbConnection = new OleDbConnection();
                                            break;
                                    }

                                    if (idbConnection.State != ConnectionState.Open)
                                    {
                                        idbConnection.ConnectionString = GetConnectionString();
                                        idbConnection.Open();
                                    }
                                    HttpContext.Current.Application["BRIDGE_CC"] = idbConnection;
                                }
                                //----------------------------------------------------------------------------------------------//

                                #endregion
                            }
                            break;
                        case ConnectionFor.DocsExecutive_Backend:
                        case ConnectionFor.DossierMgmt:
                        case ConnectionFor.DrugListing:
                        case ConnectionFor.RsrchExecutive:
                            {
                                _eConnectionType = ConnectionType.NewConnection;

                                #region ... For Transaction OR OtherDatabse ...

                                switch (DataBase)
                                {
                                    case "MSSQL":
                                    case "ORACLE":
                                        idbConnection = new OleDbConnection();
                                        break;
                                    default:
                                        idbConnection = new OleDbConnection();
                                        break;
                                }

                                if (idbConnection.State != ConnectionState.Open)
                                {
                                    idbConnection.ConnectionString = GetConnectionString();
                                    idbConnection.Open();
                                }

                                #endregion
                            }
                            break;
                    }
                }

                #endregion
            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = "DAL:- " + ex.Message;
            }
            finally
            {
                if (objGetConInfo != null)
                    objGetConInfo.Dispose();
                objGetConInfo = null;
            }
            return _bResult;
        }

        private string GetConnectionString()
        {
            objGetConInfo = null;
            string szConString = "";
            try
            {
                //----------------------------------------------------------------------------------------------//
                if (ConnAppSiteUri != null)
                    objGetConInfo = new GetConInfo.ClsConInfo(ConnAppSiteUri);
                else
                    objGetConInfo = new GetConInfo.ClsConInfo(AppXmlPath);

                if (objGetConInfo.msgError != "")
                    throw new Exception(objGetConInfo.msgError);

                szConString = objGetConInfo.GetConnectionString(Product_Name, Region, ConnectionName);
                if (objGetConInfo.msgError != "")
                    throw new Exception(objGetConInfo.msgError);

                objGetConInfo.Dispose();
                objGetConInfo = null;

                if (string.IsNullOrEmpty(szConString))
                    throw new Exception("Connection string should not be Null or Empty !!");
                //----------------------------------------------------------------------------------------------//
            }
            finally
            {
                if (objGetConInfo != null)
                {
                    objGetConInfo.Dispose();
                    objGetConInfo = null;
                }
            }
            return szConString;
        }

        /// <summary> Open Connection
        /// In This fucntion 
        ///     1]get the Connection string from Applicatoin.xml 
        ///     2]Depend upon database get connection. 
        ///     3]assing connection string to connection and open this connection 
        ///     4]return true if connection open successfully.
        /// </summary>
        public void OpenConnection()
        {
            //-----------------------------------------------------------//
            msgError = "";
            Product_Name = ProductName.DocsExecutive.ToString();
            Region = "NA";
            ConnectionName = "Common";
            _eConnectionFor = ConnectionFor.DocsExecutive;
            _eConnectionType = ConnectionType.Global;

            OpenDatabaseConnection();
            //-----------------------------------------------------------//
        }

        //public bool OpenConnection()
        //{
        //    //-----------------------------------------------------------//
        //    msgError = "";
        //    Product_Name = ProductName.DocsExecutive.ToString();
        //    Region = "NA";
        //    ConnectionName = "Common";
        //    _eConnectionFor = ConnectionFor.DocsExecutive;
        //    _eConnectionType = ConnectionType.Global;

        //    return OpenDatabaseConnection();
        //    //-----------------------------------------------------------//
        //}

        public void OpenBackEndConnection()
        {
            //-----------------------------------------------------------//
            msgError = "";
            Product_Name = ProductName.DocsExecutive.ToString();
            Region = "NA";
            ConnectionName = "Common";
            _eConnectionFor = ConnectionFor.DocsExecutive;
            _eConnectionType = ConnectionType.NewConnection;

            OpenDatabaseConnection();
            //-----------------------------------------------------------//
        }

        //public bool OpenBackEndConnection()
        //{
        //    //-----------------------------------------------------------//
        //    msgError = "";
        //    Product_Name = ProductName.DocsExecutive.ToString();
        //    Region = "NA";
        //    ConnectionName = "Common";
        //    _eConnectionFor = ConnectionFor.DocsExecutive;
        //    _eConnectionType = ConnectionType.NewConnection;

        //    return OpenDatabaseConnection();
        //    //-----------------------------------------------------------//
        //}

        public void OpenConnection(ConnectionFor eConnectionFor)
        {
            //-----------------------------------------------------------//
            msgError = "";
            ConnectionName = "Common";
            Region = "NA";
            _eConnectionFor = eConnectionFor;
            _eConnectionType = ConnectionType.Global;

            switch (_eConnectionFor)
            {
                case ConnectionFor.DocsExecutive:
                    {
                        Product_Name = ProductName.DocsExecutive.ToString();
                    }
                    break;
                case ConnectionFor.DocsExecutive_Backend:
                    {
                        _eConnectionType = ConnectionType.NewConnection;
                        Product_Name = ProductName.DocsExecutive.ToString();
                    }
                    break;
                case ConnectionFor.DocsDossierBridge:
                    {
                        Product_Name = ProductName.DocsDossierBridge.ToString();
                    }
                    break;
                case ConnectionFor.DossierMgmt:
                    {
                        Product_Name = ProductName.DossierMgmt.ToString();
                        Region = "US";
                    }
                    break;
                case ConnectionFor.DrugListing:
                    {
                        Product_Name = ProductName.DrugListing.ToString();
                    }
                    break;
                case ConnectionFor.RsrchExecutive:
                    {
                        Product_Name = ProductName.RsrchExecutive.ToString();
                    }
                    break;
            }

            OpenDatabaseConnection();
            //-----------------------------------------------------------//
        }

        //public bool OpenConnection(ConnectionFor eConnectionFor)
        //{
        //    //-----------------------------------------------------------//
        //    msgError = "";
        //    ConnectionName = "Common";
        //    Region = "NA";
        //    _eConnectionFor = eConnectionFor;
        //    _eConnectionType = ConnectionType.Global;

        //    switch (_eConnectionFor)
        //    {
        //        case ConnectionFor.DocsExecutive:
        //            {
        //                Product_Name = ProductName.DocsExecutive.ToString();
        //            }
        //            break;
        //        case ConnectionFor.DocsExecutive_Backend:
        //            {
        //                _eConnectionType = ConnectionType.NewConnection;
        //                Product_Name = ProductName.DocsExecutive.ToString();
        //            }
        //            break;
        //        case ConnectionFor.DocsDossierBridge:
        //            {
        //                Product_Name = ProductName.DocsDossierBridge.ToString();
        //            }
        //            break;
        //        case ConnectionFor.DossierMgmt:
        //            {
        //                Product_Name = ProductName.DossierMgmt.ToString();
        //                Region = "US";
        //            }
        //            break;
        //        case ConnectionFor.DrugListing:
        //            {
        //                Product_Name = ProductName.DrugListing.ToString();
        //            }
        //            break;
        //        case ConnectionFor.RsrchExecutive:
        //            {
        //                Product_Name = ProductName.RsrchExecutive.ToString();
        //            }
        //            break;
        //    }

        //    return OpenDatabaseConnection();
        //    //-----------------------------------------------------------//
        //}

        public bool OpenConnection(ConnectionFor eConnectionFor, ConnectionType eConnectionType)
        {
            //-----------------------------------------------------------//
            msgError = "";
            ConnectionName = "Common";
            Region = "NA";
            _eConnectionFor = eConnectionFor;
            _eConnectionType = eConnectionType;

            switch (_eConnectionFor)
            {
                case ConnectionFor.DocsExecutive:
                    {
                        Product_Name = ProductName.DocsExecutive.ToString();
                    }
                    break;
                case ConnectionFor.DocsExecutive_Backend:
                    {
                        _eConnectionType = ConnectionType.NewConnection;
                        Product_Name = ProductName.DocsExecutive.ToString();
                    }
                    break;
                case ConnectionFor.DocsDossierBridge:
                    {
                        Product_Name = ProductName.DocsDossierBridge.ToString();
                    }
                    break;
                case ConnectionFor.DossierMgmt:
                    {
                        Product_Name = ProductName.DossierMgmt.ToString();
                        Region = "US";
                    }
                    break;
                case ConnectionFor.DrugListing:
                    {
                        Product_Name = ProductName.DrugListing.ToString();
                    }
                    break;
                case ConnectionFor.RsrchExecutive:
                    {
                        Product_Name = ProductName.RsrchExecutive.ToString();
                    }
                    break;
            }

            return OpenDatabaseConnection();
            //-----------------------------------------------------------//
        }

        //public void OpenConnection(string szConnStringKeyName)
        //{
        //    _eConnectionFor = ConnectionFor.Transaction;
        //    DataBaseConnKeyName = szConnStringKeyName;
        //    OpenDatabaseConnection();
        //}

        public bool OpenConnection(ProductName eProductName, string szRegion)
        {
            //-----------------------------------------------------------//
            msgError = "";
            Product_Name = eProductName.ToString();
            Region = (string.IsNullOrEmpty(szRegion) || szRegion.Trim() == "" ? "NA" : szRegion);
            ConnectionName = "Common";
            _eConnectionFor = ConnectionFor.DocsExecutive;
            _eConnectionType = ConnectionType.NewConnection;

            return OpenDatabaseConnection();
            //-----------------------------------------------------------//
        }

        public bool OpenConnection(ProductName eProductName, string szRegion, string szConnectionName)
        {
            //-----------------------------------------------------------//
            msgError = "";
            Product_Name = eProductName.ToString();
            Region = (string.IsNullOrEmpty(szRegion) || szRegion.Trim() == "" ? "NA" : szRegion);
            ConnectionName = (string.IsNullOrEmpty(szConnectionName) || szConnectionName.Trim() == "" ? "Common" : szConnectionName);
            _eConnectionFor = ConnectionFor.DocsExecutive;
            _eConnectionType = ConnectionType.NewConnection;

            return OpenDatabaseConnection();
            //-----------------------------------------------------------//
        }

        public bool OpenConnection(string szProductName, string szRegion, string szConnectionName)
        {
            //-----------------------------------------------------------//
            msgError = "";
            Product_Name = szProductName;
            Region = (string.IsNullOrEmpty(szRegion) || szRegion.Trim() == "" ? "NA" : szRegion);
            ConnectionName = (string.IsNullOrEmpty(szConnectionName) || szConnectionName.Trim() == "" ? "Common" : szConnectionName);
            _eConnectionFor = ConnectionFor.DocsExecutive;
            _eConnectionType = ConnectionType.NewConnection;

            return OpenDatabaseConnection();
            //-----------------------------------------------------------//
        }

        /// <summary> Close Connection
        /// Here Conn object is global.
        /// first check connection state and then close the connection
        /// </summary>
        public void CloseConnection()
        {
            try
            {
                if (objGetConInfo != null)
                    objGetConInfo.Dispose();

                if (objDataset != null)
                    objDataset.Dispose();

                if (objDtTable != null)
                    objDtTable.Dispose();

                if (objDtReader != null)
                {
                    objDtReader.Close();
                    objDtReader.Dispose();
                }

                if (idbCommand != null)
                    idbCommand.Dispose();

                //if (idbTransaction != null)
                //    idbTransaction.Dispose();

                switch (_eConnectionType)
                {
                    case ConnectionType.Global:
                        break;
                    case ConnectionType.NewConnection:
                        {
                            //... New, Transaction OR Other Database Connection ...
                            if (idbTransaction != null)
                                idbTransaction.Dispose();
                            idbTransaction = null;

                            if (idbConnection != null)
                            {
                                if (idbConnection.State != ConnectionState.Closed)
                                    idbConnection.Close();
                                idbConnection.Dispose();
                            }
                            idbConnection = null;
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            finally
            {
                objGetConInfo = null;
                objDataset = null;
                objDtTable = null;
                objDtReader = null;

                idbCommand = null;
                idbDtAdapter = null;
                //idbTransaction = null;
                //idbConnection = null;

                //msgError = null;
            }
        }

        //public void CloseConnection(ConnectionFor ConnectionType)
        //{
        //    _eConnectionFor = ConnectionType;
        //    CloseConnection();
        //}

        # endregion

        #region .... Transaction Handlling ....

        /// <summary>
        /// This Function Begin the Transaction with IsolationLevel ReadUncommitted
        /// </summary>
        public void BeginTransaction()
        {
            msgError = "";
            try
            {
                //... IsolationLevel Changed From ReadCommitted to ReadUncommitted ...
                if (this.idbTransaction == null)
                    idbTransaction = idbConnection.BeginTransaction(IsolationLevel.ReadUncommitted);
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
        }

        public void BeginTransaction(IsolationLevel eIsolationLevel)
        {
            msgError = "";
            try
            {
                if (this.idbTransaction == null)
                {
                    switch (eIsolationLevel)
                    {
                        case IsolationLevel.Chaos:
                        case IsolationLevel.ReadCommitted:
                        case IsolationLevel.ReadUncommitted:
                        case IsolationLevel.RepeatableRead:
                        case IsolationLevel.Serializable:
                        case IsolationLevel.Snapshot:
                        case IsolationLevel.Unspecified:
                            {
                                idbTransaction = idbConnection.BeginTransaction(eIsolationLevel);
                            }
                            break;
                        default:
                            {
                                idbTransaction = idbConnection.BeginTransaction(IsolationLevel.ReadUncommitted);
                            }
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
        }

        /// <summary>
        /// Type = ReadCommitted , ReadUncommitted ,  RepeatableRead , Serializable
        /// </summary>
        /// <param name="szIsolationLevelType">ReadCommitted , ReadUncommitted ,  RepeatableRead , Serializable</param>
        public void BeginTransaction(string szIsolationLevelType)
        {
            msgError = "";
            try
            {
                switch (szIsolationLevelType.Trim().ToUpper())
                {
                    case "READCOMMITTED":
                        if (this.idbTransaction == null)
                            idbTransaction = idbConnection.BeginTransaction(IsolationLevel.ReadCommitted);
                        break;
                    case "READUNCOMMITTED":
                        if (this.idbTransaction == null)
                            idbTransaction = idbConnection.BeginTransaction(IsolationLevel.ReadUncommitted);
                        break;
                    case "REPEATABLEREAD":
                        if (this.idbTransaction == null)
                            idbTransaction = idbConnection.BeginTransaction(IsolationLevel.RepeatableRead);
                        break;
                    case "SERIALIZABLE":
                        if (this.idbTransaction == null)
                            idbTransaction = idbConnection.BeginTransaction(IsolationLevel.Serializable);
                        break;
                    default:
                        if (this.idbTransaction == null)
                            idbTransaction = idbConnection.BeginTransaction(IsolationLevel.ReadCommitted);
                        break;
                }
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
        }

        /// <summary>
        /// In this function check the transcation object is not null then commit the transcation
        /// </summary>
        public void CommitTransaction()
        {
            msgError = "";
            try
            {
                if (this.idbTransaction != null)
                    this.idbTransaction.Commit();
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
        }

        /// <summary>
        /// In this function check the transcation object is not null then Rollback the transcation
        /// </summary>
        public void RollBackTransaction()
        {
            msgError = "";
            try
            {
                if (this.idbTransaction != null)
                    this.idbTransaction.Rollback();
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
        }

        /// <summary>
        /// In this function check the transcation object is not null then Abort(Dispose)  the transcation
        /// </summary>
        public void AbortTransaction()
        {
            msgError = "";
            try
            {
                if (this.idbTransaction != null)
                    this.idbTransaction.Dispose();
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
        }

        # endregion

        # region .... Query Executation ....

        protected internal int ExecuteNonQuery(string szQuery)
        {
            msgError = "";
            int iRowsAffected = -1;
            try
            {
                this.idbCommand = GetCommand();
                idbCommand.CommandText = szQuery;
                idbCommand.Transaction = idbTransaction;
                idbCommand.Connection = idbConnection;
                idbCommand.Parameters.Clear();
                iRowsAffected = idbCommand.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                iRowsAffected = -1;
                msgError = ex.Message;
            }
            finally
            {
                if (idbCommand != null)
                    idbCommand.Dispose();
                idbCommand = null;
            }
            return iRowsAffected;
        }

        protected internal bool bExecuteNonQuery(string szQuery)
        {
            _bResult = true;

            if (ExecuteNonQuery(szQuery) == -1)
                _bResult = false;

            return _bResult;
        }

        /// <summary>
        /// This function Execute Insert,Update and Delete Query.
        /// 1]Get command by calling GetCommand function
        /// 2]Assign Query to CommandText.
        /// 3]Assign Transcation to Command
        /// 4}Assign Connection to Command
        /// 5]Execute query by using idbCommand.ExecuteNonQuery() method
        /// </summary>
        /// <param name="szQuery"></param>
        /// <returns>if Query Execute Successful then it return true else false</returns>
        protected internal bool ExecuteNonQuery(string szQuery, string szFilePath, bool bFlag, ArrayList ArrBlob)
        {
            msgError = "";
            _bResult = true;
            try
            {
                this.idbCommand = GetCommand();
                idbCommand.CommandText = szQuery;
                idbCommand.Transaction = idbTransaction;
                idbCommand.Connection = idbConnection;
                idbCommand.Parameters.Clear();

                if (bFlag)
                {
                    int iCnt;
                    string szPrmName = "";
                    for (iCnt = 0; iCnt < ArrBlob.Count; iCnt++)
                    {
                        //szFilePath = ArrBlobPath[iCnt].ToString();
                        //szPrmName = "@blob_" + iCnt.ToString();
                        //FileStream fsBLOBFile = new FileStream(szFilePath, FileMode.Open, FileAccess.Read);
                        //Byte[] bytBLOBData = new Byte[fsBLOBFile.Length - 1];
                        //fsBLOBFile.Read(bytBLOBData, 0, bytBLOBData.Length);
                        //fsBLOBFile.Close();
                        //fsBLOBFile = null;
                        szPrmName = "@blob_" + iCnt.ToString();
                        Byte[] bytData = (Byte[])Convert.ChangeType(ArrBlob[iCnt], typeof(Byte[]));
                        OleDbParameter prm = new OleDbParameter(szPrmName, OleDbType.VarBinary, bytData.Length, ParameterDirection.Input, false, 0, 0, null, DataRowVersion.Current, bytData);
                        idbCommand.Parameters.Add(prm);
                    }
                }
                if (idbCommand.ExecuteNonQuery() < 0)
                    _bResult = false;
            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = ex.Message;
            }
            finally
            {
                if (idbCommand != null)
                    idbCommand.Dispose();
                idbCommand = null;
            }
            return _bResult;
        }

        /// <summary>
        /// This function Execute Insert,Update and Delete Query.
        /// 1]Get command by calling GetCommand function
        /// 2]Assign Query to CommandText.
        /// 3]Assign Transcation to Command
        /// 4}Assign Connection to Command
        /// 5]Execute query by using idbCommand.ExecuteNonQuery() method
        /// </summary>
        /// <param name="szQuery"></param>
        /// <returns>if Query Execute Successful then it return true else false</returns>
        protected internal bool ExecuteNonQuery(string szQuery, string szFilePath, bool bFlag, ArrayList ArrBlob, bool bGetIdentity, out string szSurrogateKey)
        {
            msgError = "";
            _bResult = true;
            szSurrogateKey = "";
            try
            {
                this.idbCommand = GetCommand();
                idbCommand.CommandText = szQuery;
                idbCommand.Transaction = idbTransaction;
                idbCommand.Connection = idbConnection;

                idbCommand.Parameters.Clear();

                if (bFlag)
                {
                    int iCnt;
                    string szPrmName = "";
                    for (iCnt = 0; iCnt < ArrBlob.Count; iCnt++)
                    {
                        szPrmName = "@blob_" + iCnt.ToString();
                        Byte[] bytData = (Byte[])Convert.ChangeType(ArrBlob[iCnt], typeof(Byte[]));
                        OleDbParameter prm = new OleDbParameter(szPrmName, OleDbType.VarBinary, bytData.Length, ParameterDirection.Input, false, 0, 0, null, DataRowVersion.Current, bytData);
                        idbCommand.Parameters.Add(prm);
                    }
                }
                OleDbParameter prmIdentity = new OleDbParameter("@SurrKey", OleDbType.Numeric);
                prmIdentity.Direction = ParameterDirection.Output;
                if (bGetIdentity)
                {
                    idbCommand.Parameters.Add(prmIdentity);
                }

                if (idbCommand.ExecuteNonQuery() < 0)
                    _bResult = false;

                szSurrogateKey = prmIdentity.Value.ToString();
            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = ex.Message;
            }
            finally
            {
                if (idbCommand != null)
                    idbCommand.Dispose();
                idbCommand = null;
            }
            return _bResult;
        }

        /// <summary>
        /// This function Execute Select Query.
        /// 1]Get command by calling GetCommand function
        /// 2]Assign Query to CommandText.
        /// 4}Assign Connection to Command
        /// 5]Execute query by using idbCommand.ExecuteReader() method
        /// </summary>
        /// <param name="szQuery"></param>
        /// <returns>IDataReader</returns>
        protected internal IDataReader ExecuteReader(string szQuery)
        {
            msgError = "";
            objDtReader = null;
            try
            {
                this.idbCommand = GetCommand();
                idbCommand.CommandText = szQuery;
                idbCommand.Connection = idbConnection;
                if (idbTransaction != null)
                    idbCommand.Transaction = idbTransaction;
                objDtReader = idbCommand.ExecuteReader();
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
                if (objDtReader != null)
                {
                    objDtReader.Close();
                    objDtReader.Dispose();
                }
                objDtReader = null;
            }
            finally
            {
                if (idbCommand != null)
                    idbCommand.Dispose();
                idbCommand = null;
            }
            return objDtReader;
        }

        /// <summary>
        /// This function Execute Select Query.
        /// 1]Get command by calling GetCommand function
        /// 2]Assign Query to CommandText.
        /// 4}Assign Connection to Command
        /// 5]Execute query by using idbCommand.ExecuteReader() method
        /// </summary>
        /// <param name="szQuery"></param>
        /// <returns>IDataReader</returns>
        protected internal IDataReader ExecuteReader(string szQuery, CommandBehavior eCommandBehavior)
        {
            msgError = "";
            objDtReader = null;
            try
            {
                this.idbCommand = GetCommand();
                idbCommand.CommandText = szQuery;
                idbCommand.Connection = idbConnection;
                if (idbTransaction != null)
                    idbCommand.Transaction = idbTransaction;

                if (_eConnectionType == ConnectionType.Global)
                    objDtReader = idbCommand.ExecuteReader();
                else
                    objDtReader = idbCommand.ExecuteReader(eCommandBehavior);
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
                if (objDtReader != null)
                {
                    objDtReader.Close();
                    objDtReader.Dispose();
                }
                objDtReader = null;
            }
            finally
            {
                if (idbCommand != null)
                    idbCommand.Dispose();
                idbCommand = null;
            }
            return objDtReader;
        }

        protected internal DataSet ExecuteDataSet(string szQuery)
        {
            msgError = "";
            objDataset = null;
            try
            {
                this.idbCommand = GetCommand();
                idbDtAdapter = GetDataAdapter();
                idbCommand.Transaction = idbTransaction;
                idbDtAdapter.SelectCommand = idbCommand;
                idbDtAdapter.SelectCommand.CommandText = szQuery;
                idbDtAdapter.SelectCommand.Connection = idbConnection;

                objDataset = new DataSet();
                idbDtAdapter.Fill(objDataset);
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
                if (objDataset != null)
                    objDataset.Dispose();
                objDataset = null;
            }
            finally
            {
                idbDtAdapter = null;

                if (idbCommand != null)
                    idbCommand.Dispose();
                idbCommand = null;
            }
            return objDataset;
        }

        protected internal DataTable ExecuteDataTable(string szQuery)
        {
            msgError = "";
            objDtTable = null;
            objDataset = null;
            try
            {
                this.idbCommand = GetCommand();
                idbDtAdapter = GetDataAdapter();
                idbCommand.Transaction = idbTransaction;
                idbDtAdapter.SelectCommand = idbCommand;
                idbDtAdapter.SelectCommand.CommandText = szQuery;
                idbDtAdapter.SelectCommand.Connection = idbConnection;

                objDataset = new DataSet();
                idbDtAdapter.Fill(objDataset);
                objDtTable = objDataset.Tables[0];
                objDataset.Dispose();
                objDataset = null;
            }
            catch (Exception ex)
            {
                msgError = ex.Message;

                if (objDtTable != null)
                    objDtTable.Dispose();
                objDtTable = null;

                if (objDataset != null)
                    objDataset.Dispose();
                objDataset = null;
            }
            finally
            {
                idbDtAdapter = null;

                if (idbCommand != null)
                    idbCommand.Dispose();
                idbCommand = null;
            }
            return objDtTable;
        }

        /// <summary>
        /// This function Execute Select Query to get first column value.
        /// 1]Get command by calling GetCommand function
        /// 2]Assign Query to CommandText.
        /// 4}Assign Connection to Command
        /// 5]Execute query by using idbCommand.ExecuteScalar() method
        /// </summary>
        /// <param name="szQuery"></param>
        /// <returns>it retrun boolean value if record exist otherwise it return false</returns>
        protected internal bool ExecuteScalar_bool(string szQuery)
        {
            msgError = "";
            object objReturn = null;
            bool bReturn = false;
            try
            {
                this.idbCommand = GetCommand();
                idbCommand.CommandText = szQuery;
                idbCommand.Transaction = idbTransaction;
                idbCommand.Connection = idbConnection;
                objReturn = idbCommand.ExecuteScalar();
                if (objReturn != null)
                    bReturn = true;
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            finally
            {
                objReturn = null;
                if (idbCommand != null)
                    idbCommand.Dispose();
                idbCommand = null;
            }
            return bReturn;
        }

        /// <summary>
        /// This function Execute Insert,Update and Delete Query.
        /// 1]Get command by calling GetCommand function
        /// 2]Assign Query to CommandText.
        /// 4}Assign Connection to Command
        /// 5]Execute query by using idbCommand.ExecuteScalar() method
        /// </summary>
        /// <param name="szQuery"></param>
        /// <returns>it retrun No of records affected.</returns>
        protected internal int ExecuteScalar(string szQuery)
        {
            msgError = "";
            object objReturn = null;
            int iReturn = 0;
            try
            {
                this.idbCommand = GetCommand();
                idbCommand.CommandText = szQuery;
                idbCommand.Transaction = idbTransaction;
                idbCommand.Connection = idbConnection;
                objReturn = idbCommand.ExecuteScalar();
                if (objReturn != null && !(objReturn is DBNull))
                    iReturn = Convert.ToInt32(objReturn);
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            finally
            {
                objReturn = null;
                if (idbCommand != null)
                    idbCommand.Dispose();
                idbCommand = null;
            }
            return iReturn;
        }

        /// <summary>
        /// This function Execute Select Query to get first column value.
        /// 1]Get command by calling GetCommand function
        /// 2]Assign Query to CommandText.
        /// 4}Assign Connection to Command
        /// 5]Execute query by using idbCommand.ExecuteScalar() method
        /// </summary>
        /// <param name="szQuery"></param>
        /// <returns>it retrun string beacuse if any Exception occured then it return this Exception</returns>
        protected internal object ExecuteScalar_Object(string szQuery)
        {
            msgError = "";
            object objReturn = null;
            try
            {
                this.idbCommand = GetCommand();
                idbCommand.CommandText = szQuery;
                idbCommand.Transaction = idbTransaction;
                idbCommand.Connection = idbConnection;
                objReturn = idbCommand.ExecuteScalar();
            }
            catch (Exception ex)
            {
                objReturn = null;
                msgError = ex.Message;
            }
            finally
            {
                if (idbCommand != null)
                    idbCommand.Dispose();
                idbCommand = null;
            }
            return objReturn;
        }

        protected internal bool ExecuteStoreBLOBIntoFile(string szTemplateFilepath, string szTableName, string szBlobColumnName, string szConstraint)
        {
            msgError = "";
            int ImageCol = 0;  // position of Picture column in DataReader
            int BUFFER_LENGTH = 32768; // chunk size
            string szQuery;
            _bResult = true;
            try
            {
                szQuery = "Select ?=TEXTPTR(" + szBlobColumnName + "), ?=DataLength(" + szBlobColumnName + ") from " + szTableName + " " + szConstraint;

                this.idbCommand = GetCommand();
                idbCommand.CommandText = szQuery;
                idbCommand.Transaction = idbTransaction;
                idbCommand.Connection = idbConnection;

                OleDbParameter PointerOutParam = new OleDbParameter("@Pointer", OleDbType.VarBinary, 100);
                PointerOutParam.Direction = ParameterDirection.Output;
                OleDbParameter LengthOutParam = new OleDbParameter("@Length", OleDbType.Integer);
                LengthOutParam.Direction = ParameterDirection.Output;

                idbCommand.Parameters.Add(PointerOutParam);
                idbCommand.Parameters.Add(LengthOutParam);

                idbCommand.ExecuteNonQuery();
                if (PointerOutParam.Value == null)
                {
                    _bResult = false;
                    return _bResult;
                }

                // Set up READTEXT command, parameters, and open BinaryReader.
                idbCommand.Parameters.Clear();

                szQuery = "READTEXT " + szTableName + "." + szBlobColumnName + " ? ? ? HOLDLOCK";
                this.idbCommand = GetCommand();
                idbCommand.CommandText = szQuery;
                idbCommand.Transaction = idbTransaction;
                idbCommand.Connection = idbConnection;

                OleDbParameter PointerParam = new OleDbParameter("@Pointer", OleDbType.Binary, 16);
                OleDbParameter OffsetParam = new OleDbParameter("@Offset", OleDbType.Integer);
                OleDbParameter SizeParam = new OleDbParameter("@Size", OleDbType.Integer);

                idbCommand.Parameters.Add(PointerParam);
                idbCommand.Parameters.Add(OffsetParam);
                idbCommand.Parameters.Add(SizeParam);

                System.IO.FileStream fs = new System.IO.FileStream(szTemplateFilepath, System.IO.FileMode.Create, System.IO.FileAccess.Write);
                int Offset = 0;
                OffsetParam.Value = Offset;
                Byte[] Buffer = new Byte[BUFFER_LENGTH];

                // Read buffer full of data and write to the file stream.
                do
                {
                    PointerParam.Value = PointerOutParam.Value;

                    // Calculate buffer size - may be less than BUFFER_LENGTH for last block.
                    if ((Offset + BUFFER_LENGTH) >= System.Convert.ToInt32(LengthOutParam.Value))
                        SizeParam.Value = System.Convert.ToInt32(LengthOutParam.Value) - Offset;
                    else SizeParam.Value = BUFFER_LENGTH;

                    objDtReader = idbCommand.ExecuteReader();
                    objDtReader.Read();
                    objDtReader.GetBytes(ImageCol, 0, Buffer, 0, System.Convert.ToInt32(SizeParam.Value));
                    objDtReader.Close();
                    objDtReader.Dispose();
                    objDtReader = null;

                    fs.Write(Buffer, 0, System.Convert.ToInt32(SizeParam.Value));
                    Offset += System.Convert.ToInt32(SizeParam.Value);
                    OffsetParam.Value = Offset;
                } while (Offset < System.Convert.ToInt32(LengthOutParam.Value));
                fs.Flush();
                fs.Close();
                fs.Dispose();
                fs = null;
                _bResult = true;
            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = ex.Message;
            }
            finally
            {
                if (idbCommand != null)
                    idbCommand.Dispose();
                idbCommand = null;
            }
            return _bResult;
        }

        # endregion

        #region ..... Functions for IDisposable Interface .....

        #region Variable Declaration for Disposable Object ...

        private IntPtr handle;
        private Component CompExecuteQuery = new Component();
        private bool bDisposed = false;

        #endregion

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        [System.Runtime.InteropServices.DllImport("Kernel32")]
        private extern static Boolean CloseHandle(IntPtr handle);

        ~ClsExecuteQuery()
        {
            Dispose(false);
        }

        protected virtual void Dispose(bool bDisposing)
        {
            if (!this.bDisposed)
            {
                if (bDisposing)
                {
                    if (objGetConInfo != null)
                        objGetConInfo.Dispose();

                    if (objDataset != null)
                        objDataset.Dispose();

                    if (objDtTable != null)
                        objDtTable.Dispose();

                    if (objDtReader != null)
                    {
                        objDtReader.Close();
                        objDtReader.Dispose();
                    }

                    if (idbCommand != null)
                        idbCommand.Dispose();

                    if (_eConnectionType != ConnectionType.Global)
                    {
                        if (idbTransaction != null)
                            idbTransaction.Dispose();

                        if (idbConnection != null)
                        {
                            if (idbConnection.State != ConnectionState.Closed)
                            {
                                try { idbConnection.Close(); }
                                catch (Exception) { }
                            }
                            idbConnection.Dispose();
                        }
                    }

                    if (CompExecuteQuery != null)
                        CompExecuteQuery.Dispose();
                    CompExecuteQuery = null;
                }

                objGetConInfo = null;
                objDataset = null;
                objDtTable = null;
                objDtReader = null;

                idbConnection = null;
                idbTransaction = null;
                idbCommand = null;
                idbDtAdapter = null;

                ConnAppSiteUri = null;
                DataBase = null;
                Product_Name = null;
                Region = null;
                ConnectionName = null;
                AppXmlPath = null;
                msgError = null;

                CloseHandle(handle);
                handle = IntPtr.Zero;
                bDisposed = true;
            }
        }

        #endregion
    }
}