//... Code Changed By Manavya on 22/07/2011 For DRT-2783 ...
//... Code Changed By Manavya on 04/12/2012 For DRT-4136 ...
using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
using System.Data;
using System.Data.OleDb;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.ComponentModel;

namespace DDLLCS
{
    public class ClsBuildQuery : ClsExecuteQuery
    {
        # region .... Variables Declaration ....

        private IDataReader objDtReader;
        private DataSet objDataset;
        private ArrayList ArrBolbData = null;
        private StringBuilder _szbFields = null;
        private StringBuilder _szbValues = null;

        private bool _bResult;
        private bool _blobFlag = false;
        private string _szQuery = null;
        private string _szDatatype = null;

        # endregion

        # region .... Constructors ....

        public ClsBuildQuery()
            : base(IntPtr.Zero)
        {
            msgError = "";
            DataBase = "MSSQL";
        }

        /// <summary>
        /// In this constructor used for assing value of Connection Application site URI to clsDal calss 
        /// </summary>
        /// <param name="uriConnAppSite">Uri of Common/Connection Application site.</param>
        public ClsBuildQuery(Uri uriConnAppSite)
            : base(IntPtr.Zero)
        {
            msgError = "";
            AppXmlPath = "";
            DataBase = "MSSQL";
            ConnAppSiteUri = uriConnAppSite;
        }

        /// <summary>
        /// In this constructor used for assing value of Connection Application site URI to clsDal calss 
        /// </summary>
        /// <param name="uriConnAppSite">Uri of Common/Connection Application site.</param>
        public ClsBuildQuery(IntPtr iPtrHandle, Uri uriConnAppSite)
            : base(iPtrHandle)
        {
            msgError = "";
            AppXmlPath = "";
            DataBase = "MSSQL";
            ConnAppSiteUri = uriConnAppSite;
        }

        /// <summary>
        /// in this constructor used for assing value to DataBase and AppXmlPath property of clsDal calss 
        /// </summary>
        /// <param name="szAppXmlPath">Application.Xml Path</param>
        public ClsBuildQuery(string szAppXmlPath)
            : base(IntPtr.Zero)
        {
            msgError = "";
            AppXmlPath = szAppXmlPath;
            DataBase = "MSSQL";
            ConnAppSiteUri = null;
        }

        public ClsBuildQuery(IntPtr iPtrHandle, string szAppXmlPath)
            : base(iPtrHandle)
        {
            msgError = "";
            AppXmlPath = szAppXmlPath;
            DataBase = "MSSQL";
            ConnAppSiteUri = null;
            //DataBase = szDataBaseName;
            //DataBaseConnKeyName = "DbConn";
        }

        /// <summary>
        /// in this constructor used for assing value to DataBase and AppXmlPath property of clsDal calss 
        /// </summary>
        /// <param name="szDataBaseName">Database Name (MSSQL)</param>
        /// <param name="szAppXmlPath">Application.Xml Path</param>
        public ClsBuildQuery(string szDataBaseName, string szAppXmlPath)
            : base(IntPtr.Zero)
        {
            msgError = "";
            AppXmlPath = szAppXmlPath;
            DataBase = "MSSQL";
            ConnAppSiteUri = null;
            //DataBase = szDataBaseName;
            //DataBaseConnKeyName = "DbConn";
        }

        # endregion

        # region .... Data Insert Update Delete ....

        /// <summary>
        /// in this function insert query build according to Arraylist and tablename
        /// and pass it to ExecuteNonQuery of ClsDal Class.
        /// </summary>
        /// <param name="szTableName"></param>
        /// <param name="ArrParameterList"></param>
        /// <returns>if Query Execute Successfuly then it return DONE else Exception </returns>
        public bool Insert(string szTableName, ArrayList ArrParameterList)
        {
            _bResult = true;
            _szbFields = new StringBuilder("");
            _szbValues = new StringBuilder("");
            ArrBolbData = new ArrayList();
            ArrBolbData.Clear();
            try
            {
                int iCnt, iSeqNo;
                string szMyQuery = null;
                int iListCount = ArrParameterList.Count;
                for (iCnt = 0; iCnt < iListCount; iCnt += 3)
                {
                    _szDatatype = null;
                    _szDatatype = ArrParameterList[iCnt + 1].ToString().ToUpper();
                    //if Values Are Not None
                    if (_szbValues.ToString() == "")
                    {
                        switch (_szDatatype)
                        {
                            case "VARCHAR":
                                _szbValues.Append("'" + ArrParameterList[iCnt + 2].ToString() + "'");
                                break;
                            case "NUMBER":
                            case "BIT":
                                _szbValues.Append(ArrParameterList[iCnt + 2].ToString());
                                break;
                            case "DATE":
                                if (DataBase.ToUpper() == "MSSQL")
                                    _szbValues.Append("CONVERT(varchar, '" + ArrParameterList[iCnt + 2].ToString() + "', 101)");
                                else if (DataBase.ToUpper() == "ORACLE")
                                    _szbValues.Append("to_date('" + ArrParameterList[iCnt + 2].ToString() + "', 'MM-DD-YYYY')");
                                break;
                            case "TIME":
                                if (DataBase.ToUpper() == "MSSQL")
                                    _szbValues.Append("CONVERT(varchar, '" + ArrParameterList[iCnt + 2].ToString() + "', 22)");
                                else if (DataBase.ToUpper() == "ORACLE")
                                    _szbValues.Append("to_date('" + ArrParameterList[iCnt + 2].ToString() + "', 'HH24:MI:SS')");
                                break;
                            case "DATETIME":
                                if (DataBase.ToUpper() == "MSSQL")
                                    _szbValues.Append("CONVERT(varchar, '" + ArrParameterList[iCnt + 2].ToString() + "', 100)");
                                else if (DataBase.ToUpper() == "ORACLE")
                                    _szbValues.Append("to_date('" + ArrParameterList[iCnt + 2].ToString() + "', 'MM-DD-YYYY HH24:MI:SS')");
                                break;
                            case "SEQUENCE":
                                if (DataBase.ToUpper() == "MSSQL")
                                    continue;
                                else if (DataBase.ToUpper() == "ORACLE")
                                    szMyQuery = "SELECT SQ_" + szTableName + ".NEXTVAL FROM DUAL";
                                iSeqNo = ExecuteScalar(szMyQuery);
                                _szbValues.Append(iSeqNo.ToString());
                                break;
                            case "BLOB":
                                _blobFlag = true;
                                _szbValues.Append("?");
                                ArrBolbData.Add(ArrParameterList[iCnt + 2]);
                                break;
                            default:
                                _szbValues.Append(ArrParameterList[iCnt + 2].ToString());
                                break;
                        }

                    } //  if (szValues == "")

                    //if Values Are Not None
                    else
                    {
                        switch (_szDatatype)
                        {
                            case "VARCHAR":
                                _szbValues.Append(" ,'" + ArrParameterList[iCnt + 2].ToString() + "'");
                                break;
                            case "NUMBER":
                            case "BIT":
                                _szbValues.Append(" ," + ArrParameterList[iCnt + 2].ToString());
                                break;
                            case "DATE":
                                if (DataBase.ToUpper() == "MSSQL")
                                    _szbValues.Append(", CONVERT(varchar, '" + ArrParameterList[iCnt + 2].ToString() + "', 101)");
                                else if (DataBase.ToUpper() == "ORACLE")
                                    _szbValues.Append(", to_date('" + ArrParameterList[iCnt + 2].ToString() + "', 'MM-DD-YYYY')");
                                break;
                            case "TIME":
                                if (DataBase.ToUpper() == "MSSQL")
                                    _szbValues.Append(", CONVERT(varchar,'" + ArrParameterList[iCnt + 2].ToString() + "', 22)");
                                else if (DataBase.ToUpper() == "ORACLE")
                                    _szbValues.Append(" , to_date('" + ArrParameterList[iCnt + 2].ToString() + "', 'HH24:MI:SS')");
                                break;
                            case "DATETIME":
                                if (DataBase.ToUpper() == "MSSQL")
                                    _szbValues.Append(", CONVERT(varchar, '" + ArrParameterList[iCnt + 2].ToString() + "', 100)");
                                else if (DataBase.ToUpper() == "ORACLE")
                                    _szbValues.Append(", to_date('" + ArrParameterList[iCnt + 2].ToString() + "', 'MM-DD-YYYY HH24:MI:SS')");
                                break;
                            case "SEQUENCE":
                                if (DataBase == "MSSQL")
                                    continue;
                                else if (DataBase == "ORACLE")
                                    szMyQuery = "SELECT SQ_" + szTableName + ".NEXTVAL FROM DUAL";
                                _szbValues.Append(" ," + ExecuteScalar(szMyQuery));
                                break;
                            case "BLOB":
                                _szbValues.Append(", " + "?");
                                _blobFlag = true;
                                ArrBolbData.Add(ArrParameterList[iCnt + 2]);
                                break;
                            default:
                                _szbValues.Append(" ," + ArrParameterList[iCnt + 2].ToString());
                                break;
                        }
                    } //else  if (szValues == "")

                    if (_szbFields.ToString() == "")
                        _szbFields.Append(ArrParameterList[iCnt].ToString());
                    else
                        _szbFields.Append(", " + ArrParameterList[iCnt].ToString());

                } //for (iCnt = 0; iCnt <= iArrCount; iCnt++)
                _szDatatype = null;

                _szQuery = " Insert into " + szTableName + " (" + _szbFields.ToString() + ") values (" + _szbValues.ToString() + ")";
                _szbFields = null;
                _szbValues = null;

                _bResult = ExecuteNonQuery(_szQuery, "", _blobFlag, ArrBolbData);
            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = ex.Message;
            }
            finally
            {
                _szbFields = null;
                _szbValues = null;
                _szDatatype = null;
                _szQuery = null;
                if (ArrBolbData != null)
                    ArrBolbData.Clear();
                ArrBolbData = null;
            }
            return _bResult;
        }

        /// <summary>
        /// in this function insert query build according to Arraylist and tablename
        /// and pass it to ExecuteNonQuery of ClsDal Class.
        /// </summary>
        /// <param name="szTableName"></param>
        /// <param name="ArrParameterList"></param>
        /// <returns>if Query Execute Successfuly then it return DONE else Exception </returns>
        public bool Insert(string szTableName, ArrayList ArrParameterList, bool bGetIdentity, out string szSurrogateKey)
        {
            _bResult = true;
            _szbFields = new StringBuilder("");
            _szbValues = new StringBuilder("");
            ArrBolbData = new ArrayList();
            ArrBolbData.Clear();
            szSurrogateKey = "";
            try
            {
                int iListCount = ArrParameterList.Count;
                int iCnt, iSeqNo;
                string szMyQuery = null;
                for (iCnt = 0; iCnt < iListCount; iCnt += 3)
                {
                    _szDatatype = null;
                    _szDatatype = ArrParameterList[iCnt + 1].ToString().ToUpper();
                    //if Values Are Not None
                    if (_szbValues.ToString() == "")
                    {
                        switch (_szDatatype)
                        {
                            case "VARCHAR":
                                _szbValues.Append("'" + ArrParameterList[iCnt + 2].ToString() + "'");
                                break;
                            case "NUMBER":
                            case "BIT":
                                _szbValues.Append(ArrParameterList[iCnt + 2].ToString());
                                break;
                            case "DATE":
                                if (DataBase.ToUpper() == "MSSQL")
                                    _szbValues.Append("CONVERT(varchar, '" + ArrParameterList[iCnt + 2].ToString() + "', 101)");
                                else if (DataBase.ToUpper() == "ORACLE")
                                    _szbValues.Append("to_date('" + ArrParameterList[iCnt + 2].ToString() + "', 'MM-DD-YYYY')");
                                break;
                            case "TIME":
                                if (DataBase.ToUpper() == "MSSQL")
                                    _szbValues.Append("CONVERT(varchar, '" + ArrParameterList[iCnt + 2].ToString() + "', 22)");
                                else if (DataBase.ToUpper() == "ORACLE")
                                    _szbValues.Append("to_date('" + ArrParameterList[iCnt + 2].ToString() + "', 'HH24:MI:SS')");
                                break;
                            case "DATETIME":
                                if (DataBase.ToUpper() == "MSSQL")
                                    _szbValues.Append("CONVERT(varchar, '" + ArrParameterList[iCnt + 2].ToString() + "', 100)");
                                else if (DataBase.ToUpper() == "ORACLE")
                                    _szbValues.Append("to_date('" + ArrParameterList[iCnt + 2].ToString() + "', 'MM-DD-YYYY HH24:MI:SS')");
                                break;
                            case "SEQUENCE":
                                if (DataBase.ToUpper() == "MSSQL")
                                    continue;
                                else if (DataBase.ToUpper() == "ORACLE")
                                    szMyQuery = "SELECT SQ_" + szTableName + ".NEXTVAL FROM DUAL";
                                iSeqNo = ExecuteScalar(szMyQuery);
                                _szbValues.Append(iSeqNo.ToString());
                                break;
                            case "BLOB":
                                _blobFlag = true;
                                _szbValues.Append("?");
                                ArrBolbData.Add(ArrParameterList[iCnt + 2]);
                                break;
                            default:
                                _szbValues.Append(ArrParameterList[iCnt + 2].ToString());
                                break;
                        }

                    } //  if (szValues == "")

                    //if Values Are Not None
                    else
                    {
                        switch (_szDatatype)
                        {
                            case "VARCHAR":
                                _szbValues.Append(" ,'" + ArrParameterList[iCnt + 2].ToString() + "'");
                                break;
                            case "NUMBER":
                            case "BIT":
                                _szbValues.Append(" ," + ArrParameterList[iCnt + 2].ToString());
                                break;
                            case "DATE":
                                if (DataBase.ToUpper() == "MSSQL")
                                    _szbValues.Append(", CONVERT(varchar, '" + ArrParameterList[iCnt + 2].ToString() + "', 101)");
                                else if (DataBase.ToUpper() == "ORACLE")
                                    _szbValues.Append(", to_date('" + ArrParameterList[iCnt + 2].ToString() + "', 'MM-DD-YYYY')");
                                break;
                            case "TIME":
                                if (DataBase.ToUpper() == "MSSQL")
                                    _szbValues.Append(", CONVERT(varchar,'" + ArrParameterList[iCnt + 2].ToString() + "', 22)");
                                else if (DataBase.ToUpper() == "ORACLE")
                                    _szbValues.Append(" , to_date('" + ArrParameterList[iCnt + 2].ToString() + "', 'HH24:MI:SS')");
                                break;
                            case "DATETIME":
                                if (DataBase.ToUpper() == "MSSQL")
                                    _szbValues.Append(", CONVERT(varchar, '" + ArrParameterList[iCnt + 2].ToString() + "', 100)");
                                else if (DataBase.ToUpper() == "ORACLE")
                                    _szbValues.Append(", to_date('" + ArrParameterList[iCnt + 2].ToString() + "', 'MM-DD-YYYY HH24:MI:SS')");
                                break;
                            case "SEQUENCE":
                                if (DataBase == "MSSQL")
                                    continue;
                                else if (DataBase == "ORACLE")
                                    szMyQuery = "SELECT SQ_" + szTableName + ".NEXTVAL FROM DUAL";
                                _szbValues.Append(" ," + ExecuteScalar(szMyQuery));
                                break;
                            case "BLOB":
                                _szbValues.Append(", " + "?");
                                _blobFlag = true;
                                ArrBolbData.Add(ArrParameterList[iCnt + 2]);
                                break;
                            default:
                                _szbValues.Append(" ," + ArrParameterList[iCnt + 2].ToString());
                                break;
                        }
                    } //else  if (szValues == "")

                    if (_szbFields.ToString() == "")
                        _szbFields.Append(ArrParameterList[iCnt].ToString());
                    else
                        _szbFields.Append(", " + ArrParameterList[iCnt].ToString());

                } //for (iCnt = 0; iCnt <= iArrCount; iCnt++)
                _szDatatype = null;

                _szQuery = " Insert into " + szTableName + " (" + _szbFields.ToString() + ") values (" + _szbValues.ToString() + ")";
                _szbFields = null;
                _szbValues = null;

                if (bGetIdentity && DataBase.ToUpper() == "MSSQL")
                    _szQuery = _szQuery + " SET ? = SCOPE_IDENTITY()";

                _bResult = ExecuteNonQuery(_szQuery, "", _blobFlag, ArrBolbData, bGetIdentity, out szSurrogateKey);
            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = ex.Message;
            }
            finally
            {
                _szbFields = null;
                _szbValues = null;
                _szDatatype = null;
                _szQuery = null;
                if (ArrBolbData != null)
                    ArrBolbData.Clear();
                ArrBolbData = null;
            }
            return _bResult;
        }

        /// <summary>
        /// in this function insert query build according to Arraylist and tablename
        /// and pass it to ExecuteNonQuery of ClsDal Class.
        /// if Query execute Successefuly and no rows updated then it return NOTDONE
        /// </summary>
        /// <param name="szTableName"></param>
        /// <param name="ArrParameterList"></param>
        /// <param name="ArrWParameterList"></param>
        /// <returns>if Query Execute Successfuly then it return DONE or NOTDONE else Exception</returns>
        public bool Update(string szTableName, ArrayList ArrParameterList, ArrayList ArrWParameterList)
        {
            _bResult = true;
            _szbFields = new StringBuilder("");
            _szbValues = new StringBuilder("");
            ArrBolbData = new ArrayList();
            ArrBolbData.Clear();
            try
            {
                int iListCount = ArrParameterList.Count;
                int iCnt;
                for (iCnt = 0; iCnt < iListCount; iCnt += 3)
                {
                    _szDatatype = null;
                    _szDatatype = ArrParameterList[iCnt + 1].ToString().ToUpper();
                    //if Values Are Not None
                    if (_szbFields.ToString() == "")
                    {
                        switch (_szDatatype)
                        {
                            case "VARCHAR":
                                _szbFields.Append(ArrParameterList[iCnt].ToString() + " = '" + ArrParameterList[iCnt + 2].ToString() + "'");
                                break;
                            case "NUMBER":
                            case "BIT":
                                _szbFields.Append(ArrParameterList[iCnt].ToString() + " = " + ArrParameterList[iCnt + 2].ToString());
                                break;
                            case "DATE":
                                if (DataBase.ToUpper() == "MSSQL")
                                    _szbFields.Append(ArrParameterList[iCnt].ToString() + " = CONVERT(varchar, '" + ArrParameterList[iCnt + 2].ToString() + "', 101)");
                                else if (DataBase.ToUpper() == "ORACLE")
                                    _szbFields.Append(ArrParameterList[iCnt].ToString() + " = to_date('" + ArrParameterList[iCnt + 2].ToString() + "', 'MM-DD-YYYY')");
                                break;
                            case "TIME":
                                if (DataBase.ToUpper() == "MSSQL")
                                    _szbFields.Append(ArrParameterList[iCnt].ToString() + " = CONVERT(varchar, '" + ArrParameterList[iCnt + 2].ToString() + "', 22)");
                                else if (DataBase.ToUpper() == "ORACLE")
                                    _szbFields.Append(ArrParameterList[iCnt].ToString() + " = to_date('" + ArrParameterList[iCnt + 2].ToString() + "', 'HH24:MI:SS')");
                                break;
                            case "DATETIME":
                                if (DataBase.ToUpper() == "MSSQL")
                                    _szbFields.Append(ArrParameterList[iCnt].ToString() + " = CONVERT(varchar, '" + ArrParameterList[iCnt + 2].ToString() + "', 100)");
                                else if (DataBase.ToUpper() == "ORACLE")
                                    _szbFields.Append(ArrParameterList[iCnt].ToString() + " = to_date('" + ArrParameterList[iCnt + 2].ToString() + "', 'MM-DD-YYYY HH24:MI:SS')");
                                break;
                            case "BLOB":
                                //blobFlag = true;
                                //szFilePath = ArrParameterList[iCnt + 1].ToString();
                                //szFields = ArrParameterList[iCnt].ToString() + " = " + ArrParameterList[iCnt + 2].ToString();
                                _blobFlag = true;
                                _szbFields.Append(ArrParameterList[iCnt].ToString() + " = ?");
                                ArrBolbData.Add(ArrParameterList[iCnt + 2]);
                                break;
                            default:
                                _szbFields.Append(ArrParameterList[iCnt].ToString() + " = " + ArrParameterList[iCnt + 2].ToString());
                                break;
                        }
                    } //  if (szFields == "")
                    //if Values Are Not None
                    else
                    {
                        switch (_szDatatype)
                        {
                            case "VARCHAR":
                                _szbFields.Append(", " + ArrParameterList[iCnt].ToString() + " = '" + ArrParameterList[iCnt + 2].ToString() + "'");
                                break;

                            case "NUMBER":
                            case "BIT":
                                _szbFields.Append(", " + ArrParameterList[iCnt].ToString() + " = " + ArrParameterList[iCnt + 2].ToString());
                                break;

                            case "DATE":
                                if (DataBase.ToUpper() == "MSSQL")
                                    _szbFields.Append(", " + ArrParameterList[iCnt].ToString() + " = CONVERT(varchar, '" + ArrParameterList[iCnt + 2].ToString() + "', 101)");
                                else if (DataBase.ToUpper() == "ORACLE")
                                    _szbFields.Append(", " + ArrParameterList[iCnt].ToString() + " = to_date('" + ArrParameterList[iCnt + 2].ToString() + "', 'MM-DD-YYYY')");
                                break;

                            case "TIME":
                                if (DataBase.ToUpper() == "MSSQL")
                                    _szbFields.Append(", " + ArrParameterList[iCnt].ToString() + " = CONVERT(varchar, '" + ArrParameterList[iCnt + 2].ToString() + "', 22)");
                                else if (DataBase.ToUpper() == "ORACLE")
                                    _szbFields.Append(", " + ArrParameterList[iCnt].ToString() + " = to_date('" + ArrParameterList[iCnt + 2].ToString() + "', 'HH24:MI:SS')");
                                break;
                            case "DATETIME":
                                if (DataBase.ToUpper() == "MSSQL")
                                    _szbFields.Append(", " + ArrParameterList[iCnt].ToString() + " = CONVERT(varchar, '" + ArrParameterList[iCnt + 2].ToString() + "', 100)");
                                else if (DataBase.ToUpper() == "ORACLE")
                                    _szbFields.Append(", " + ArrParameterList[iCnt].ToString() + " = to_date('" + ArrParameterList[iCnt + 2].ToString() + "', 'MM-DD-YYYY HH24:MI:SS')");
                                break;
                            case "BLOB":
                                //blobFlag = true;
                                //szFilePath = ArrParameterList[iCnt + 1].ToString();
                                //szFields += ", " + ArrParameterList[iCnt].ToString() + " = " + ArrParameterList[iCnt + 2].ToString();

                                _blobFlag = true;
                                _szbFields.Append(", " + ArrParameterList[iCnt].ToString() + " = ?");
                                ArrBolbData.Add(ArrParameterList[iCnt + 2]);
                                break;

                            default:
                                _szbFields.Append(", " + ArrParameterList[iCnt].ToString() + " = " + ArrParameterList[iCnt + 2].ToString());
                                break;
                        }
                    } //else  if (szFields == "")
                } //for (iCnt = 0; iCnt <= iArrCount; iCnt++)
                _szDatatype = null;

                //------------------------------------ Where Parameter List -----------------------------
                iListCount = 0;
                iListCount = ArrWParameterList.Count;
                for (iCnt = 0; iCnt < iListCount; iCnt += 3)
                {
                    _szDatatype = null;
                    //if Values Are Not None
                    if (_szbValues.ToString() == "")
                    {
                        _szDatatype = ArrWParameterList[iCnt + 1].ToString().ToUpper();
                        switch (_szDatatype)
                        {
                            case "VARCHAR":
                                _szbValues.Append(ArrWParameterList[iCnt].ToString() + "'" + ArrWParameterList[iCnt + 2].ToString() + "'");
                                break;

                            case "NUMBER":
                            case "BIT":
                                _szbValues.Append(ArrWParameterList[iCnt].ToString() + ArrWParameterList[iCnt + 1].ToString() + ArrWParameterList[iCnt + 2].ToString());
                                break;

                            case "DATE":
                                if (DataBase.ToUpper() == "MSSQL")
                                    _szbValues.Append(ArrWParameterList[iCnt].ToString() + "CONVERT(varchar, '" + ArrWParameterList[iCnt + 2].ToString() + "', 101)");
                                else if (DataBase.ToUpper() == "ORACLE")
                                    _szbValues.Append(ArrWParameterList[iCnt].ToString() + "to_date('" + ArrWParameterList[iCnt + 2].ToString() + "', 'MM-DD-YYYY')");
                                break;

                            case "TIME":
                                if (DataBase.ToUpper() == "MSSQL")
                                    _szbValues.Append(ArrWParameterList[iCnt].ToString() + ArrWParameterList[iCnt + 1].ToString() + "CONVERT(varchar,'" + ArrWParameterList[iCnt + 2].ToString() + "', 22)");
                                else if (DataBase.ToUpper() == "ORACLE")
                                    _szbValues.Append(ArrWParameterList[iCnt].ToString() + ArrWParameterList[iCnt + 1].ToString() + "to_date('" + ArrWParameterList[iCnt + 2].ToString() + "', 'HH24:MI:SS')");
                                break;
                            case "DATETIME":
                                if (DataBase.ToUpper() == "MSSQL")
                                    _szbValues.Append(ArrWParameterList[iCnt].ToString() + "CONVERT(varchar, '" + ArrWParameterList[iCnt + 2].ToString() + "', 100)");
                                else if (DataBase.ToUpper() == "ORACLE")
                                    _szbValues.Append(ArrWParameterList[iCnt].ToString() + "to_date('" + ArrWParameterList[iCnt + 2].ToString() + "', 'MM-DD-YYYY HH24:MI:SS')");
                                break;
                            default:
                                _szbValues.Append(ArrWParameterList[iCnt].ToString() + ArrWParameterList[iCnt + 1].ToString() + ArrWParameterList[iCnt + 2].ToString());
                                break;
                        }
                    } //  if (szValues == "")
                    //if Values Are Not None
                    else
                    {
                        _szDatatype = ArrWParameterList[iCnt + 2].ToString().ToUpper();
                        switch (_szDatatype)
                        {
                            case "VARCHAR":
                                _szbValues.Append(ArrWParameterList[iCnt].ToString() + ArrWParameterList[iCnt + 1].ToString() + "'" + ArrWParameterList[iCnt + 3].ToString() + "'");
                                break;

                            case "NUMBER":
                            case "BIT":
                                _szbValues.Append(ArrWParameterList[iCnt].ToString() + ArrWParameterList[iCnt + 1].ToString() + ArrWParameterList[iCnt + 2].ToString() + ArrWParameterList[iCnt + 3].ToString());
                                break;
                            case "DATE":
                                if (DataBase.ToUpper() == "MSSQL")
                                    _szbValues.Append(ArrWParameterList[iCnt].ToString() + ArrWParameterList[iCnt + 1].ToString() + "CONVERT(varchar, '" + ArrWParameterList[iCnt + 3].ToString() + "', 101)");
                                else if (DataBase.ToUpper() == "ORACLE")
                                    _szbValues.Append(ArrWParameterList[iCnt].ToString() + ArrWParameterList[iCnt + 1].ToString() + " to_date('" + ArrWParameterList[iCnt + 3].ToString() + "', 'MM-DD-YYYY')");
                                break;

                            case "TIME":
                                if (DataBase.ToUpper() == "MSSQL")
                                    _szbValues.Append(ArrWParameterList[iCnt].ToString() + ArrWParameterList[iCnt + 1].ToString() + " CONVERT(varchar, '" + ArrWParameterList[iCnt + 3].ToString() + "', 22)");
                                else if (DataBase.ToUpper() == "ORACLE")
                                    _szbValues.Append(ArrWParameterList[iCnt].ToString() + ArrWParameterList[iCnt + 1].ToString() + " to_date('" + ArrWParameterList[iCnt + 3].ToString() + "', 'HH24:MI:SS')");
                                break;
                            case "DATETIME":
                                if (DataBase.ToUpper() == "MSSQL")
                                    _szbValues.Append(ArrWParameterList[iCnt].ToString() + ArrWParameterList[iCnt + 1].ToString() + "CONVERT(varchar, '" + ArrWParameterList[iCnt + 3].ToString() + "', 100)");
                                else if (DataBase.ToUpper() == "ORACLE")
                                    _szbValues.Append(ArrWParameterList[iCnt].ToString() + ArrWParameterList[iCnt + 1].ToString() + " to_date('" + ArrWParameterList[iCnt + 3].ToString() + "', 'MM-DD-YYYY HH24:MI:SS')");
                                break;
                            default:
                                _szbValues.Append(ArrWParameterList[iCnt].ToString() + ArrWParameterList[iCnt + 1].ToString() + ArrWParameterList[iCnt + 2].ToString() + ArrWParameterList[iCnt + 3].ToString());
                                break;
                        }
                    } //else  if (szValues == "")
                } //for (iCnt = 0; iCnt <= iArrCount; iCnt++)
                _szDatatype = null;
                // --------------------------------------------------------------------------

                _szQuery = " Update " + szTableName + " Set " + _szbFields.ToString() + " Where " + _szbValues.ToString();
                _szbFields = null;
                _szbValues = null;

                _bResult = ExecuteNonQuery(_szQuery, "", _blobFlag, ArrBolbData);
            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = ex.Message;
            }
            finally
            {
                _szbFields = null;
                _szbValues = null;
                _szDatatype = null;
                _szQuery = null;
                if (ArrBolbData != null)
                    ArrBolbData.Clear();
                ArrBolbData = null;
            }
            return _bResult;
        }

        /// <summary>
        /// in this function insert query build according to Arraylist and tablename
        /// and pass it to ExecuteNonQuery of ClsDal Class.
        /// if Query execute Successefuly and no rows updated then it return NOTDONE
        /// </summary>
        /// <param name="szTableName"></param>
        /// <param name="ArrWParameterList"></param>
        /// <returns>if Query Execute Successfuly then it return DONE or NOTDONE else Exception</returns>
        public bool Delete(string szTableName, ArrayList ArrWParameterList)
        {
            _bResult = true;
            _szbValues = new StringBuilder("");
            try
            {
                int iListCount = 0;
                iListCount = ArrWParameterList.Count;
                int iCnt;
                for (iCnt = 0; iCnt < iListCount; iCnt += 3)
                {
                    _szDatatype = null;
                    _szDatatype = ArrWParameterList[iCnt + 1].ToString().ToUpper();
                    //if Values Are Not None
                    if (_szbValues.ToString() == "")
                    {
                        switch (_szDatatype)
                        {
                            case "VARCHAR":
                                _szbValues.Append(ArrWParameterList[iCnt].ToString() + "'" + ArrWParameterList[iCnt + 2].ToString() + "'");
                                break;

                            case "NUMBER":
                            case "BIT":
                                _szbValues.Append(ArrWParameterList[iCnt].ToString() + ArrWParameterList[iCnt + 1].ToString() + ArrWParameterList[iCnt + 2].ToString());
                                break;

                            case "DATE":
                                if (DataBase.ToUpper() == "MSSQL")
                                    _szbValues.Append(ArrWParameterList[iCnt].ToString() + "CONVERT(varchar, '" + ArrWParameterList[iCnt + 2].ToString() + "', 101)");
                                else if (DataBase.ToUpper() == "ORACLE")
                                    _szbValues.Append(ArrWParameterList[iCnt].ToString() + "to_date('" + ArrWParameterList[iCnt + 2].ToString() + "', 'MM-DD-YYYY')");
                                break;
                            case "TIME":
                                if (DataBase.ToUpper() == "MSSQL")
                                    _szbValues.Append(ArrWParameterList[iCnt].ToString() + "CONVERT(varchar,'" + ArrWParameterList[iCnt + 2].ToString() + "', 22)");
                                else if (DataBase.ToUpper() == "ORACLE")
                                    _szbValues.Append(ArrWParameterList[iCnt].ToString() + "to_date('" + ArrWParameterList[iCnt + 2].ToString() + "', 'HH24:MI:SS')");
                                break;
                            case "DATETIME":
                                if (DataBase.ToUpper() == "MSSQL")
                                    _szbValues.Append(ArrWParameterList[iCnt].ToString() + "CONVERT(varchar, '" + ArrWParameterList[iCnt + 2].ToString() + "', 100)");
                                else if (DataBase.ToUpper() == "ORACLE")
                                    _szbValues.Append(ArrWParameterList[iCnt].ToString() + "to_date('" + ArrWParameterList[iCnt + 2].ToString() + "', 'MM-DD-YYYY HH24:MI:SS')");
                                break;
                            default:
                                _szbValues.Append(ArrWParameterList[iCnt].ToString() + ArrWParameterList[iCnt + 1].ToString() + ArrWParameterList[iCnt + 2].ToString());
                                break;
                        }
                    } //  if (szValues == "")
                    //if Values Are Not None
                    else
                    {
                        _szDatatype = ArrWParameterList[iCnt + 2].ToString().ToUpper();
                        switch (_szDatatype)
                        {
                            case "VARCHAR":
                                _szbValues.Append(ArrWParameterList[iCnt].ToString() + ArrWParameterList[iCnt + 1].ToString() + "'" + ArrWParameterList[iCnt + 3].ToString() + "'");
                                break;

                            case "NUMBER":
                            case "BIT":
                                _szbValues.Append(ArrWParameterList[iCnt].ToString() + ArrWParameterList[iCnt + 1].ToString() + ArrWParameterList[iCnt + 2].ToString() + ArrWParameterList[iCnt + 3].ToString());
                                break;

                            case "DATE":
                                if (DataBase.ToUpper() == "MSSQL")
                                    _szbValues.Append(ArrWParameterList[iCnt].ToString() + ArrWParameterList[iCnt + 1].ToString() + "CONVERT(varchar, '" + ArrWParameterList[iCnt + 3].ToString() + "', 101)");
                                else if (DataBase.ToUpper() == "ORACLE")
                                    _szbValues.Append(ArrWParameterList[iCnt].ToString() + ArrWParameterList[iCnt + 1].ToString() + " to_date('" + ArrWParameterList[iCnt + 3].ToString() + "', 'MM-DD-YYYY')");
                                break;

                            case "TIME":
                                if (DataBase.ToUpper() == "MSSQL")
                                    _szbValues.Append(ArrWParameterList[iCnt].ToString() + ArrWParameterList[iCnt + 1].ToString() + " CONVERT(varchar, '" + ArrWParameterList[iCnt + 3].ToString() + "', 22)");
                                else if (DataBase.ToUpper() == "ORACLE")
                                    _szbValues.Append(ArrWParameterList[iCnt].ToString() + ArrWParameterList[iCnt + 1].ToString() + " to_date('" + ArrWParameterList[iCnt + 3].ToString() + "', 'HH24:MI:SS')");
                                break;
                            case "DATETIME":
                                if (DataBase.ToUpper() == "MSSQL")
                                    _szbValues.Append(ArrWParameterList[iCnt].ToString() + ArrWParameterList[iCnt + 1].ToString() + "CONVERT(varchar, '" + ArrWParameterList[iCnt + 3].ToString() + "', 100)");
                                else if (DataBase.ToUpper() == "ORACLE")
                                    _szbValues.Append(ArrWParameterList[iCnt].ToString() + ArrWParameterList[iCnt + 1].ToString() + " to_date('" + ArrWParameterList[iCnt + 3].ToString() + "', 'MM-DD-YYYY HH24:MI:SS')");
                                break;
                            default:
                                _szbValues.Append(ArrWParameterList[iCnt].ToString() + ArrWParameterList[iCnt + 1].ToString() + ArrWParameterList[iCnt + 2].ToString() + ArrWParameterList[iCnt + 3].ToString());
                                break;
                        }
                        iCnt = iCnt + 1;
                    } //else  if (szValues == "")
                } //for (iCnt = 0; iCnt <= iArrCount; iCnt++)
                _szDatatype = null;

                _szQuery = " Delete from " + szTableName + " Where " + _szbValues.ToString();
                _szbValues = null;

                _bResult = bExecuteNonQuery(_szQuery);
            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = ex.Message;
            }
            finally
            {
                _szbValues = null;
                _szDatatype = null;
                _szQuery = null;
                ArrBolbData = null;
            }
            return _bResult;
        }

        /// <summary>
        /// This function used for Execution of Direct Query.
        /// e.g. if we insert record into a A table Selecting record form B table
        /// </summary>
        /// <param name="szQuery"></param>
        /// <returns></returns>
        public bool ExecuteQuery(string szQuery)
        {
            return bExecuteNonQuery(szQuery);
        }

        /// <summary>
        /// This function used for Execution of Query and returns no of affected records.
        /// </summary>
        /// <param name="szQuery">Insert, Update or Delete Query.</param>
        /// <returns>Number of records affected by SQL Query. If affected records is -1 means there is error while executive the Query.</returns>
        public int iExecuteQuery(string szQuery)
        {
            return ExecuteNonQuery(szQuery);
        }

        # endregion

        #region .... Select DataReader, Select DataSet ....

        /// <summary>
        /// This Function Build Query from Arrary list. 
        /// ArrParameterList contain field. ArrWParameterList Contains where condition
        /// </summary>
        /// <param name="szTableName"></param>
        /// <param name="ArrParameterList"></param>
        /// <param name="ArrWParameterList"></param>
        /// <returns>IDataReader</returns>
        public IDataReader DecideDatabaseDR(string szTableName, ArrayList ArrParameterList, ArrayList ArrWParameterList)
        {
            _szbFields = new StringBuilder("");
            _szbValues = new StringBuilder("");
            objDtReader = null;
            try
            {
                int iListCount = ArrParameterList.Count;
                int iCnt;
                for (iCnt = 0; iCnt <= iListCount - 1; iCnt++)
                {
                    if (_szbFields.ToString() == "")
                        _szbFields.Append(ArrParameterList[iCnt].ToString());
                    else
                        _szbFields.Append(", " + ArrParameterList[iCnt].ToString());
                }
                iListCount = 0;
                iListCount = ArrWParameterList.Count;

                if (ArrWParameterList[0].ToString() == "0")
                    _szQuery = " Select " + _szbFields.ToString() + " From " + szTableName;
                else
                {
                    iCnt = 0;
                    for (iCnt = 0; iCnt <= iListCount - 1; iCnt += 3)
                    {
                        _szDatatype = null;
                        //if Values Are Not None
                        if (_szbValues.ToString() == "")
                        {
                            _szDatatype = ArrWParameterList[iCnt + 1].ToString().ToUpper();
                            switch (_szDatatype)
                            {
                                case "VARCHAR":
                                    _szbValues.Append(ArrWParameterList[iCnt].ToString() + " '" + ArrWParameterList[iCnt + 2].ToString() + "'");
                                    break;

                                case "NUMBER":
                                case "BIT":
                                    _szbValues.Append(ArrWParameterList[iCnt].ToString() + " " + ArrWParameterList[iCnt + 1].ToString() + ArrWParameterList[iCnt + 2].ToString());
                                    break;

                                case "DATE":
                                    if (DataBase.ToUpper() == "MSSQL")
                                        _szbValues.Append(ArrWParameterList[iCnt].ToString() + " CONVERT(varchar, '" + ArrWParameterList[iCnt + 2].ToString() + "', 101)");
                                    else if (DataBase.ToUpper() == "ORACLE")
                                        _szbValues.Append(ArrWParameterList[iCnt].ToString() + " to_date('" + ArrWParameterList[iCnt + 2].ToString() + "', 'MM-DD-YYYY')");
                                    break;

                                case "TIME":
                                    if (DataBase.ToUpper() == "MSSQL")
                                        _szbValues.Append(ArrWParameterList[iCnt].ToString() + " CONVERT(varchar,'" + ArrWParameterList[iCnt + 2].ToString() + "', 22)");
                                    else if (DataBase.ToUpper() == "ORACLE")
                                        _szbValues.Append(ArrWParameterList[iCnt].ToString() + " to_date('" + ArrWParameterList[iCnt + 2].ToString() + "', 'HH24:MI:SS')");
                                    break;
                                case "DATETIME":
                                    if (DataBase.ToUpper() == "MSSQL")
                                        _szbValues.Append(ArrWParameterList[iCnt].ToString() + " CONVERT(varchar, '" + ArrWParameterList[iCnt + 2].ToString() + "', 100)");
                                    else if (DataBase.ToUpper() == "ORACLE")
                                        _szbValues.Append(ArrWParameterList[iCnt].ToString() + " to_date('" + ArrWParameterList[iCnt + 2].ToString() + "', 'MM-DD-YYYY HH24:MI:SS')");
                                    break;
                                default:
                                    _szbValues.Append(ArrWParameterList[iCnt].ToString() + " " + ArrWParameterList[iCnt + 1].ToString() + " " + ArrWParameterList[iCnt + 2].ToString());
                                    break;
                            }
                        } //  if (szValues == "")
                        //if Values Are Not None
                        else
                        {
                            _szDatatype = ArrWParameterList[iCnt + 2].ToString().ToUpper();
                            switch (_szDatatype)
                            {
                                case "VARCHAR":
                                    _szbValues.Append(" " + ArrWParameterList[iCnt].ToString() + " " + ArrWParameterList[iCnt + 1].ToString() + " '" + ArrWParameterList[iCnt + 3].ToString() + "'");
                                    break;

                                case "NUMBER":
                                case "BIT":
                                    _szbValues.Append(" " + ArrWParameterList[iCnt].ToString() + " " + ArrWParameterList[iCnt + 1].ToString() + " " + ArrWParameterList[iCnt + 2].ToString() + " " + ArrWParameterList[iCnt + 3].ToString());
                                    break;

                                case "DATE":
                                    if (DataBase.ToUpper() == "MSSQL")
                                        _szbValues.Append(" " + ArrWParameterList[iCnt].ToString() + " " + ArrWParameterList[iCnt + 1].ToString() + " CONVERT(varchar, '" + ArrWParameterList[iCnt + 3].ToString() + "', 101)");
                                    else if (DataBase.ToUpper() == "ORACLE")
                                        _szbValues.Append(" " + ArrWParameterList[iCnt].ToString() + " " + ArrWParameterList[iCnt + 1].ToString() + " to_date('" + ArrWParameterList[iCnt + 3].ToString() + "', 'MM-DD-YYYY')");
                                    break;

                                case "TIME":
                                    if (DataBase.ToUpper() == "MSSQL")
                                        _szbValues.Append(" " + ArrWParameterList[iCnt].ToString() + " " + ArrWParameterList[iCnt + 1].ToString() + " CONVERT(varchar, '" + ArrWParameterList[iCnt + 3].ToString() + "', 22)");
                                    else if (DataBase.ToUpper() == "ORACLE")
                                        _szbValues.Append(" " + ArrWParameterList[iCnt].ToString() + " " + ArrWParameterList[iCnt + 1].ToString() + " to_date('" + ArrWParameterList[iCnt + 3].ToString() + "', 'HH24:MI:SS')");
                                    break;
                                case "DATETIME":
                                    if (DataBase.ToUpper() == "MSSQL")
                                        _szbValues.Append(" " + ArrWParameterList[iCnt].ToString() + " " + ArrWParameterList[iCnt + 1].ToString() + " CONVERT(varchar, '" + ArrWParameterList[iCnt + 3].ToString() + "', 100)");
                                    else if (DataBase.ToUpper() == "ORACLE")
                                        _szbValues.Append(" " + ArrWParameterList[iCnt].ToString() + " " + ArrWParameterList[iCnt + 1].ToString() + " to_date('" + ArrWParameterList[iCnt + 3].ToString() + "', 'MM-DD-YYYY HH24:MI:SS')");
                                    break;
                                default:
                                    _szbValues.Append(" " + ArrWParameterList[iCnt].ToString() + ArrWParameterList[iCnt + 1].ToString() + ArrWParameterList[iCnt + 2].ToString() + ArrWParameterList[iCnt + 3].ToString());
                                    break;
                            }
                            iCnt = iCnt + 1;
                        } //else  if (szValues == "")
                    } //for (iCnt = 0; iCnt <= iArrCount; iCnt++)
                    _szDatatype = null;

                    _szQuery = " Select " + _szbFields.ToString() + " from " + szTableName + " Where " + _szbValues.ToString();
                }
                _szbFields = null;
                _szbValues = null;

                objDtReader = ExecuteReader(_szQuery);
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
                _szbFields = null;
                _szbValues = null;
                _szDatatype = null;
                _szQuery = null;
            }
            return objDtReader;
        }

        /// <summary>
        /// This functoin take select query and execute query and return Datareader
        /// </summary>
        /// <param name="szQuery"></param>
        /// <returns>IDataReader</returns>
        public IDataReader DecideDatabaseQDR(string szQuery)
        {
            return ExecuteReader(szQuery);
        }

        public IDataReader DecideDatabaseQDR(string szQuery, CommandBehavior eCommandBehavior)
        {
            return ExecuteReader(szQuery, eCommandBehavior);
        }

        /// <summary>
        /// This Function Build Query from Arrary list. 
        /// ArrParameterList contain field. ArrWParameterList Contains where condition
        /// </summary>
        /// <param name="szTableName"></param>
        /// <param name="ArrParameterList"></param>
        /// <param name="ArrWParameterList"></param>
        /// <returns>DataSet</returns>
        public DataSet DecideDatabaseDS(string szTableName, ArrayList ArrParameterList, ArrayList ArrWParameterList)
        {
            objDataset = null;
            _szbFields = new StringBuilder("");
            _szbValues = new StringBuilder("");
            try
            {
                int iListCount = ArrParameterList.Count;
                int iCnt;
                for (iCnt = 0; iCnt <= iListCount - 1; iCnt++)
                {
                    if (_szbFields.ToString() == "")
                        _szbFields.Append(ArrParameterList[iCnt].ToString());
                    else
                        _szbFields.Append(", " + ArrParameterList[iCnt].ToString());
                }
                iListCount = 0;
                iListCount = ArrWParameterList.Count;

                if (ArrWParameterList[0].ToString() == "0")
                    _szQuery = " Select " + _szbFields.ToString() + " From " + szTableName;
                else
                {
                    iCnt = 0;
                    for (iCnt = 0; iCnt <= iListCount - 1; iCnt += 3)
                    {
                        _szDatatype = null;
                        //if Values Are Not None
                        if (_szbValues.ToString() == "")
                        {
                            _szDatatype = ArrWParameterList[iCnt + 1].ToString().ToUpper();
                            switch (_szDatatype)
                            {
                                case "VARCHAR":
                                    _szbValues.Append(ArrWParameterList[iCnt].ToString() + " '" + ArrWParameterList[iCnt + 2].ToString() + "'");
                                    break;

                                case "NUMBER":
                                case "BIT":
                                    _szbValues.Append(ArrWParameterList[iCnt].ToString() + " " + ArrWParameterList[iCnt + 1].ToString() + " " + ArrWParameterList[iCnt + 2].ToString());
                                    break;

                                case "DATE":
                                    if (DataBase.ToUpper() == "MSSQL")
                                        _szbValues.Append(ArrWParameterList[iCnt].ToString() + " CONVERT(varchar, '" + ArrWParameterList[iCnt + 2].ToString() + "', 101)");
                                    else if (DataBase.ToUpper() == "ORACLE")
                                        _szbValues.Append(ArrWParameterList[iCnt].ToString() + " to_date('" + ArrWParameterList[iCnt + 2].ToString() + "', 'MM-DD-YYYY')");
                                    break;

                                case "TIME":
                                    if (DataBase.ToUpper() == "MSSQL")
                                        _szbValues.Append(ArrWParameterList[iCnt].ToString() + " CONVERT(varchar,'" + ArrWParameterList[iCnt + 2].ToString() + "', 22)");
                                    else if (DataBase.ToUpper() == "ORACLE")
                                        _szbValues.Append(ArrWParameterList[iCnt].ToString() + " to_date('" + ArrWParameterList[iCnt + 2].ToString() + "', 'HH24:MI:SS')");
                                    break;

                                case "DATETIME":
                                    if (DataBase.ToUpper() == "MSSQL")
                                        _szbValues.Append(ArrWParameterList[iCnt].ToString() + " CONVERT(varchar, '" + ArrWParameterList[iCnt + 2].ToString() + "', 100)");
                                    else if (DataBase.ToUpper() == "ORACLE")
                                        _szbValues.Append(ArrWParameterList[iCnt].ToString() + " to_date('" + ArrWParameterList[iCnt + 2].ToString() + "', 'MM-DD-YYYY HH24:MI:SS')");
                                    break;

                                default:
                                    _szbValues.Append(ArrWParameterList[iCnt].ToString() + " " + ArrWParameterList[iCnt + 1].ToString() + " " + ArrWParameterList[iCnt + 2].ToString());
                                    break;
                            }
                        } //  if (szValues == "")
                        //if Values Are Not None
                        else
                        {
                            _szDatatype = ArrWParameterList[iCnt + 2].ToString().ToUpper();
                            switch (_szDatatype)
                            {
                                case "VARCHAR":
                                    _szbValues.Append(" " + ArrWParameterList[iCnt].ToString() + " " + ArrWParameterList[iCnt + 1].ToString() + " '" + ArrWParameterList[iCnt + 3].ToString() + "'");
                                    break;

                                case "NUMBER":
                                case "BIT":
                                    _szbValues.Append(" " + ArrWParameterList[iCnt].ToString() + " " + ArrWParameterList[iCnt + 1].ToString() + " " + ArrWParameterList[iCnt + 2].ToString() + " " + ArrWParameterList[iCnt + 3].ToString());
                                    break;

                                case "DATE":
                                    if (DataBase.ToUpper() == "MSSQL")
                                        _szbValues.Append(" " + ArrWParameterList[iCnt].ToString() + " " + ArrWParameterList[iCnt + 1].ToString() + " CONVERT(varchar, '" + ArrWParameterList[iCnt + 3].ToString() + "', 101)");
                                    else if (DataBase.ToUpper() == "ORACLE")
                                        _szbValues.Append(" " + ArrWParameterList[iCnt].ToString() + " " + ArrWParameterList[iCnt + 1].ToString() + " to_date('" + ArrWParameterList[iCnt + 3].ToString() + "', 'MM-DD-YYYY')");
                                    break;

                                case "TIME":
                                    if (DataBase.ToUpper() == "MSSQL")
                                        _szbValues.Append(" " + ArrWParameterList[iCnt].ToString() + " " + ArrWParameterList[iCnt + 1].ToString() + " CONVERT(varchar, '" + ArrWParameterList[iCnt + 3].ToString() + "', 22)");
                                    else if (DataBase.ToUpper() == "ORACLE")
                                        _szbValues.Append(" " + ArrWParameterList[iCnt].ToString() + " " + ArrWParameterList[iCnt + 1].ToString() + " to_date('" + ArrWParameterList[iCnt + 3].ToString() + "', 'HH24:MI:SS')");
                                    break;
                                case "DATETIME":
                                    if (DataBase.ToUpper() == "MSSQL")
                                        _szbValues.Append(" " + ArrWParameterList[iCnt].ToString() + " " + ArrWParameterList[iCnt + 1].ToString() + " CONVERT(varchar, '" + ArrWParameterList[iCnt + 3].ToString() + "', 100)");
                                    else if (DataBase.ToUpper() == "ORACLE")
                                        _szbValues.Append(" " + ArrWParameterList[iCnt].ToString() + " " + ArrWParameterList[iCnt + 1].ToString() + " to_date('" + ArrWParameterList[iCnt + 3].ToString() + "', 'MM-DD-YYYY HH24:MI:SS')");
                                    break;
                                default:
                                    _szbValues.Append(" " + ArrWParameterList[iCnt].ToString() + " " + ArrWParameterList[iCnt + 1].ToString() + " " + ArrWParameterList[iCnt + 2].ToString() + " " + ArrWParameterList[iCnt + 3].ToString());
                                    break;
                            }
                            iCnt = iCnt + 1;
                        } //else  if (szValues == "")
                    } //for (iCnt = 0; iCnt <= iArrCount; iCnt++)
                    _szDatatype = null;

                    _szQuery = " Select " + _szbFields.ToString() + " From " + szTableName + " Where " + _szbValues.ToString();
                }
                _szbFields = null;
                _szbValues = null;

                objDataset = ExecuteDataSet(_szQuery);
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
                _szbFields = null;
                _szbValues = null;
                _szDatatype = null;
                _szQuery = null;
            }
            return objDataset;
        }

        /// <summary>
        /// This functoin take select query and execute query and return Dataset
        /// </summary>
        /// <param name="szQuery"></param>
        /// <returns>DataSet</returns>
        public DataSet DecideDatabaseQDS(string szQuery)
        {
            return ExecuteDataSet(szQuery);
        }

        /// <summary>
        /// This functoin take select query and execute query and return DataTable
        /// </summary>
        /// <param name="szQuery"></param>
        /// <returns>DataTable</returns>
        public DataTable DecideDatabaseQDT(string szQuery)
        {
            return ExecuteDataTable(szQuery);
        }

        /// <summary>
        /// This function used to Check record is exist or not?
        /// This function is only use for Select script.
        /// </summary>
        /// <param name="szQuery"></param>
        /// <returns>bool</returns>
        public bool IsRecordExist(string szQuery)
        {
            return ExecuteScalar_bool(szQuery);
        }

        /// <summary>
        /// This function used for executiong agreegate functions
        /// such as getting Max,Min,Sum etc. 
        /// </summary>
        /// <param name="szQuery"></param>
        /// <returns>int (default value is 0)</returns>
        public int GetSingleValue(string szQuery)
        {
            return ExecuteScalar(szQuery);
        }

        /// <summary>
        /// This function used to get single column value
        /// i.e.first column and first row value.
        /// </summary>
        /// <param name="szQuery"></param>
        /// <returns>string (default value is null-- Record not Exist)</returns>
        public object GetFirstColumnValue(string szQuery)
        {
            return ExecuteScalar_Object(szQuery);
        }

        /// <summary>
        /// This Function used for storing bolb data into file
        /// </summary>
        /// <param name="szTemplateFilepath"></param>
        /// <param name="szTableName"></param>
        /// <param name="szBlobColumnName">name of cloumn where blod data stored</param>
        /// <param name="szConstraint">where Condition</param>
        /// <returns></returns>
        public bool StoreBLOBIntoFile(string szTemplateFilepath, string szTableName, string szBlobColumnName, string szConstraint)
        {
            return ExecuteStoreBLOBIntoFile(szTemplateFilepath, szTableName, szBlobColumnName, szConstraint);
        }

        # endregion

        #region ..... Functions for IDisposable Interface .....

        #region Variable Declaration for Disposable Object ...

        private Component CompBuildQuery = new Component();
        private bool bDisposed = false;

        #endregion

        protected override void Dispose(bool bDisposing)
        {
            if (!this.bDisposed)
            {
                if (bDisposing)
                {
                    if (objDtReader != null)
                    {
                        objDtReader.Close();
                        objDtReader.Dispose();
                    }

                    if (objDataset != null)
                        objDataset.Dispose();

                    if (ArrBolbData != null)
                        ArrBolbData.Clear();

                    if (CompBuildQuery != null)
                        CompBuildQuery.Dispose();
                    CompBuildQuery = null;
                }

                objDtReader = null;
                objDataset = null;
                ArrBolbData = null;

                _szQuery = null;
                _szbFields = null;
                _szbValues = null;
                _szDatatype = null;
                msgError = null;

                bDisposed = true;

                base.Dispose(bDisposing);
            }
        }

        #endregion
    }
}