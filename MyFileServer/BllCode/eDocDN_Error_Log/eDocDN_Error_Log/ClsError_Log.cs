using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DDLLCS;
using eDocsDN_ReadAppXml;

namespace eDocDN_Error_Log
{
    public class ClsError_Log
    {
        #region ..... Variable Declaration ....
        ClsBuildQuery _objDal = null;
        clsReadAppXml _objINI = null;
        Dictionary<string, string> dic_AppVariables = null;

        string _szSqlQuery = string.Empty;
        string _szAppXmlPath = string.Empty;
        string _szDBName = string.Empty;
        string _szCurrunt_Location = string.Empty;

        #endregion

        #region ..... Property ....
        public string msgError { get; set; }
        public List<LogRecord> ErrorList { get; set; }
        public bool bDebugLog { get; set; }
        public bool bErrorLog { get; set; }

        #endregion

        #region ..... Constructor ....
        public ClsError_Log(string szAppXmlPath)
        {
            msgError = string.Empty;
            _szAppXmlPath = szAppXmlPath;
            _objDal = new ClsBuildQuery(_szAppXmlPath);
            _objINI = new clsReadAppXml(_szAppXmlPath);
            _szCurrunt_Location = _objINI.GetCurrentLocation();
            _objINI.ReadLocationSetting(_szCurrunt_Location, "");
            dic_AppVariables = _objINI.AppVariables;
            if (dic_AppVariables != null)
            {
                bErrorLog = Convert.ToBoolean(dic_AppVariables["ErrorStatus"]);
                bDebugLog = Convert.ToBoolean(dic_AppVariables["DebugStatus"]);
            }
            dic_AppVariables = null;

        }

        #endregion

        #region ..... Public Functions ....
        public bool AddLog(string szProgType, string szProgKey, string szProgDesc, string szSessionId,
                          string szStartDate, string szStartTime, string szUserID, string szUserPwd, bool bError)
        {
            int iSurrogateKey = 0;
            bool blogResultreturn = true;
            try
            {
                if (_objDal != null)
                    _objDal.OpenConnection();

                _szSqlQuery = "Insert into zespl_redaeh_gol (epyt_gorp,yek_gorp,csed_gorp,di_noisses,di_resu,dwp_resu,td_trats,mt_trats,epyt_gol, td_dne, mt_dne) " +
                              "Values ('" + szProgType + "','" + szProgKey + "','" + szProgDesc + "','" + szSessionId + "','" + szUserID + "','" + szUserPwd + "','" + szStartDate + "','" + szStartTime + "'," + (bError == true ? "NULL" : "'E'") + ",'" + DateTime.Now.ToString("MM/dd/yyyy") + "','" + DateTime.Now.ToString("HH:mm:ss") + "') SELECT @@IDENTITY AS 'Identity'";
                iSurrogateKey = _objDal.GetSingleValue(_szSqlQuery);
                if (_objDal.msgError != "")
                    throw new Exception(_objDal.msgError);

                for (int iCnt = 0; iCnt < ErrorList.Count; iCnt++)
                {
                    _szSqlQuery = "Insert into zespl_liated_gol (yek_rrus_gol, on_rs, epyt_gol, txet_gol) " +
                                   "Values (" + iSurrogateKey + "," + iCnt + 1 + ",'" + ErrorList[iCnt].eLogType.ToString() + "','" + ErrorList[iCnt].LogText.Replace("'", "''") + "')";
                    if (!_objDal.ExecuteQuery(_szSqlQuery))
                        throw new Exception(_objDal.msgError);
                }
            }
            catch (Exception ex)
            {
                blogResultreturn = false;
                msgError = ex.Message;
            }
            finally
            {
                if (ErrorList != null)
                    ErrorList.Clear();
                ErrorList = null;

                if (_objDal != null)
                {
                    _objDal.CloseConnection();
                    _objDal.Dispose();
                }
                _objDal = null;

            }
            return blogResultreturn;

        }

        #endregion

        #region ..... Private Functions ....


        #endregion
    }
}
