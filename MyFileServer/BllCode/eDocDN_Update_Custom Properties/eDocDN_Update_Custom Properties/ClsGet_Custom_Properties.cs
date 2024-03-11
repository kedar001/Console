using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using DDLLCS;

namespace eDocDN_Update_Custom_Properties
{
    class ClsGet_Custom_Properties : ClsUpdate_Custom_Properties
    {
        #region .... Variable Declaration ....
        ClsBuildQuery _objDal = null;
        IDataReader _objDataReader_DCR_Info = null;
        IDataReader objDataReader = null;

        string _szQuery = string.Empty;

        bool _bResult = false;

        object _objReturnVal = null;

        #endregion


        public enum Process
        {
            Controller_Live = 0,
            Transfer_Document,
            ScanSign_Updation,
            Controller_Publish,
            Document_Recall,
            Obsolute_document
        }

        #region .... Constructor ....

        public ClsGet_Custom_Properties(ClsBuildQuery objDal, string szFilePath)
        {
            msgError = "";
            _objDal = objDal;
            FileName = szFilePath;
        }
        #endregion

        #region .... Properties ....

        public string InProcessDocStatus { get; set; }

        #endregion

        #region .... Public Functions ....
        public bool Get_Custom_variables(IDataReader objDataReader_Frc, Process eProcess)
        {
            msgError = "";
            _bResult = true;
            try
            {
                switch (eProcess)
                {
                    case Process.Controller_Live:

                        break;
                    case Process.Transfer_Document:
                        break;
                    case Process.ScanSign_Updation:
                        break;
                    case Process.Controller_Publish:
                        break;
                    case Process.Document_Recall:
                        break;
                    case Process.Obsolute_document:
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = ex.Message;
            }
            finally
            {

            }
            return _bResult;
        }
        #endregion

        #region .... Private Functions ...
        private void Custom_Properties()
        {
            int iCount = 0;
            try
            {

                using (ClsUpdate_Custom_Properties objUpdate_Cust = new ClsUpdate_Custom_Properties(FileName))
                {
                    objUpdate_Cust.lstCustom_Properties = new List<Custom_Property>();

                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "title", PropertyValue = _objDataReader_DCR_Info["eltit"] is DBNull ? "title" : Convert.ToString(_objDataReader_DCR_Info["eltit"]) });
                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Document Number", PropertyValue = _objDataReader_DCR_Info["on_cod_qer"] is DBNull ? "Document Number" : Convert.ToString(_objDataReader_DCR_Info["on_cod_qer"]) });
                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Document Type", PropertyValue = _objDataReader_DCR_Info["epyt_cod_qer"] is DBNull ? "Document Type" : Convert.ToString(_objDataReader_DCR_Info["epyt_cod_qer"]) });
                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Version Number", PropertyValue = _objDataReader_DCR_Info["rev_wen_cod_qer"] is DBNull ? "Version Number" : Convert.ToString(_objDataReader_DCR_Info["rev_wen_cod_qer"]) });
                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Comments", PropertyValue = _objDataReader_DCR_Info["segnahc_cod"] is DBNull ? "Comments" : Convert.ToString(_objDataReader_DCR_Info["segnahc_cod"]) });
                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Manager", PropertyValue = Get_Full_Name_of_User(Convert.ToString(_objDataReader_DCR_Info["renwo_cod"])) });
                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Company", PropertyValue = _objDataReader_DCR_Info["ynapmoc"] is DBNull ? "Company" : Convert.ToString(_objDataReader_DCR_Info["ynapmoc"]) });
                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Author", PropertyValue = _objDataReader_DCR_Info["rohtua_cod"] is DBNull ? "Author" : Convert.ToString(_objDataReader_DCR_Info["rohtua_cod"]) });
                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Doc_exp_Date", PropertyValue = _objDataReader_DCR_Info["ot_rud"] is DBNull ? "Doc_exp_Date" : Convert.ToString(_objDataReader_DCR_Info["ot_rud"]) });
                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Doc_eff_Date", PropertyValue = _objDataReader_DCR_Info["morf_rud"] is DBNull ? "Doc_eff_Date" : Convert.ToString(_objDataReader_DCR_Info["morf_rud"]) });
                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Location", PropertyValue = _objDataReader_DCR_Info["noitacol"] is DBNull ? "Location" : Convert.ToString(_objDataReader_DCR_Info["noitacol"]) });
                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Loc_Name", PropertyValue = Get_Code_Description(Convert.ToString(_objDataReader_DCR_Info["noitacol"]), "LOC") });
                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Department", PropertyValue = Convert.ToString(_objDataReader_DCR_Info["tnemtraped"]) });
                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Label4", PropertyValue = Convert.ToString(_objDataReader_DCR_Info["lebal4"]) });
                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Label5", PropertyValue = Convert.ToString(_objDataReader_DCR_Info["lebal5"]) });
                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Label6", PropertyValue = Convert.ToString(_objDataReader_DCR_Info["lebal6"]) });

                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Company_Desc", PropertyValue = Get_Code_Description(Convert.ToString(_objDataReader_DCR_Info["ynapmoc"]), "COMP") });
                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Location_Desc", PropertyValue = Get_Code_Description(Convert.ToString(_objDataReader_DCR_Info["noitacol"]), "LOC") });
                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Department_Desc", PropertyValue = Get_Code_Description(Convert.ToString(_objDataReader_DCR_Info["tnemtraped"]), "DPT") });
                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "DocType_Desc", PropertyValue = Get_Code_Description(Convert.ToString(_objDataReader_DCR_Info["epyt_cod_qer"]), "DOC") });
                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Label4_Desc", PropertyValue = Get_HRK_Code_Desc(Convert.ToString(_objDataReader_DCR_Info["ynapmoc"]), Convert.ToString(_objDataReader_DCR_Info["noitacol"]), Convert.ToString(_objDataReader_DCR_Info["tnemtraped"]), Convert.ToString(_objDataReader_DCR_Info["epyt_cod_qer"]), Convert.ToString(_objDataReader_DCR_Info["lebal4"]), "LEBAL4") });
                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Label5_Desc", PropertyValue = Get_HRK_Code_Desc(Convert.ToString(_objDataReader_DCR_Info["ynapmoc"]), Convert.ToString(_objDataReader_DCR_Info["noitacol"]), Convert.ToString(_objDataReader_DCR_Info["tnemtraped"]), Convert.ToString(_objDataReader_DCR_Info["epyt_cod_qer"]), Convert.ToString(_objDataReader_DCR_Info["lebal5"]), "LEBAL5") });
                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Label6_Desc", PropertyValue = Get_HRK_Code_Desc(Convert.ToString(_objDataReader_DCR_Info["ynapmoc"]), Convert.ToString(_objDataReader_DCR_Info["noitacol"]), Convert.ToString(_objDataReader_DCR_Info["tnemtraped"]), Convert.ToString(_objDataReader_DCR_Info["epyt_cod_qer"]), Convert.ToString(_objDataReader_DCR_Info["lebal6"]), "LEBAL6") });


                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Target Date", PropertyValue = "Target Date" });
                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "FileName", PropertyValue = "FileName" });
                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "ARF", PropertyValue = _objDataReader_DCR_Info["on_frc"] is DBNull ? "ARF" : Convert.ToString(_objDataReader_DCR_Info["on_frc"]) });
                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Author_Dt", PropertyValue = "Author_Dt" });

                    for (iCount = 1; iCount <= 10; iCount++)
                    {
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "R" + iCount, PropertyValue = "R" + iCount });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "R" + iCount + "_Dt", PropertyValue = "R" + iCount + "_Dt" });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "AddLoc" + iCount, PropertyValue = "AddLoc" + iCount });
                        if (iCount <= 5)
                        {
                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "A" + iCount, PropertyValue = "A" + iCount });
                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "A" + iCount + "_Dt", PropertyValue = "A" + iCount + "_Dt" });
                        }
                    }

                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Checked By", PropertyValue = "Checked By" });
                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Review Date", PropertyValue = "Review Date" });
                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "AddIn_Changes", PropertyValue = "AddIn_Changes" });
                    if (InProcessDocStatus == "")
                        InProcessDocStatus = "Draft";
                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Status", PropertyValue = InProcessDocStatus });
                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Ref No", PropertyValue = Convert.ToString(_objDataReader_DCR_Info["on_fer_ruoy"]) });
                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Change_Control_Number", PropertyValue = _objDataReader_DCR_Info["on_codd"] is DBNull ? "Change_Control_Number" : Convert.ToString(_objDataReader_DCR_Info["on_codd"]) });

                    iCount = 0;
                    string Info = string.Empty;
                    for (iCount = 1; iCount <= 20; iCount++)
                    {
                        Info = "info" + iCount;
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = Info, PropertyValue = Info });
                    }

                    //_szQuery = "select eulav_lebal,elbairav_motsuc from zespl_atad_lebal where on_frc =" & objARFDR("on_frc") & " order by di_cer";










                }
            }
            finally
            {

            }
        }


        private string Get_Full_Name_of_User(string szUserID)
        {

            StringBuilder szOwnerName = null;
            try
            {
                _szQuery = "Select eltit,emanf,emanm,emanl from ZESPL_tsm_resu where upper(di_resu) = '" + szUserID + "'";
                objDataReader = _objDal.DecideDatabaseQDR(_szQuery);
                if (objDataReader == null)
                {
                    msgError = "Error while Getting Record from User master table-->" + _objDal.msgError;
                    throw new Exception(msgError);
                }

                if (objDataReader.Read())
                {
                    szOwnerName = new StringBuilder(objDataReader["eltit"] is DBNull ? " " : Convert.ToString(objDataReader["eltit"]));
                    szOwnerName.Append(objDataReader["emanf"] is DBNull ? " " : Convert.ToString(" " + objDataReader["emanf"]) + " ");
                    szOwnerName.Append(objDataReader["emanm"] is DBNull ? " " : Convert.ToString(" " + objDataReader["emanm"]) + " ");
                    szOwnerName.Append(objDataReader["emanl"] is DBNull ? " " : Convert.ToString(" " + objDataReader["emanl"]));

                }

                objDataReader.Close();
                objDataReader.Dispose();
                objDataReader = null;
            }
            finally
            {
                if (objDataReader != null)
                {
                    objDataReader.Close();
                    objDataReader.Dispose();
                    objDataReader = null;
                }
            }
            return szOwnerName.ToString();
        }

        private string Get_Code_Description(string szCode, string szType)
        {
            objDataReader = null;
            string szCodeDescription = string.Empty;
            try
            {
                _szQuery = "select csed_dc from zespl_tsm_edoc where upper(edoc) ='" + szCode.ToUpper() + "' and epyt='" + szType.ToUpper() + "'";
                _objReturnVal = _objDal.GetFirstColumnValue(_szQuery);
                if (_objDal.msgError != "")
                    throw new Exception("Error while Gettting Code Desc : " + _objDal.msgError);
                if (_objReturnVal != null)
                {
                    szCodeDescription = Convert.ToString(_objReturnVal);
                }
                _objReturnVal = null;

            }
            finally
            {
                _objReturnVal = null;
            }
            return szCodeDescription;
        }

        private string Get_HRK_Code_Desc(string szComp, string szLoc, string szDept, string szDocType, string szCode, string szType)
        {

            string szDesc = szCode;
            string szHdrSurrKey = "";
            _objReturnVal = null;
            try
            {
                if (Is_Labels_ExistIn_Mapping())
                {
                    switch (szType)
                    {
                        case "LEBAL4":
                            _szQuery = "Select lebal_csed From zespl_tsm_edoc_lebal_tsuc Where lebal_type in ('Label4','LABEL4','LEBAL4','lebal4') And lebal_edoc='" + szCode.Trim() + "'";
                            break;
                        case "LEBAL5":
                            _szQuery = "Select lebal_csed From zespl_tsm_edoc_lebal_tsuc Where lebal_type in ('Label5','LABEL5','LEBAL5','lebal5') And lebal_edoc='" + szCode.Trim() + "'";
                            break;
                        case "LEBAL6":
                            _szQuery = "Select lebal_csed From zespl_tsm_edoc_lebal_tsuc Where lebal_type in ('Label6','LABEL6','LEBAL6','lebal6') And lebal_edoc='" + szCode.Trim() + "'";
                            break;
                        default:
                            break;
                    }
                }
                else
                {
                    szHdrSurrKey = "SELECT yek_rrus_rdh FROM zespl_rdh_tsm_edoc_krh " +
                                   "WHERE UPPER (ynapmoc) = '" + szComp.ToUpper() + "' AND UPPER (noitacol) = '" + szLoc.ToUpper() +
                                   "' AND UPPER (tnemtraped) = '" + szDept.ToUpper() + "' AND UPPER (epyt_cod) = '" + szDocType.ToUpper() + "'";

                    _szQuery = "SELECT csed_dc FROM zespl_ltd_tsm_edoc_krh Dtl, zespl_rdh_tsm_edoc_krh Hdr " +
                            " WHERE Hdr.yek_rrus_rdh = Dtl.yek_rrus_rdh AND Dtl.yek_rrus_rdh = (" + szHdrSurrKey + ")" +
                            " AND UPPER (Dtl.epyt) = '" + szType.ToUpper() +
                            "' AND UPPER (Dtl.edoc) = '" + szCode.ToUpper() + "'";
                }

                _objReturnVal = _objDal.GetFirstColumnValue(_szQuery);
                if (_objDal.msgError != "")
                    throw new Exception("Error while Gettting Code Desc : " + _objDal.msgError);

                if (_objReturnVal != null)
                {
                    if (!(_objReturnVal is DBNull))
                        szDesc = _objReturnVal.ToString();
                    else
                        szDesc = "NA";

                }
            }
            finally
            {
                _objReturnVal = null;
            }
            return szDesc;
        }
        private bool Is_Labels_ExistIn_Mapping()
        {
            msgError = "";
            bool bReturn = true;
            try
            {

                _szQuery = "Select yek_rrus_rdh From zespl_redaeh_lebal_tsuc";
                bReturn = _objDal.IsRecordExist(_szQuery);
                if (_objDal.msgError != "")
                    throw new Exception("Error while selecting record from zespl_redaeh_lebal_tsuc table" + _objDal.msgError);

            }
            finally
            {

            }
            return bReturn;
        }


        #endregion

        #region .... IDISPOSABLE ....

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {


            }
            else
            {
                if (lstCustom_Properties != null)
                {
                    lstCustom_Properties.Clear();
                    lstCustom_Properties = null;
                }
            }
        }

        ~ClsGet_Custom_Properties()
        {
            Dispose(false);
        }


        #endregion

    }
}
