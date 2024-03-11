using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using DDLLCS;
using System.IO;
using eDocsDN_Common_LogWriter;

namespace eDocsDN_OpenXml_Operations
{

    //public enum Process { CL, TR, TR_4, CP, OD, DR, ADDVAR, ADS };

    public class clsUpdateCustomProperty
    {
        #region .... Variable Declaration ...
        ClsBuildQuery _objDAL = null;
        clsLogWriter _objCommonLog = null;
        IDataReader _objDataReader = null;

        string _szSqlQuery = string.Empty;
        string _szCodeDescription = string.Empty;

        object _objReturnVal = null;

        #endregion


        #region .... Property ....
        public string msgError { get; set; }
        public bool DebugLog { get; set; }
        public string FileName { get; set; }
        public string DCRNo { get; set; }
        public string Company { get; set; }
        public string Location { get; set; }
        public string Department { get; set; }
        public string DocType { get; set; }
        public string DocNumber { get; set; }
        public string DocVersion { get; set; }
        public string Label4 { get; set; }
        public string Label5 { get; set; }
        public string Label6 { get; set; }
        public string RefNo { get; set; }
        public string DDocNo { get; set; }
        public string AutherDate { get; set; }
        public string StatusDateFormat { get; set; }
        public string Issue_Date { get; set; }
        public string InProcessDocStatus { get; set; }
        //public Process CurruntProcess { get; set; }
        public bool isNGMP { get; set; }
        public bool isMigrated { get; set; }
        #endregion

        #region .... Constructor  ....
        public clsUpdateCustomProperty(ClsBuildQuery objDAL, clsLogWriter objCommonLog)
        {
            msgError = "";
            _objDAL = objDAL;
            _objCommonLog = objCommonLog;
            DebugLog = true;
        }

        public clsUpdateCustomProperty()
        {
            msgError = "";
            DebugLog = true;
        }

        #endregion

        #region .... Public Functions ...
        public bool UpdateCustumProperty(IDataReader objDataReader, bool isAttachment)
        {
            msgError = "";
            StringBuilder szOwnersName = null;
            List<CustomProperty> lstCustomProperty = new List<CustomProperty>();

            try
            {

                lstCustomProperty.Add(new CustomProperty() { PropertyName = "Author", PropertyValue = objDataReader["rohtua_cod"] is DBNull ? "" : Convert.ToString(objDataReader["rohtua_cod"]) });
                lstCustomProperty.Add(new CustomProperty() { PropertyName = "title", PropertyValue = objDataReader["eltit"] is DBNull ? "" : Convert.ToString(objDataReader["eltit"]) });
                lstCustomProperty.Add(new CustomProperty() { PropertyName = "Document Number", PropertyValue = DocNumber });
                lstCustomProperty.Add(new CustomProperty() { PropertyName = "Document Type", PropertyValue = DocType });
                lstCustomProperty.Add(new CustomProperty() { PropertyName = "Version Number", PropertyValue = DocVersion });
                lstCustomProperty.Add(new CustomProperty() { PropertyName = "Comments", PropertyValue = objDataReader["segnahc_cod"] is DBNull ? "" : Convert.ToString(objDataReader["segnahc_cod"]) });

                _szSqlQuery = "Select eltit,emanf,emanm,emanl from ZESPL_tsm_resu where upper(di_resu) = '" + objDataReader["renwo_cod"] + "'";
                _objDataReader = _objDAL.DecideDatabaseQDR(_szSqlQuery);
                if (_objDAL.msgError != "")
                    throw new Exception(_objDAL.msgError);

                if (_objDataReader.Read())
                {
                    szOwnersName = new StringBuilder(_objDataReader["eltit"] is DBNull ? " " : Convert.ToString(_objDataReader["eltit"]));
                    szOwnersName.Append(_objDataReader["emanf"] is DBNull ? " " : Convert.ToString(" " + _objDataReader["emanf"]) + " ");
                    szOwnersName.Append(_objDataReader["emanm"] is DBNull ? " " : Convert.ToString(" " + _objDataReader["emanm"]) + " ");
                    szOwnersName.Append(_objDataReader["emanl"] is DBNull ? " " : Convert.ToString(" " + _objDataReader["emanl"]));
                }
                if (_objDataReader != null)
                {
                    _objDataReader.Close();
                    _objDataReader.Dispose();
                    _objDataReader = null;
                }

                lstCustomProperty.Add(new CustomProperty() { PropertyName = "Manager", PropertyValue = szOwnersName.ToString() });
                lstCustomProperty.Add(new CustomProperty() { PropertyName = "Company", PropertyValue = Company });
                lstCustomProperty.Add(new CustomProperty() { PropertyName = "Location", PropertyValue = Location });
                lstCustomProperty.Add(new CustomProperty() { PropertyName = "Department", PropertyValue = Department });
                lstCustomProperty.Add(new CustomProperty() { PropertyName = "Label4", PropertyValue = Label4 });
                lstCustomProperty.Add(new CustomProperty() { PropertyName = "Label5", PropertyValue = Label5 });
                lstCustomProperty.Add(new CustomProperty() { PropertyName = "Label6", PropertyValue = Label6 });
                lstCustomProperty.Add(new CustomProperty() { PropertyName = "Doc_exp_Date", PropertyValue = objDataReader["ot_rud"] is DBNull ? "" : Convert.ToDateTime(objDataReader["ot_rud"]).ToString("dd/MM/yyyy") });
                lstCustomProperty.Add(new CustomProperty() { PropertyName = "Doc_eff_Date", PropertyValue = objDataReader["morf_rud"] is DBNull ? "" : Convert.ToDateTime(objDataReader["morf_rud"]).ToString("dd/MM/yyyy") });


                #region ... Location Name ....

                _szSqlQuery = "select edoc,csed_dc from zespl_tsm_edoc where upper(epyt) = 'LOC' And upper(edoc)='" + Convert.ToString(Location) + "'";
                _objDataReader = _objDAL.DecideDatabaseQDR(_szSqlQuery);
                if (_objDAL.msgError != "")
                    throw new Exception(_objDAL.msgError);
                if (_objDataReader.Read())
                {
                    lstCustomProperty.Add(new CustomProperty() { PropertyName = "Loc_Name", PropertyValue = _objDataReader["csed_dc"] is DBNull ? "" : Convert.ToString(_objDataReader["csed_dc"]) });
                }
                else
                {
                    lstCustomProperty.Add(new CustomProperty() { PropertyName = "Loc_Name", PropertyValue = "" });
                }

                if (_objDataReader != null)
                {
                    _objDataReader.Close();
                    _objDataReader.Dispose();
                    _objDataReader = null;
                }
                #endregion

                #region ... Company Description ...

                _szCodeDescription = Get_Code_Description(Company, "COMP");
                if (msgError != "")
                {
                    throw new Exception("Error while getting Company Desc : " + msgError);
                }
                lstCustomProperty.Add(new CustomProperty() { PropertyName = "Company_Desc", PropertyValue = _szCodeDescription });

                #endregion

                #region ... Location Description ...

                _szCodeDescription = Get_Code_Description(Location, "LOC");
                if (msgError != "")
                {
                    throw new Exception("Error while getting Location Desc : " + msgError);
                }
                lstCustomProperty.Add(new CustomProperty() { PropertyName = "Location_Desc", PropertyValue = _szCodeDescription });

                #endregion

                #region ... Department Description ...

                _szCodeDescription = Get_Code_Description(Department, "DPT");
                if (msgError != "")
                {
                    throw new Exception("Error while getting Department Desc : " + msgError);
                }
                lstCustomProperty.Add(new CustomProperty() { PropertyName = "Department_Desc", PropertyValue = _szCodeDescription });

                #endregion

                #region ... DocType Description ...

                _szCodeDescription = Get_Code_Description(DocType, "DOC");
                if (msgError != "")
                {
                    throw new Exception("Error while getting DocType Desc : " + msgError);
                }
                lstCustomProperty.Add(new CustomProperty() { PropertyName = "DocType_Desc", PropertyValue = _szCodeDescription });

                #endregion

                #region ... Label4 Description ...

                _szCodeDescription = Get_HRK_Code_Desc(Company, Location, Department, DocType, Label4, "LEBAL4");
                if (msgError != "")
                {
                    throw new Exception("Error while getting Label4 Desc : " + msgError);
                }
                lstCustomProperty.Add(new CustomProperty() { PropertyName = "Label4_Desc", PropertyValue = _szCodeDescription });

                #endregion

                #region ... Label5 Description ...

                _szCodeDescription = Get_HRK_Code_Desc(Company, Location, Department, DocType, Label5, "LEBAL5");
                if (msgError != "")
                {
                    throw new Exception("Error while getting Label5 Desc : " + msgError);
                }
                lstCustomProperty.Add(new CustomProperty() { PropertyName = "Label5_Desc", PropertyValue = _szCodeDescription });

                #endregion

                #region ... Label6 Description ...

                _szCodeDescription = Get_HRK_Code_Desc(Company, Location, Department, DocType, Label6, "LEBAL6");
                if (msgError != "")
                {
                    throw new Exception("Error while getting Label6 Desc : " + msgError);
                }
                lstCustomProperty.Add(new CustomProperty() { PropertyName = "Label6_Desc", PropertyValue = _szCodeDescription });

                #endregion


                lstCustomProperty.Add(new CustomProperty() { PropertyName = "Target Date", PropertyValue = "" });
                lstCustomProperty.Add(new CustomProperty() { PropertyName = "FileName", PropertyValue = Path.GetFileName(FileName) });
                lstCustomProperty.Add(new CustomProperty() { PropertyName = "ARF", PropertyValue = objDataReader["on_frc"] is DBNull ? "" : Convert.ToString(objDataReader["on_frc"]) });
                lstCustomProperty.Add(new CustomProperty() { PropertyName = "Author_Dt", PropertyValue = "" });

                for (int i = 1; i <= 10; i++)
                {
                    lstCustomProperty.Add(new CustomProperty() { PropertyName = "R" + i, PropertyValue = "" });
                    lstCustomProperty.Add(new CustomProperty() { PropertyName = "R" + i + "_Dt", PropertyValue = "" });
                    lstCustomProperty.Add(new CustomProperty() { PropertyName = "AddLoc" + i, PropertyValue = "" });
                    if (i <= 5)
                    {
                        lstCustomProperty.Add(new CustomProperty() { PropertyName = "A" + i, PropertyValue = "" });
                        lstCustomProperty.Add(new CustomProperty() { PropertyName = "A" + i + "_Dt", PropertyValue = "" });
                    }
                }

                lstCustomProperty.Add(new CustomProperty() { PropertyName = "Checked By", PropertyValue = "" });
                lstCustomProperty.Add(new CustomProperty() { PropertyName = "Issue Date", PropertyValue = Issue_Date == null ? "" : Issue_Date });
                lstCustomProperty.Add(new CustomProperty() { PropertyName = "Review Date", PropertyValue = objDataReader["ot_rud"] is DBNull ? "" : Convert.ToString(objDataReader["ot_rud"]) });
                lstCustomProperty.Add(new CustomProperty() { PropertyName = "AddIn_Changes", PropertyValue = "" });


                if (InProcessDocStatus == "") InProcessDocStatus = "Draft";
                lstCustomProperty.Add(new CustomProperty() { PropertyName = "Status", PropertyValue = InProcessDocStatus });
                lstCustomProperty.Add(new CustomProperty() { PropertyName = "Ref No", PropertyValue = RefNo });
                lstCustomProperty.Add(new CustomProperty() { PropertyName = "Change_Control_Number", PropertyValue = DDocNo });


                #region .... Info Values ....


                string szInfo = string.Empty;
                for (int i = 1; i <= 20; i++)
                {
                    szInfo = "info" + i;
                    lstCustomProperty.Add(new CustomProperty() { PropertyName = szInfo, PropertyValue = "" });
                }
                

                #endregion

                #region .... Label Data ....
                _szSqlQuery = "select eulav_lebal,elbairav_motsuc from zespl_atad_lebal where on_frc =" + DCRNo + " order by di_cer";
                _objDataReader = _objDAL.DecideDatabaseQDR(_szSqlQuery);
                if (_objDAL.msgError != "")
                {
                    throw new Exception("Error while getting record from Label data" + _objDAL.msgError);
                }
                while (_objDataReader.Read())
                {
                    if (!(_objDataReader["eulav_lebal"] is DBNull))
                    {
                        szInfo = _objDataReader["elbairav_motsuc"].ToString().ToLower();
                        lstCustomProperty.Add(new CustomProperty() { PropertyName = szInfo, PropertyValue = Convert.ToString(_objDataReader["eulav_lebal"]) });
                    }
                }
                if (_objDataReader != null)
                {
                    _objDataReader.Close();
                    _objDataReader.Dispose();
                    _objDataReader = null;
                }

                #endregion


                UpdateDocumentCustomProperty(lstCustomProperty);
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            finally
            {
                if (_objDataReader != null)
                {
                    _objDataReader.Close();
                    _objDataReader.Dispose();
                    _objDataReader = null;
                }

                lstCustomProperty = null;
            }
            return true;
        }

        //public bool UpdateCustumProperty(IDataReader objDataReader_frc, IDataReader objDataReader_Doc_Info, bool isAttachment)
        //{
        //    msgError = "";

        //    string szOwnersName = null;
        //    IDataReader objDrReviewer = null;
        //    IDataReader objDrApprover = null;
        //    List<CustomProperty> lstCustomProperty = new List<CustomProperty>();

        //    try
        //    {

        //        //IIf(IsDBNull(drDocInfo("no_degnahc")), drDocInfo("no_detaerc"), drDocInfo("no_degnahc"))

        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "Author", PropertyValue = objDataReader_frc["rohtua_cod"] is DBNull ? "" : Convert.ToString(objDataReader_frc["rohtua_cod"]) });
        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "Author_Dt", PropertyValue = AutherDate });
        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "title", PropertyValue = objDataReader_frc["eltit"] is DBNull ? "" : Convert.ToString(objDataReader_frc["eltit"]) });

        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "Document Number", PropertyValue = DocNumber });
        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "Document Type", PropertyValue = DocType });
        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "Version Number", PropertyValue = DocVersion });





        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "Company", PropertyValue = Company });
        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "Location", PropertyValue = Location });
        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "Department", PropertyValue = Department });
        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "Label4", PropertyValue = Label4 });
        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "Label5", PropertyValue = Label5 });
        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "Label6", PropertyValue = Label6 });

        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "Doc_exp_Date", PropertyValue = objDataReader_Doc_Info["td_pxe"] is DBNull ? "" : Convert.ToDateTime(objDataReader_Doc_Info["td_pxe"]).ToString(StatusDateFormat) });
        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "Doc_eff_Date", PropertyValue = objDataReader_Doc_Info["td_ffe"] is DBNull ? "" : Convert.ToDateTime(objDataReader_Doc_Info["td_ffe"]).ToString(StatusDateFormat) });
        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "Target Date", PropertyValue = objDataReader_Doc_Info["td_tegrat"] is DBNull ? "" : Convert.ToDateTime(objDataReader_Doc_Info["td_tegrat"]).ToString(StatusDateFormat) });
        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "Issue Date", PropertyValue = Issue_Date == null ? "" : Issue_Date });


        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "Change_Control_Number", PropertyValue = DDocNo });
        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "Keywords", PropertyValue = objDataReader_Doc_Info["sdrowyek_cod"] is DBNull ? "" : Convert.ToString(objDataReader_Doc_Info["sdrowyek_cod"]) });
        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "FileName", PropertyValue = Path.GetFileName(FileName) });
        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "ARF", PropertyValue = objDataReader_frc["on_frc"] is DBNull ? "" : Convert.ToString(objDataReader_frc["on_frc"]) });
        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "Comments", PropertyValue = objDataReader_Doc_Info["mmoc_giro"] is DBNull ? "" : Convert.ToString(objDataReader_Doc_Info["mmoc_giro"]) });


        //        #region .... Author/Owner ....

        //        szOwnersName = Get_Users_FullName(objDataReader_frc["renwo_cod"].ToString());
        //        if (szOwnersName == null)
        //        {
        //            if (DebugLog) _objCommonLog.writeToLog(_objCommonLog.szSurrKey + "-$-E-$-  Error occured while getting Owners Information from Master table...  " + msgError);
        //            msgError = "Error occured while getting Owners Information from Master table..:" + msgError;
        //            throw new Exception(msgError);
        //        }
        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "Manager", PropertyValue = szOwnersName });

        //        #endregion

        //        #region ... Location Name ....

        //        _szSqlQuery = "select edoc,csed_dc from zespl_tsm_edoc where upper(epyt) = 'LOC' And upper(edoc)='" + Convert.ToString(Location) + "'";
        //        _objDataReader = _objDAL.DecideDatabaseQDR(_szSqlQuery);
        //        if (_objDAL.msgError != "")
        //            throw new Exception(_objDAL.msgError);
        //        if (_objDataReader.Read())
        //        {
        //            lstCustomProperty.Add(new CustomProperty() { PropertyName = "Loc_Name", PropertyValue = _objDataReader["csed_dc"] is DBNull ? "" : Convert.ToString(_objDataReader["csed_dc"]) });
        //        }
        //        else
        //        {
        //            lstCustomProperty.Add(new CustomProperty() { PropertyName = "Loc_Name", PropertyValue = "" });
        //        }

        //        if (_objDataReader != null)
        //        {
        //            _objDataReader.Close();
        //            _objDataReader.Dispose();
        //            _objDataReader = null;
        //        }
        //        #endregion

        //        #region ... Company Description ...

        //        _szCodeDescription = Get_Code_Description(Company, "COMP");
        //        if (msgError != "")
        //        {
        //            _objCommonLog.writeToLog(_objCommonLog.szSurrKey + "-$-E-$- Error while getting Company Desc :  " + msgError);
        //            throw new Exception("Error while getting Company Desc : " + msgError);
        //        }
        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "Company_Desc", PropertyValue = _szCodeDescription });

        //        #endregion

        //        #region ... Location Description ...

        //        _szCodeDescription = Get_Code_Description(Location, "LOC");
        //        if (msgError != "")
        //        {
        //            _objCommonLog.writeToLog(_objCommonLog.szSurrKey + "-$-E-$- Error while getting Location Desc :  " + msgError);
        //            throw new Exception("Error while getting Location Desc : " + msgError);
        //        }
        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "Location_Desc", PropertyValue = _szCodeDescription });

        //        #endregion

        //        #region ... Department Description ...

        //        _szCodeDescription = Get_Code_Description(Department, "DPT");
        //        if (msgError != "")
        //        {
        //            _objCommonLog.writeToLog(_objCommonLog.szSurrKey + "-$-E-$- Error while getting Department Desc :  " + msgError);
        //            throw new Exception("Error while getting Department Desc : " + msgError);
        //        }
        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "Department_Desc", PropertyValue = _szCodeDescription });

        //        #endregion

        //        #region ... DocType Description ...

        //        _szCodeDescription = Get_Code_Description(DocType, "DOC");
        //        if (msgError != "")
        //        {
        //            _objCommonLog.writeToLog(_objCommonLog.szSurrKey + "-$-E-$- Error while getting DocType Desc :  " + msgError);
        //            throw new Exception("Error while getting DocType Desc : " + msgError);
        //        }
        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "DocType_Desc", PropertyValue = _szCodeDescription });

        //        #endregion

        //        #region ... Label4 Description ...

        //        _szCodeDescription = Get_HRK_Code_Desc(Company, Location, Department, DocType, Label4, "LEBAL4");
        //        if (msgError != "")
        //        {
        //            _objCommonLog.writeToLog(_objCommonLog.szSurrKey + "-$-E-$- Error while getting Label4 Desc :  " + msgError);
        //            throw new Exception("Error while getting Label4 Desc : " + msgError);
        //        }
        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "Label4_Desc", PropertyValue = _szCodeDescription });

        //        #endregion

        //        #region ... Label5 Description ...

        //        _szCodeDescription = Get_HRK_Code_Desc(Company, Location, Department, DocType, Label5, "LEBAL5");
        //        if (msgError != "")
        //        {
        //            _objCommonLog.writeToLog(_objCommonLog.szSurrKey + "-$-E-$- Error while getting Label5 Desc :  " + msgError);
        //            throw new Exception("Error while getting Label5 Desc : " + msgError);
        //        }
        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "Label5_Desc", PropertyValue = _szCodeDescription });

        //        #endregion

        //        #region ... Label6 Description ...

        //        _szCodeDescription = Get_HRK_Code_Desc(Company, Location, Department, DocType, Label6, "LEBAL6");
        //        if (msgError != "")
        //        {
        //            _objCommonLog.writeToLog(_objCommonLog.szSurrKey + "-$-E-$- Error while getting Label6 Desc :  " + msgError);
        //            throw new Exception("Error while getting Label6 Desc : " + msgError);
        //        }
        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "Label6_Desc", PropertyValue = _szCodeDescription });

        //        #endregion


        //        //... Review /Approve
        //        //Approver
        //        //if (Process.CP.Equals(CurruntProcess))
        //        //{

        //            lstCustomProperty.Add(new CustomProperty() { PropertyName = "Review Date", PropertyValue = objDataReader_Doc_Info["td_pxe"] is DBNull ? "" : Convert.ToDateTime(objDataReader_Doc_Info["td_pxe"]).ToString(StatusDateFormat) });

        //            int iReviewer_Approver = 0;
        //            StringBuilder sbReviewersName = null;
        //            StringBuilder sbApproversName = null;
        //            StringBuilder sbTime = null;
        //            string szTime = string.Empty;

        //            _szSqlQuery = "select d.epyt_resu,d.lvl_qes,d.di_resu,d.yek_cod,U.eltit ,U.emanf ,U.emanm ,U.emanl,U.tamrof_etad_resu ,D.td_sutats,D.mt_sutats,D.ot_detageled,D.codno_tnirp " +
        //               "from zespl_tsm_resu U,zespl_tsid_cod D where D.epyt_resu='A' And D.yek_cod = " + DCRNo + " And U.di_resu = D.di_resu order by D.on_rs";

        //            objDrApprover = _objDAL.DecideDatabaseQDR(_szSqlQuery);
        //            if (_objDAL.msgError != "")
        //            {
        //                if (DebugLog) _objCommonLog.writeToLog(_objCommonLog.szSurrKey + "-$-E-$- Error while getting Approve List " + _objDAL.msgError);
        //                msgError = "Error while getting Approve List  :" + _objDAL.msgError;
        //                throw new Exception(msgError);
        //            }


        //            while (objDrApprover.Read())
        //            {
        //                if (Convert.ToString(objDrApprover["codno_tnirp"]) == "True" || Convert.ToString(objDrApprover["codno_tnirp"]) == "1")
        //                {
        //                    if (!(objDrApprover["mt_sutats"] is DBNull))
        //                    {
        //                        szTime = Convert.ToDateTime(objDrApprover["mt_sutats"]).ToString("T");
        //                    }
        //                    if (objDrApprover["ot_detageled"] is DBNull)
        //                    {
        //                        sbApproversName = new StringBuilder(objDrApprover["eltit"] is DBNull ? "" : Convert.ToString(objDrApprover["eltit"]));
        //                        sbApproversName.Append(objDrApprover["emanf"] is DBNull ? "" : Convert.ToString(" " + objDrApprover["emanf"]));
        //                        sbApproversName.Append(objDrApprover["emanm"] is DBNull ? "" : Convert.ToString(" " + objDrApprover["emanm"]));
        //                        sbApproversName.Append(objDrApprover["emanl"] is DBNull ? "" : Convert.ToString(" " + objDrApprover["emanl"]));

        //                        lstCustomProperty.Add(new CustomProperty() { PropertyName = "A" + iReviewer_Approver, PropertyValue = sbApproversName.ToString() });
        //                        if (!(objDrApprover["td_sutats"] is DBNull))
        //                        {

        //                            sbTime = new StringBuilder(objDrApprover["td_sutats"] is DBNull ? "" : Convert.ToDateTime(objDrApprover["td_sutats"]).ToString(StatusDateFormat) + " " + szTime);
        //                            lstCustomProperty.Add(new CustomProperty() { PropertyName = "A" + iReviewer_Approver + "_Dt", PropertyValue = sbTime.ToString() });
        //                        }
        //                    }
        //                    else
        //                    {
        //                        szOwnersName = Get_Users_FullName(Convert.ToString(objDrApprover["ot_detageled"]));
        //                        if (szOwnersName == null)
        //                        {
        //                            if (DebugLog) _objCommonLog.writeToLog(_objCommonLog.szSurrKey + "-$-E-$- Error while getting Delegators information from User master for Reviewers  " + msgError);
        //                            msgError = "Error while getting Delegators information from User master for Approver   :" + msgError;
        //                            throw new Exception(msgError);
        //                        }
        //                        lstCustomProperty.Add(new CustomProperty() { PropertyName = "A" + iReviewer_Approver, PropertyValue = szOwnersName });
        //                        if (!(objDrApprover["td_sutats"] is DBNull))
        //                        {

        //                            sbTime = new StringBuilder(objDrApprover["td_sutats"] is DBNull ? "" : Convert.ToDateTime(objDrApprover["td_sutats"]).ToString(StatusDateFormat) + " " + szTime);
        //                            lstCustomProperty.Add(new CustomProperty() { PropertyName = "A" + iReviewer_Approver + "_Dt", PropertyValue = sbTime.ToString() });
        //                        }
        //                    }
        //                //}
        //            }
        //            if (objDrApprover != null)
        //            {
        //                objDrApprover.Close();
        //                objDrApprover.Dispose();
        //                objDrApprover = null;
        //            }
        //            sbApproversName = null;
        //            //Reviewer
        //            _szSqlQuery = "select d.epyt_resu,d.lvl_qes,d.di_resu,d.yek_cod,U.eltit ,U.emanf ,U.emanm ,U.emanl ,D.td_sutats,D.mt_sutats,D.ot_detageled,D.codno_tnirp,U.tamrof_etad_resu " +
        //                "from zespl_tsm_resu U,zespl_tsid_cod D " +
        //                "where D.epyt_resu='R' And D.yek_cod = " + DCRNo + " And U.di_resu = D.di_resu order by D.on_rs";

        //            objDrReviewer = _objDAL.DecideDatabaseQDR(_szSqlQuery);
        //            if (_objDAL.msgError != "")
        //            {
        //                if (DebugLog) _objCommonLog.writeToLog(_objCommonLog.szSurrKey + "-$-E-$- Error while getting Reviewer List " + _objDAL.msgError);
        //                msgError = "Error while getting Reviewer List  :" + _objDAL.msgError;
        //                throw new Exception(msgError);
        //            }

        //            while (objDrReviewer.Read())
        //            {
        //                if (Convert.ToString(objDrReviewer["codno_tnirp"]) == "True" || Convert.ToString(objDrReviewer["codno_tnirp"]) == "1")
        //                {
        //                    if (!(objDrReviewer["mt_sutats"] is DBNull))
        //                    {
        //                        szTime = Convert.ToDateTime(objDrReviewer["mt_sutats"]).ToString("T");
        //                    }
        //                    if (objDrReviewer["ot_detageled"] is DBNull)
        //                    {
        //                        sbReviewersName = new StringBuilder(objDrReviewer["eltit"] is DBNull ? "" : Convert.ToString(objDrReviewer["eltit"]));
        //                        sbReviewersName.Append(objDrReviewer["emanf"] is DBNull ? "" : Convert.ToString(" " + objDrReviewer["emanf"]));
        //                        sbReviewersName.Append(objDrReviewer["emanm"] is DBNull ? "" : Convert.ToString(" " + objDrReviewer["emanm"]));
        //                        sbReviewersName.Append(objDrReviewer["emanl"] is DBNull ? "" : Convert.ToString(" " + objDrReviewer["emanl"]));

        //                        lstCustomProperty.Add(new CustomProperty() { PropertyName = "R" + iReviewer_Approver, PropertyValue = sbReviewersName.ToString() });
        //                        if (!(objDrReviewer["td_sutats"] is DBNull))
        //                        {

        //                            sbTime = new StringBuilder(objDrReviewer["td_sutats"] is DBNull ? "" : Convert.ToDateTime(objDrReviewer["td_sutats"]).ToString(StatusDateFormat) + " " + szTime);
        //                            lstCustomProperty.Add(new CustomProperty() { PropertyName = "R" + iReviewer_Approver + "_Dt", PropertyValue = sbTime.ToString() });
        //                        }
        //                    }
        //                    else
        //                    {
        //                        szOwnersName = Get_Users_FullName(Convert.ToString(objDrReviewer["ot_detageled"]));
        //                        if (szOwnersName == null)
        //                        {
        //                            if (DebugLog) _objCommonLog.writeToLog(_objCommonLog.szSurrKey + "-$-E-$- Error while getting Delegators information from User master for Reviewers  " + msgError);
        //                            msgError = "Error while getting Delegators information from User master for Reviewers   :" + msgError;
        //                            throw new Exception(msgError);
        //                        }
        //                        lstCustomProperty.Add(new CustomProperty() { PropertyName = "R" + iReviewer_Approver, PropertyValue = szOwnersName });
        //                        if (!(objDrReviewer["td_sutats"] is DBNull))
        //                        {

        //                            sbTime = new StringBuilder(objDrReviewer["td_sutats"] is DBNull ? "" : Convert.ToDateTime(objDrReviewer["td_sutats"]).ToString(StatusDateFormat) + " " + szTime);
        //                            lstCustomProperty.Add(new CustomProperty() { PropertyName = "R" + iReviewer_Approver + "_Dt", PropertyValue = sbTime.ToString() });
        //                        }
        //                    }
        //                }
        //            }
        //            if (objDrReviewer != null)
        //            {
        //                objDrReviewer.Close();
        //                objDrReviewer.Dispose();
        //                objDrReviewer = null;
        //            }
        //            sbReviewersName = null;


        //            //...  Last R/A
        //            _szSqlQuery = "select U.eltit ,U.emanf ,U.emanm ,U.emanl ,D.td_sutats,D.mt_sutats,D.ot_detageled,U.tamrof_etad_resu " +
        //          "from zespl_tsm_resu U,zespl_tsid_cod D " +
        //          "where D.epyt_resu='A' And D.yek_cod = '" + DCRNo + "' and U.di_resu = D.di_resu order by D.lvl_qes desc";

        //            if (!_objDAL.IsRecordExist(_szSqlQuery))
        //            {
        //                if (_objDAL.msgError != "")
        //                {
        //                    if (DebugLog) _objCommonLog.writeToLog(_objCommonLog.szSurrKey + "-$-E-$- Error while Last Approvers Approvers List " + _objDAL.msgError);
        //                    msgError = "Error while Last Approvers Approvers List  :" + _objDAL.msgError;
        //                    throw new Exception(msgError);
        //                }
        //                if (isNGMP == false && isMigrated == false)
        //                {
        //                    if (DebugLog) _objCommonLog.writeToLog(_objCommonLog.szSurrKey + "-$-E-$- No Last Approvers Found");
        //                    msgError = "No Last Approvers Found";
        //                    throw new Exception(msgError);
        //                }
        //            }
        //            _objDataReader = _objDAL.DecideDatabaseQDR(_szSqlQuery);
        //            if (_objDAL.msgError != "")
        //            {
        //                if (DebugLog) _objCommonLog.writeToLog(_objCommonLog.szSurrKey + "-$-E-$- Error while Last Approvers Approvers List " + _objDAL.msgError);
        //                msgError = "Error while Last Approvers Approvers List  :" + _objDAL.msgError;
        //                throw new Exception(msgError);
        //            }
        //            if (_objDataReader.Read())
        //            {
        //                if (DebugLog) _objCommonLog.writeToLog(_objCommonLog.szSurrKey + "-$-D-$- Last Approvers Found " + _objDAL.msgError);
        //                if (!(_objDataReader["mt_sutats"] is DBNull))
        //                {
        //                    szTime = Convert.ToDateTime(_objDataReader["mt_sutats"]).ToString("T");
        //                }
        //                if (_objDataReader["ot_detageled"] is DBNull)
        //                {
        //                    sbApproversName = null;
        //                    sbApproversName = new StringBuilder(_objDataReader["eltit"] is DBNull ? "" : Convert.ToString(_objDataReader["eltit"]));
        //                    sbApproversName.Append(_objDataReader["emanf"] is DBNull ? "" : Convert.ToString(" " + _objDataReader["emanf"]));
        //                    sbApproversName.Append(_objDataReader["emanm"] is DBNull ? "" : Convert.ToString(" " + _objDataReader["emanm"]));
        //                    sbApproversName.Append(_objDataReader["emanl"] is DBNull ? "" : Convert.ToString(" " + _objDataReader["emanl"]));

        //                    lstCustomProperty.Add(new CustomProperty() { PropertyName = "Checked By", PropertyValue = sbApproversName.ToString() });
        //                }
        //                else
        //                {
        //                    szOwnersName = Get_Users_FullName(Convert.ToString(_objDataReader["ot_detageled"]));
        //                    if (szOwnersName == null)
        //                    {
        //                        if (DebugLog) _objCommonLog.writeToLog(_objCommonLog.szSurrKey + "-$-E-$- Error while getting Delegators information from User master for Reviewers  " + msgError);
        //                        msgError = "Error while getting Delegators information from User master for checked-by   :" + msgError;
        //                        throw new Exception(msgError);
        //                    }
        //                    lstCustomProperty.Add(new CustomProperty() { PropertyName = "A" + iReviewer_Approver, PropertyValue = szOwnersName });
        //                    if (!(objDrReviewer["td_sutats"] is DBNull))
        //                    {

        //                        sbTime = new StringBuilder(_objDataReader["td_sutats"] is DBNull ? "" : Convert.ToDateTime(_objDataReader["td_sutats"]).ToString(StatusDateFormat) + " " + szTime);
        //                        lstCustomProperty.Add(new CustomProperty() { PropertyName = "A" + iReviewer_Approver + "_Dt", PropertyValue = sbTime.ToString() });
        //                    }

        //                }

        //            }

        //            if (_objDataReader != null)
        //            {
        //                _objDataReader.Close();
        //                _objDataReader.Dispose();
        //                _objDataReader = null;
        //            }

        //            //...
        //            //...AddLoc
        //            IDataReader drLocations;
        //            iReviewer_Approver = 1;
        //            drLocations = GetLocations(DCRNo);
        //            while (drLocations.Read())
        //            {
        //                lstCustomProperty.Add(new CustomProperty() { PropertyName = "AddLoc" + iReviewer_Approver, PropertyValue = Convert.ToString(drLocations["noitacol"]) });
        //                iReviewer_Approver = iReviewer_Approver + 1;
        //            }
        //            if (drLocations != null)
        //            {
        //                drLocations.Close();
        //                drLocations.Dispose();
        //                drLocations = null;
        //            }
        //            iReviewer_Approver = 0;
        //            //..
        //        }
        //        else
        //        {

        //            lstCustomProperty.Add(new CustomProperty() { PropertyName = "Review Date", PropertyValue = objDataReader_frc["ot_rud"] is DBNull ? "" : Convert.ToDateTime(objDataReader_frc["ot_rud"]).ToString(StatusDateFormat) });
        //            lstCustomProperty.Add(new CustomProperty() { PropertyName = "Checked By", PropertyValue = "" });
        //            for (int i = 1; i <= 10; i++)
        //            {
        //                lstCustomProperty.Add(new CustomProperty() { PropertyName = "AddLoc" + i, PropertyValue = "" });
        //            }
        //            for (int i = 1; i <= 10; i++)
        //            {
        //                lstCustomProperty.Add(new CustomProperty() { PropertyName = "R" + i, PropertyValue = "" });
        //                lstCustomProperty.Add(new CustomProperty() { PropertyName = "R" + i + "_Dt", PropertyValue = "" });

        //                if (i <= 5)
        //                {
        //                    lstCustomProperty.Add(new CustomProperty() { PropertyName = "A" + i, PropertyValue = "" });
        //                    lstCustomProperty.Add(new CustomProperty() { PropertyName = "A" + i + "_Dt", PropertyValue = "" });
        //                }
        //            }
        //        }


        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "AddIn_Changes", PropertyValue = "" });
        //        if (InProcessDocStatus == "") InProcessDocStatus = "Draft";
        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "Status", PropertyValue = InProcessDocStatus });
        //        lstCustomProperty.Add(new CustomProperty() { PropertyName = "Ref No", PropertyValue = RefNo });



        //        #region .... Info Values ....


        //        string szInfo = string.Empty;
        //        for (int i = 1; i <= 20; i++)
        //        {
        //            szInfo = "info" + i;
        //            lstCustomProperty.Add(new CustomProperty() { PropertyName = szInfo, PropertyValue = "" });
        //        }
        //        if (DebugLog)
        //            _objCommonLog.writeToLog(_objCommonLog.szSurrKey + "-$-D-$- Udated Info Variables Successfully");

        //        #endregion

        //        #region .... Label Data ....
        //        _szSqlQuery = "select eulav_lebal,elbairav_motsuc from zespl_atad_lebal where on_frc =" + DCRNo + " order by di_cer";
        //        _objDataReader = _objDAL.DecideDatabaseQDR(_szSqlQuery);
        //        if (_objDAL.msgError != "")
        //        {
        //            if (DebugLog) _objCommonLog.writeToLog(_objCommonLog.szSurrKey + "-$-E-$- Error while getting record from Label data :" + _objDAL.msgError);
        //            throw new Exception("Error while getting record from Label data" + _objDAL.msgError);
        //        }
        //        while (_objDataReader.Read())
        //        {
        //            if (!(_objDataReader["eulav_lebal"] is DBNull))
        //            {
        //                szInfo = _objDataReader["elbairav_motsuc"].ToString().ToLower();
        //                lstCustomProperty.Add(new CustomProperty() { PropertyName = szInfo, PropertyValue = Convert.ToString(_objDataReader["eulav_lebal"]) });
        //            }
        //        }
        //        if (_objDataReader != null)
        //        {
        //            _objDataReader.Close();
        //            _objDataReader.Dispose();
        //            _objDataReader = null;
        //        }
        //        #endregion


        //        UpdateDocumentCustomProperty(lstCustomProperty);
        //    }
        //    catch (Exception ex)
        //    {
        //        msgError = ex.Message;
        //        _objCommonLog.writeToLog(_objCommonLog.szSurrKey + "-$-E-$- Error occured while Updating Custom Variables :  " + msgError);
        //    }
        //    finally
        //    {
        //        if (_objDataReader != null)
        //        {
        //            _objDataReader.Close();
        //            _objDataReader.Dispose();
        //            _objDataReader = null;
        //        }

        //        lstCustomProperty = null;
        //    }
        //    return true;
        //}

        public bool UpdateCustumProperty(List<CustomProperty> lstCustomProperty)
        {
            msgError = "";
            try
            {
                UpdateDocumentCustomProperty(lstCustomProperty);
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            finally
            {
                lstCustomProperty = null;
            }
            return true;
        }



        #endregion

        #region .... Private Functions ....

        private IDataReader GetLocations(string szDCRNo)
        {

            IDataReader objdr_Locations = null;
            _szSqlQuery = "select noitacol from zespl_col_lppa where on_frc = " + szDCRNo + " order by on_rs";
            objdr_Locations = _objDAL.DecideDatabaseQDR(_szSqlQuery);
            if (_objDAL.msgError != "")
            {
                throw new Exception("Error while getting Locations(ADD_LOC)" + _objDAL.msgError);
            }
            return objdr_Locations;
        }

        private string Get_Users_FullName(string szUserID)
        {

            StringBuilder szOwnersName = null;
            string szUsers_Full_Name = string.Empty;
            try
            {
                _szSqlQuery = "Select eltit,emanf,emanm,emanl from ZESPL_tsm_resu where upper(di_resu) = '" + szUserID + "'";
                _objDataReader = _objDAL.DecideDatabaseQDR(_szSqlQuery);
                if (_objDAL.msgError != "")
                    throw new Exception(_objDAL.msgError);

                if (_objDataReader.Read())
                {
                    szOwnersName = new StringBuilder(_objDataReader["eltit"] is DBNull ? " " : Convert.ToString(_objDataReader["eltit"]));
                    szOwnersName.Append(_objDataReader["emanf"] is DBNull ? " " : Convert.ToString(" " + _objDataReader["emanf"]) + " ");
                    szOwnersName.Append(_objDataReader["emanm"] is DBNull ? " " : Convert.ToString(" " + _objDataReader["emanm"]) + " ");
                    szOwnersName.Append(_objDataReader["emanl"] is DBNull ? " " : Convert.ToString("" + _objDataReader["emanl"]));
                    szUsers_Full_Name = szOwnersName.ToString();
                }
                if (_objDataReader != null)
                {
                    _objDataReader.Close();
                    _objDataReader.Dispose();
                    _objDataReader = null;
                }
            }
            catch (Exception ex)
            {
                szUsers_Full_Name = null;
                msgError = ex.Message;
            }
            finally
            {
                if (_objDataReader != null)
                {
                    _objDataReader.Close();
                    _objDataReader.Dispose();
                    _objDataReader = null;
                }
            }
            return szUsers_Full_Name;
        }


        internal bool UpdateDocumentCustomProperty(List<CustomProperty> lstCustomProperty)
        {
            bool bResult = true;
            try
            {
                using (eDocsDN_DocX.DocX document = eDocsDN_DocX.DocX.Load(FileName))
                {
                    foreach (var item in lstCustomProperty)
                    {
                        document.AddCustomProperty(new eDocsDN_DocX.CustomProperty(item.PropertyName, Convert.ToString(item.PropertyValue)));
                    }
                    document.SaveAs(FileName);
                }
            }
            catch (Exception ex)
            {
                bResult = false;
                msgError = ex.Message;
            }
            finally
            {
                lstCustomProperty = null;
            }
            return bResult;
        }

        private string Get_Code_Description(string szCode, string szType)
        {
            msgError = "";
            string szCodeDescription = string.Empty;
            try
            {
                _szSqlQuery = "select csed_dc from zespl_tsm_edoc where upper(edoc) ='" + szCode.ToUpper() + "' and epyt='" + szType.ToUpper() + "'";
                _objReturnVal = _objDAL.GetFirstColumnValue(_szSqlQuery);
                if (_objDAL.msgError != "")
                    throw new Exception(_objDAL.msgError);

                if (_objReturnVal != null)
                {
                    if (!(_objReturnVal is DBNull))
                        szCodeDescription = Convert.ToString(_objReturnVal);
                }
                _objReturnVal = null;
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
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
                            _szSqlQuery = "Select lebal_csed From zespl_tsm_edoc_lebal_tsuc Where lebal_type in ('Label4','LABEL4','LEBAL4','lebal4') And lebal_edoc='" + szCode.Trim() + "'";
                            break;
                        case "LEBAL5":
                            _szSqlQuery = "Select lebal_csed From zespl_tsm_edoc_lebal_tsuc Where lebal_type in ('Label5','LABEL5','LEBAL5','lebal5') And lebal_edoc='" + szCode.Trim() + "'";
                            break;
                        case "LEBAL6":
                            _szSqlQuery = "Select lebal_csed From zespl_tsm_edoc_lebal_tsuc Where lebal_type in ('Label6','LABEL6','LEBAL6','lebal6') And lebal_edoc='" + szCode.Trim() + "'";
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

                    _szSqlQuery = "SELECT csed_dc FROM zespl_ltd_tsm_edoc_krh Dtl, zespl_rdh_tsm_edoc_krh Hdr " +
                            " WHERE Hdr.yek_rrus_rdh = Dtl.yek_rrus_rdh AND Dtl.yek_rrus_rdh = (" + szHdrSurrKey + ")" +
                            " AND UPPER (Dtl.epyt) = '" + szType.ToUpper() +
                            "' AND UPPER (Dtl.edoc) = '" + szCode.ToUpper() + "'";
                }

                _objReturnVal = _objDAL.GetFirstColumnValue(_szSqlQuery);
                if (_objDAL.msgError != "")
                    throw new Exception("Error while Gettting Code Desc : " + _objDAL.msgError);

                if (_objReturnVal != null)
                {
                    if (!(_objReturnVal is DBNull))
                        szDesc = _objReturnVal.ToString();
                    else
                        szDesc = "NA";

                }
            }
            catch (Exception ex)
            {
                msgError = ex.Message;

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

                _szSqlQuery = "Select yek_rrus_rdh From zespl_redaeh_lebal_tsuc";
                bReturn = _objDAL.IsRecordExist(_szSqlQuery);
                if (_objDAL.msgError != "")
                    throw new Exception("Error while selecting record from zespl_redaeh_lebal_tsuc table" + _objDAL.msgError);

            }
            catch (Exception ex)
            {
                bReturn = false;
                msgError = ex.Message;
            }
            return bReturn;
        }

        private string GetDateFormat()
        {
            string szConfigDate = string.Empty;
            try
            {
                _szSqlQuery = "SELECT dt_ngis_ele FROM zespl_setalpmet_cod WHERE UPPER(ynapmoc) = '" + Company.ToUpper() +
                        "' AND UPPER(noitacol) = '" + Location.ToUpper() + "' AND UPPER(tnemtraped) = '" + Department.ToUpper() +
                        "' AND UPPER(epyt_cod) = '" + DocType.ToUpper() + "'";
                _objReturnVal = _objDAL.GetFirstColumnValue(_szSqlQuery);
                if (_objDAL.msgError != "")
                    throw new Exception(_objDAL.msgError);

                if (_objReturnVal != null)
                    szConfigDate = _objReturnVal.ToString();

                _objReturnVal = null;
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            return szConfigDate;
        }

        #endregion

    }
}
