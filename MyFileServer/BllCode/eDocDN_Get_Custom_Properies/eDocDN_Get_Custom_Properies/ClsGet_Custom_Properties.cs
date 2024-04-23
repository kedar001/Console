﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DDLLCS;
using eDocDN_Update_Custom_Properties;
using System.Data;
using eDocsDN_CommonFunctions;
using eDocsDN_Get_Directory_Info;
using System.IO;
using eDocsDN_ReadAppXml;



namespace eDocDN_Get_Custom_Properies
{
    public class ClsGet_Custom_Properties : ClsUpdate_Custom_Properties, IDisposable
    {
        #region .... Variable Declaration ....

        ClsBuildQuery _objDal = null;
        clsReadAppXml _objINI = null;
        ClsCommonFunction objComm;
        IDataReader _objDataReader_DCR_Info = null;
        IDataReader _objDataReader_Doc_Info = null;
        IDataReader _objDataReader_Doc_Dist = null;
        IDataReader _objDataReader_Issueance_Details = null;
        IDataReader _objDataReader = null;
        IDataReader _objCustomVariables = null;
        Dictionary<string, string> _dicUserName = null;
        int iCount = 0;
        string szCreatedAT = string.Empty;
        string szChangedAT = string.Empty;
        string szAuthDate = string.Empty;
        string szRevDate = string.Empty, szRevTime = string.Empty;
        string szDateFormat = string.Empty;
        string Info = string.Empty;
        ClsUpdate_Custom_Properties objUpdate_Cust = null;

        string _szQuery = string.Empty;
        string _szAppXmlPath = string.Empty;
        string _szDBName = string.Empty;
        bool _bisRevApproverExist = true;
        bool _bAutoGenerate_Sinature_Page = true;
        object _ObjReturnVal = null;


        object _objReturnVal = null;

        #endregion

        #region .... Constructor ....

        public ClsGet_Custom_Properties(ClsBuildQuery objDal, string szFilePath)
        {
            msgError = "";
            _objDal = objDal;
            FileName = szFilePath;
            _strmDocument = null;
        }
        public ClsGet_Custom_Properties(string szDBName, string szAppXmlPath, Stream strmDocument)
        {
            msgError = "";
            _szAppXmlPath = szAppXmlPath;
            _szDBName = szDBName;
            _objDal = new ClsBuildQuery(_szDBName, _szAppXmlPath);
            _objDal.OpenConnection(ConnectionFor.DocsExecutive, ConnectionType.NewConnection);
            FileName = null;
            _strmDocument = strmDocument;
        }
        public ClsGet_Custom_Properties(string szDBName, string szAppXmlPath, string szFilePath)
        {
            msgError = "";
            _szAppXmlPath = szAppXmlPath;
            _szDBName = szDBName;
            _objDal = new ClsBuildQuery(_szDBName, _szAppXmlPath);
            _objDal.OpenConnection(ConnectionFor.DocsExecutive, ConnectionType.NewConnection);
            _strmDocument = null;
            FileName = szFilePath;
        }

        #endregion

        #region .... Properties ....

        public new string msgError { get; set; }
        public string Document_Status { get; set; }
        public int DCRNo { get; set; }
        public bool isNGMP { get; set; }
        public bool isMigrated { get; set; }
        public string AppXmlPath { get; set; }
        public bool Batch_Card_Templete { get; set; }

        #endregion

        #region .... Public Functions ....
        public Stream Get_Custom_variables(int iDcrNo, Documents_Status eStatus, Documents_Process eProcess)
        {
            msgError = "";
            try
            {

                switch (eStatus)
                {
                    case Documents_Status.Draft:
                        using (_objINI = new clsReadAppXml(_szAppXmlPath))
                        {
                            Document_Status = _objINI.GetApplicationVariable("InProcessDocStatus");
                        }

                        break;
                    case Documents_Status.Draft_Approved:
                        using (_objINI = new clsReadAppXml(_szAppXmlPath))
                        {
                            Document_Status = _objINI.GetApplicationVariable("DraftApprovedStatus");
                        }

                        break;
                    case Documents_Status.Expired:
                        using (_objINI = new clsReadAppXml(_szAppXmlPath))
                        {
                            Document_Status = _objINI.GetApplicationVariable("ExpiredDocStatus");
                        }
                        break;
                    case Documents_Status.Issued:
                        using (_objINI = new clsReadAppXml(_szAppXmlPath))
                        {
                            Document_Status = _objINI.GetApplicationVariable("IssuedDocStatus");
                        }
                        break;
                    case Documents_Status.Obsolete:
                        using (_objINI = new clsReadAppXml(_szAppXmlPath))
                        {
                            Document_Status = _objINI.GetApplicationVariable("ObsoleteDocStatus");
                        }
                        break;
                    case Documents_Status.Publish:
                        using (_objINI = new clsReadAppXml(_szAppXmlPath))
                        {
                            Document_Status = _objINI.GetApplicationVariable("PublishDocStatus");
                        }
                        break;
                    default:
                        //Document_Status = "Draft";
                        break;
                }


                switch (eProcess)
                {
                    case Documents_Process.Controller_Live:

                        #region .... Get DCR Information .....

                        _szQuery = "select * from zespl_frc where on_frc=" + iDcrNo;
                        _objDataReader_DCR_Info = _objDal.DecideDatabaseQDR(_szQuery);
                        if (_objDataReader_DCR_Info == null)
                        {
                            throw new Exception(_objDal.msgError);
                        }

                        #endregion

                        break;
                    case Documents_Process.Transfer_Document:
                    case Documents_Process.Preview:

                        #region .... Get DCR Information .....

                        _szQuery = "select * from zespl_frc where on_frc=" + iDcrNo;
                        _objDataReader_DCR_Info = _objDal.DecideDatabaseQDR(_szQuery);
                        if (_objDataReader_DCR_Info == null)
                        {
                            throw new Exception(_objDal.msgError);
                        }

                        #endregion

                        #region .....  Get Document Information .....

                        _szQuery = "select * from ZESPL_ofni_cod where on_frc=" + iDcrNo;
                        _objDataReader_Doc_Info = _objDal.DecideDatabaseQDR(_szQuery);
                        if (_objDataReader_Doc_Info == null)
                        {
                            throw new Exception(_objDal.msgError);
                        }

                        #endregion

                        #region .... Get Rev/App Info ...

                        _szQuery = "select epyt_resu, lvl_qes, di_resu, td_sutats, mt_sutats, ot_detageled, codno_tnirp,noitangised,sutats from zespl_tsid_cod where epyt_resu in('R','A') And yek_cod = " + iDcrNo + "  order by lvl_qes";
                        _objDataReader_Doc_Dist = _objDal.DecideDatabaseQDR(_szQuery);
                        if (_objDataReader_Doc_Dist == null)
                            throw new Exception(_objDal.msgError);

                        _bisRevApproverExist = _objDal.IsRecordExist(_szQuery);
                        if (_objDal.msgError != "")
                            throw new Exception(_objDal.msgError);

                        #endregion


                        break;

                    case Documents_Process.Controller_Publish:
                    case Documents_Process.Document_Recall:

                        #region .... Get DCR Information .....

                        _szQuery = "select * from zespl_frc where on_frc=" + iDcrNo;
                        _objDataReader_DCR_Info = _objDal.DecideDatabaseQDR(_szQuery);
                        if (_objDataReader_DCR_Info == null)
                        {
                            throw new Exception(_objDal.msgError);
                        }

                        #endregion

                        #region .....  Get Document Information .....

                        _szQuery = "select * from ZESPL_ofni_cod where on_frc=" + iDcrNo;
                        _objDataReader_Doc_Info = _objDal.DecideDatabaseQDR(_szQuery);
                        if (_objDataReader_Doc_Info == null)
                        {
                            throw new Exception(_objDal.msgError);
                        }

                        #endregion

                        DCRNo = iDcrNo;
                        break;
                    case Documents_Process.Document_Issuance:

                        #region .... Issuance Details ....

                        //_szQuery = "select ynapmoc,noitacol,tnemtraped,epyt_cod,etad_eussi,yb_eussi,no_eussi,noitacol_eussi,tnemtraped_eussi,td_ffe from zespl_col_lppa A inner join zespl_ofni_cod D on A.on_frc_noitacol=D.on_frc  where on_frc_noitacol=" + iDcrNo;
                        _szQuery = "select S.ynapmoc,S.noitacol,S.tnemtraped,S.epyt_cod,etad_eussi,yb_eussi,no_eussi,noitacol_eussi,tnemtraped_eussi,D.td_ffe " +
                                 " from zespl_col_lppa A inner join zespl_ofni_cod D on A.on_frc_noitacol=D.on_frc" +
                                 " inner join  zespl_ofni_cod S on S.on_frc=A.on_frc where on_frc_noitacol=" + iDcrNo;
                        _objDataReader_Issueance_Details = _objDal.DecideDatabaseQDR(_szQuery);
                        if (_objDataReader_Issueance_Details == null)
                        {
                            throw new Exception(_objDal.msgError);
                        }

                        #endregion

                        break;
                    case Documents_Process.Attach_Custom_Variables:

                        DCRNo = iDcrNo;

                        _szQuery = "Select * from zespl_selbairav_motsuc order by on_rs";
                        _objCustomVariables = _objDal.DecideDatabaseQDR(_szQuery);
                        if (_objCustomVariables == null)
                            throw new Exception(_objDal.msgError);

                        break;
                    case Documents_Process.Attach_Custom_Variables_To_Template:

                        _szQuery = "Select * from zespl_selbairav_motsuc order by on_rs";
                        _objCustomVariables = _objDal.DecideDatabaseQDR(_szQuery);
                        if (_objCustomVariables == null)
                            throw new Exception(_objDal.msgError);

                        break;

                    case Documents_Process.TR4:

                        string szExt = string.Empty;
                        string szUserName = string.Empty;
                        string szRevTime = string.Empty;
                        string szRevDate = string.Empty;

                        #region .....  Get Document Information .....

                        _szQuery = "select * from ZESPL_ofni_cod where on_frc=" + iDcrNo;
                        _objDataReader_Doc_Info = _objDal.DecideDatabaseQDR(_szQuery);
                        if (_objDataReader_Doc_Info == null)
                        {
                            throw new Exception(_objDal.msgError);
                        }

                        #endregion

                        #region .... Get Rev/App Info ...

                        _szQuery = "select epyt_resu, lvl_qes, di_resu, td_sutats, mt_sutats, ot_detageled, codno_tnirp,noitangised from zespl_tsid_cod where epyt_resu in('R','A') And yek_cod = " + iDcrNo + "  order by lvl_qes";
                        _objDataReader_Doc_Dist = _objDal.DecideDatabaseQDR(_szQuery);
                        if (_objDataReader_Doc_Dist == null)
                            throw new Exception(_objDal.msgError);

                        #endregion

                        #region .... Auto Generate Signature Page ....
                        _szQuery = " select t.otua_erutangis_etareneg from zespl_setalpmet_cod T inner join zespl_ofni_cod D on T.ynapmoc=D.ynapmoc and T.noitacol=D.noitacol" +
                                  "  AND T.tnemtraped = D.tnemtraped and T.epyt_cod = D.epyt_cod" +
                                  "  WHERE d.on_frc =" + iDcrNo;

                        _ObjReturnVal = _objDal.GetFirstColumnValue(_szQuery);
                        if (_objDal.msgError != "")
                            throw new Exception(_objDal.msgError);

                        if (_objReturnVal != null)
                            _bAutoGenerate_Sinature_Page = Convert.ToBoolean(_objReturnVal);
                        _objReturnVal = null;
                        #endregion

                        break;

                    case Documents_Process.Expired_Document:

                        #region .....  Get Document Information .....

                        _szQuery = "select * from ZESPL_ofni_cod where on_frc=" + iDcrNo;
                        _objDataReader_Doc_Info = _objDal.DecideDatabaseQDR(_szQuery);
                        if (_objDataReader_Doc_Info == null)
                        {
                            throw new Exception(_objDal.msgError);
                        }

                        #endregion

                        break;
                    case Documents_Process.obsolete_Document:
                        break;
                    default:
                        break;

                }
                _strmDocument = Custom_Properties(eProcess);

            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            finally
            {
                if (_objDataReader_DCR_Info != null)
                {
                    _objDataReader_DCR_Info.Close();
                    _objDataReader_DCR_Info.Dispose();
                    _objDataReader_DCR_Info = null;
                }
                if (_objDataReader_Doc_Info != null)
                {
                    _objDataReader_Doc_Info.Close();
                    _objDataReader_Doc_Info.Dispose();
                    _objDataReader_Doc_Info = null;
                }
                if (_objDataReader_Issueance_Details != null)
                {
                    _objDataReader_Issueance_Details.Close();
                    _objDataReader_Issueance_Details.Dispose();
                    _objDataReader_Issueance_Details = null;
                }
                if (_objCustomVariables != null)
                {
                    _objCustomVariables.Close();
                    _objCustomVariables.Dispose();
                }
                if (_objDataReader_Doc_Dist != null)
                {
                    _objDataReader_Doc_Dist.Close();
                    _objDataReader_Doc_Dist.Dispose();
                }
                _objDataReader_Doc_Dist = null;

                _objCustomVariables = null;

                if (_objDataReader != null)
                {
                    _objDataReader.Close();
                    _objDataReader.Dispose();
                    _objDataReader = null;
                }

            }
            return _strmDocument;
        }
        #endregion

        #region .... Private Functions ...
        private Stream Custom_Properties(Documents_Process eProcess)
        {


            try
            {

                eDocsDN_BaseFunctions.ClsBaseFunctions objDate = null;
                if (_strmDocument != null)
                    objUpdate_Cust = new ClsUpdate_Custom_Properties(_strmDocument);
                else
                    objUpdate_Cust = new ClsUpdate_Custom_Properties(FileName);

                objUpdate_Cust.lstCustom_Properties = new List<Custom_Property>();


                switch (eProcess)
                {
                    case Documents_Process.Attach_Custom_Variables:

                        while (_objCustomVariables.Read())
                        {
                            objUpdate_Cust.lstCustom_Properties.Add(new Custom_Property { PropertyName = _objCustomVariables["rav_tsuc"].ToString(), PropertyValue = _objCustomVariables["eulav_rav_tsuc"].ToString() });
                        }

                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "ARF", PropertyValue = Convert.ToString(DCRNo) });


                        break;
                    case Documents_Process.Attach_Custom_Variables_To_Template:

                        while (_objCustomVariables.Read())
                        {
                            objUpdate_Cust.lstCustom_Properties.Add(new Custom_Property { PropertyName = _objCustomVariables["rav_tsuc"].ToString(), PropertyValue = _objCustomVariables["eulav_rav_tsuc"].ToString() });
                        }

                        break;

                    case Documents_Process.Controller_Live:
                        _objDataReader_DCR_Info.Read();

                        #region .... Custom Property ....

                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Doc_exp_Date", PropertyValue = _objDataReader_DCR_Info["ot_rud"] is DBNull ? "Doc_exp_Date" : Convert.ToString(_objDataReader_DCR_Info["ot_rud"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Doc_eff_Date", PropertyValue = _objDataReader_DCR_Info["morf_rud"] is DBNull ? "Doc_eff_Date" : Convert.ToString(_objDataReader_DCR_Info["morf_rud"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Target Date", PropertyValue = "Target Date" });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Change_Control_Number", PropertyValue = _objDataReader_DCR_Info["on_codd"] is DBNull ? "Change_Control_Number" : Convert.ToString(_objDataReader_DCR_Info["on_codd"]) });

                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Company_Desc", PropertyValue = Get_Code_Description(Convert.ToString(_objDataReader_DCR_Info["ynapmoc"]), "COMP") });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Location_Desc", PropertyValue = Get_Code_Description(Convert.ToString(_objDataReader_DCR_Info["noitacol"]), "LOC") });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Department_Desc", PropertyValue = Get_Code_Description(Convert.ToString(_objDataReader_DCR_Info["tnemtraped"]), "DPT") });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "DocType_Desc", PropertyValue = Get_Code_Description(Convert.ToString(_objDataReader_DCR_Info["epyt_cod_qer"]), "DOC") });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Label4_Desc", PropertyValue = Get_HRK_Code_Desc(Convert.ToString(_objDataReader_DCR_Info["ynapmoc"]), Convert.ToString(_objDataReader_DCR_Info["noitacol"]), Convert.ToString(_objDataReader_DCR_Info["tnemtraped"]), Convert.ToString(_objDataReader_DCR_Info["epyt_cod_qer"]), Convert.ToString(_objDataReader_DCR_Info["lebal4"]), "LEBAL4", Convert.ToString(_objDataReader_DCR_Info["lebal4"]), Convert.ToString(_objDataReader_DCR_Info["lebal5"])) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Label5_Desc", PropertyValue = Get_HRK_Code_Desc(Convert.ToString(_objDataReader_DCR_Info["ynapmoc"]), Convert.ToString(_objDataReader_DCR_Info["noitacol"]), Convert.ToString(_objDataReader_DCR_Info["tnemtraped"]), Convert.ToString(_objDataReader_DCR_Info["epyt_cod_qer"]), Convert.ToString(_objDataReader_DCR_Info["lebal5"]), "LEBAL5", Convert.ToString(_objDataReader_DCR_Info["lebal4"]), Convert.ToString(_objDataReader_DCR_Info["lebal5"])) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Label6_Desc", PropertyValue = Get_HRK_Code_Desc(Convert.ToString(_objDataReader_DCR_Info["ynapmoc"]), Convert.ToString(_objDataReader_DCR_Info["noitacol"]), Convert.ToString(_objDataReader_DCR_Info["tnemtraped"]), Convert.ToString(_objDataReader_DCR_Info["epyt_cod_qer"]), Convert.ToString(_objDataReader_DCR_Info["lebal6"]), "LEBAL6", Convert.ToString(_objDataReader_DCR_Info["lebal4"]), Convert.ToString(_objDataReader_DCR_Info["lebal5"])) });

                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Manager", PropertyValue = Get_Full_Name_of_User(Convert.ToString(_objDataReader_DCR_Info["renwo_cod"])) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Comments", PropertyValue = _objDataReader_DCR_Info["segnahc_cod"] is DBNull ? "Comments" : Convert.ToString(_objDataReader_DCR_Info["segnahc_cod"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Author_Dt", PropertyValue = "Author_Dt" });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Author_Desig", PropertyValue = "Author_Desig" });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Author_Dept", PropertyValue = "Author_Dept" });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "FileName", PropertyValue = "FileName" });

                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Title", PropertyValue = _objDataReader_DCR_Info["eltit"] is DBNull ? "Title" : Convert.ToString(_objDataReader_DCR_Info["eltit"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Document Number", PropertyValue = _objDataReader_DCR_Info["on_cod_qer"] is DBNull ? "Document Number" : Convert.ToString(_objDataReader_DCR_Info["on_cod_qer"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Document Type", PropertyValue = _objDataReader_DCR_Info["epyt_cod_qer"] is DBNull ? "Document Type" : Convert.ToString(_objDataReader_DCR_Info["epyt_cod_qer"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Version Number", PropertyValue = _objDataReader_DCR_Info["rev_wen_cod_qer"] is DBNull ? "Version Number" : Convert.ToString(_objDataReader_DCR_Info["rev_wen_cod_qer"]) });


                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "ARF", PropertyValue = _objDataReader_DCR_Info["on_frc"] is DBNull ? "ARF" : Convert.ToString(_objDataReader_DCR_Info["on_frc"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Company", PropertyValue = _objDataReader_DCR_Info["ynapmoc"] is DBNull ? "Company" : Convert.ToString(_objDataReader_DCR_Info["ynapmoc"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Author", PropertyValue = _objDataReader_DCR_Info["rohtua_cod"] is DBNull ? "Author" : Convert.ToString(_objDataReader_DCR_Info["rohtua_cod"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Location", PropertyValue = _objDataReader_DCR_Info["noitacol"] is DBNull ? "Location" : Convert.ToString(_objDataReader_DCR_Info["noitacol"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Loc_Name", PropertyValue = Get_Code_Description(Convert.ToString(_objDataReader_DCR_Info["noitacol"]), "LOC") });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Department", PropertyValue = Convert.ToString(_objDataReader_DCR_Info["tnemtraped"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Label4", PropertyValue = Convert.ToString(_objDataReader_DCR_Info["lebal4"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Label5", PropertyValue = Convert.ToString(_objDataReader_DCR_Info["lebal5"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Label6", PropertyValue = Convert.ToString(_objDataReader_DCR_Info["lebal6"]) });

                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Label6", PropertyValue = Convert.ToString(_objDataReader_DCR_Info["lebal6"]) });

                        #region ..... BTP/DR ....

                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Template_Author", PropertyValue = "Template_Author" });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Template_Author_Dt", PropertyValue = "Template_Author_Dt" });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Template_Author_Sign", PropertyValue = "Template_Author_Sign" });


                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Template_R1", PropertyValue = "Template_R1" });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Template_R1_Dt", PropertyValue = "Template_R1_Dt" });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Template_R1_Sign", PropertyValue = "Template_R1_Sign" });


                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Template_A1", PropertyValue = "Template_A1" });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Template_A1_Dt", PropertyValue = "Template_A1_Dt" });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Template_A1_Sign", PropertyValue = "Template_A1_Sign" });


                        #endregion

                        #endregion

                        break;
                    case Documents_Process.Transfer_Document:
                        objDate = new eDocsDN_BaseFunctions.ClsBaseFunctions(_objDal, AppXmlPath);

                        _objDataReader_DCR_Info.Read();
                        _objDataReader_Doc_Info.Read();

                        #region ... Get Date Format ....
                        //Get Date Format
                        szDateFormat = GetDateFormat(Convert.ToString(_objDataReader_Doc_Info["ynapmoc"]), Convert.ToString(_objDataReader_Doc_Info["noitacol"]), Convert.ToString(_objDataReader_Doc_Info["tnemtraped"]), Convert.ToString(_objDataReader_Doc_Info["epyt_cod"]));
                        if (szDateFormat != "")
                            szDateFormat = objDate.GetCodeDescription(szDateFormat.ToString(), "DT");
                        else
                            szDateFormat = objDate.GetCodeDescription("1", "DT");

                        #endregion

                        #region .... Get Author Date .....

                        if (!(_objDataReader_Doc_Info["ta_detaerc"] is DBNull))
                            szCreatedAT = Convert.ToDateTime(_objDataReader_Doc_Info["ta_detaerc"]).ToString("T");

                        if (!(_objDataReader_Doc_Info["ta_degnahc"] is DBNull))
                            szChangedAT = Convert.ToDateTime(_objDataReader_Doc_Info["ta_degnahc"]).ToString("T");

                        szAuthDate = _objDataReader_Doc_Info["no_degnahc"] is DBNull ? Convert.ToString(_objDataReader_Doc_Info["no_detaerc"]) : Convert.ToString(_objDataReader_Doc_Info["no_degnahc"]);

                        szAuthDate = Convert.ToDateTime(szAuthDate).ToString(szDateFormat);

                        if (szChangedAT != "")
                            szAuthDate = szAuthDate + " " + szChangedAT;
                        else
                            szAuthDate = szAuthDate + " " + szCreatedAT;

                        #endregion

                        #region .... Custom Properies ....

                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Keywords", PropertyValue = _objDataReader_Doc_Info["sdrowyek_cod"] is DBNull ? "Keywords" : Convert.ToString(_objDataReader_Doc_Info["sdrowyek_cod"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Comments", PropertyValue = _objDataReader_Doc_Info["mmoc_giro"] is DBNull ? "Comments" : Convert.ToString(_objDataReader_Doc_Info["mmoc_giro"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Manager", PropertyValue = Get_Full_Name_of_User(Convert.ToString(_objDataReader_DCR_Info["renwo_cod"])) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Author", PropertyValue = Get_Full_Name_of_User(Convert.ToString(_objDataReader_DCR_Info["rohtua_cod"])) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Company", PropertyValue = _objDataReader_DCR_Info["ynapmoc"] is DBNull ? "Company" : Convert.ToString(_objDataReader_DCR_Info["ynapmoc"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Location", PropertyValue = _objDataReader_DCR_Info["noitacol"] is DBNull ? "Location" : Convert.ToString(_objDataReader_DCR_Info["noitacol"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Department", PropertyValue = Convert.ToString(_objDataReader_DCR_Info["tnemtraped"]) });

                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Label4", PropertyValue = Convert.ToString(_objDataReader_Doc_Info["lebal4"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Label5", PropertyValue = Convert.ToString(_objDataReader_Doc_Info["lebal5"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Label6", PropertyValue = Convert.ToString(_objDataReader_Doc_Info["lebal6"]) });

                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Document Number", PropertyValue = _objDataReader_Doc_Info["on_cod"] is DBNull ? "Document Number" : Convert.ToString(_objDataReader_Doc_Info["on_cod"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Document Type", PropertyValue = _objDataReader_Doc_Info["epyt_cod"] is DBNull ? "Document Type" : Convert.ToString(_objDataReader_Doc_Info["epyt_cod"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Version Number", PropertyValue = _objDataReader_Doc_Info["rev_cod"] is DBNull ? "Version Number" : Convert.ToString(_objDataReader_Doc_Info["rev_cod"]) });

                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Company_Desc", PropertyValue = Get_Code_Description(Convert.ToString(_objDataReader_DCR_Info["ynapmoc"]), "COMP") });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Location_Desc", PropertyValue = Get_Code_Description(Convert.ToString(_objDataReader_DCR_Info["noitacol"]), "LOC") });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Department_Desc", PropertyValue = Get_Code_Description(Convert.ToString(_objDataReader_DCR_Info["tnemtraped"]), "DPT") });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "DocType_Desc", PropertyValue = Get_Code_Description(Convert.ToString(_objDataReader_DCR_Info["epyt_cod_qer"]), "DOC") });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Label4_Desc", PropertyValue = Get_HRK_Code_Desc(Convert.ToString(_objDataReader_DCR_Info["ynapmoc"]), Convert.ToString(_objDataReader_DCR_Info["noitacol"]), Convert.ToString(_objDataReader_DCR_Info["tnemtraped"]), Convert.ToString(_objDataReader_DCR_Info["epyt_cod_qer"]), Convert.ToString(_objDataReader_DCR_Info["lebal4"]), "LEBAL4", Convert.ToString(_objDataReader_DCR_Info["lebal4"]), Convert.ToString(_objDataReader_DCR_Info["lebal5"])) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Label5_Desc", PropertyValue = Get_HRK_Code_Desc(Convert.ToString(_objDataReader_DCR_Info["ynapmoc"]), Convert.ToString(_objDataReader_DCR_Info["noitacol"]), Convert.ToString(_objDataReader_DCR_Info["tnemtraped"]), Convert.ToString(_objDataReader_DCR_Info["epyt_cod_qer"]), Convert.ToString(_objDataReader_DCR_Info["lebal5"]), "LEBAL5", Convert.ToString(_objDataReader_DCR_Info["lebal4"]), Convert.ToString(_objDataReader_DCR_Info["lebal5"])) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Label6_Desc", PropertyValue = Get_HRK_Code_Desc(Convert.ToString(_objDataReader_DCR_Info["ynapmoc"]), Convert.ToString(_objDataReader_DCR_Info["noitacol"]), Convert.ToString(_objDataReader_DCR_Info["tnemtraped"]), Convert.ToString(_objDataReader_DCR_Info["epyt_cod_qer"]), Convert.ToString(_objDataReader_DCR_Info["lebal6"]), "LEBAL6", Convert.ToString(_objDataReader_DCR_Info["lebal4"]), Convert.ToString(_objDataReader_DCR_Info["lebal5"])) });

                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "ARF", PropertyValue = _objDataReader_DCR_Info["on_frc"] is DBNull ? "ARF" : Convert.ToString(_objDataReader_DCR_Info["on_frc"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "FileName", PropertyValue = "FileName" });

                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Title", PropertyValue = _objDataReader_Doc_Info["eltit_cod"] is DBNull ? "Title" : Convert.ToString(_objDataReader_Doc_Info["eltit_cod"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Doc_exp_Date", PropertyValue = _objDataReader_Doc_Info["td_pxe"] is DBNull ? "Doc_exp_Date" : Convert.ToDateTime(_objDataReader_Doc_Info["td_pxe"]).ToString(szDateFormat) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Doc_eff_Date", PropertyValue = _objDataReader_Doc_Info["td_ffe"] is DBNull ? "Doc_eff_Date" : Convert.ToDateTime(_objDataReader_Doc_Info["td_ffe"]).ToString(szDateFormat) });

                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Author_Dt", PropertyValue = szAuthDate });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Author_Desig", PropertyValue = "Author_Desig" });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Author_Dept", PropertyValue = "Author_Dept" });

                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Target Date", PropertyValue = _objDataReader_Doc_Info["td_tegrat"] is DBNull ? "Target Date" : Convert.ToDateTime(_objDataReader_Doc_Info["td_tegrat"]).ToString(szDateFormat) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Change_Control_Number", PropertyValue = _objDataReader_Doc_Info["on_codd"] is DBNull ? "Change_Control_Number" : Convert.ToString(_objDataReader_Doc_Info["on_codd"]) });



                        #endregion

                        break;
                    case Documents_Process.Preview:
                        int iApp = 0, iRev = 0;
                        szAuthDate = "Author_Dt";
                        objDate = new eDocsDN_BaseFunctions.ClsBaseFunctions(_objDal, AppXmlPath);

                        _objDataReader_DCR_Info.Read();
                        _objDataReader_Doc_Info.Read();

                        #region ... Get Date Format ....
                        //Get Date Format
                        szDateFormat = GetDateFormat(Convert.ToString(_objDataReader_Doc_Info["ynapmoc"]), Convert.ToString(_objDataReader_Doc_Info["noitacol"]), Convert.ToString(_objDataReader_Doc_Info["tnemtraped"]), Convert.ToString(_objDataReader_Doc_Info["epyt_cod"]));
                        if (szDateFormat != "")
                            szDateFormat = objDate.GetCodeDescription(szDateFormat.ToString(), "DT");
                        else
                            szDateFormat = objDate.GetCodeDescription("1", "DT");

                        #endregion

                        #region .... Get Author Date .....

                        if (!(_objDataReader_Doc_Info["ta_detaerc"] is DBNull))
                            szCreatedAT = Convert.ToDateTime(_objDataReader_Doc_Info["ta_detaerc"]).ToString("T");

                        if (!(_objDataReader_Doc_Info["ta_degnahc"] is DBNull))
                            szChangedAT = Convert.ToDateTime(_objDataReader_Doc_Info["ta_degnahc"]).ToString("T");

                        if (!DBNull.Value.Equals(_objDataReader_Doc_Info["no_detaerc"]))
                        {
                            szAuthDate = _objDataReader_Doc_Info["no_degnahc"] is DBNull ? Convert.ToString(_objDataReader_Doc_Info["no_detaerc"]) : Convert.ToString(_objDataReader_Doc_Info["no_degnahc"]);

                            szAuthDate = Convert.ToDateTime(szAuthDate).ToString(szDateFormat);

                            if (szChangedAT != "")
                                szAuthDate = szAuthDate + " " + szChangedAT;
                            else
                                szAuthDate = szAuthDate + " " + szCreatedAT;
                        }
                        #endregion

                        #region .... Custom Properies ....

                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Keywords", PropertyValue = _objDataReader_Doc_Info["sdrowyek_cod"] is DBNull ? "Keywords" : Convert.ToString(_objDataReader_Doc_Info["sdrowyek_cod"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Comments", PropertyValue = _objDataReader_Doc_Info["mmoc_giro"] is DBNull ? "Comments" : Convert.ToString(_objDataReader_Doc_Info["mmoc_giro"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Title", PropertyValue = _objDataReader_Doc_Info["eltit_cod"] is DBNull ? "Title" : Convert.ToString(_objDataReader_Doc_Info["eltit_cod"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Doc_exp_Date", PropertyValue = _objDataReader_Doc_Info["td_pxe"] is DBNull ? "Doc_exp_Date" : Convert.ToDateTime(_objDataReader_Doc_Info["td_pxe"]).ToString(szDateFormat) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Doc_eff_Date", PropertyValue = _objDataReader_Doc_Info["td_ffe"] is DBNull ? "Doc_eff_Date" : Convert.ToDateTime(_objDataReader_Doc_Info["td_ffe"]).ToString(szDateFormat) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Target Date", PropertyValue = _objDataReader_Doc_Info["td_tegrat"] is DBNull ? "Target Date" : Convert.ToDateTime(_objDataReader_Doc_Info["td_tegrat"]).ToString(szDateFormat) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Label4_Desc", PropertyValue = Get_HRK_Code_Desc(Convert.ToString(_objDataReader_DCR_Info["ynapmoc"]), Convert.ToString(_objDataReader_DCR_Info["noitacol"]), Convert.ToString(_objDataReader_DCR_Info["tnemtraped"]), Convert.ToString(_objDataReader_DCR_Info["epyt_cod_qer"]), Convert.ToString(_objDataReader_DCR_Info["lebal4"]), "LEBAL4", Convert.ToString(_objDataReader_DCR_Info["lebal4"]), Convert.ToString(_objDataReader_DCR_Info["lebal5"])) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Label5_Desc", PropertyValue = Get_HRK_Code_Desc(Convert.ToString(_objDataReader_DCR_Info["ynapmoc"]), Convert.ToString(_objDataReader_DCR_Info["noitacol"]), Convert.ToString(_objDataReader_DCR_Info["tnemtraped"]), Convert.ToString(_objDataReader_DCR_Info["epyt_cod_qer"]), Convert.ToString(_objDataReader_DCR_Info["lebal5"]), "LEBAL5", Convert.ToString(_objDataReader_DCR_Info["lebal4"]), Convert.ToString(_objDataReader_DCR_Info["lebal5"])) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Label6_Desc", PropertyValue = Get_HRK_Code_Desc(Convert.ToString(_objDataReader_DCR_Info["ynapmoc"]), Convert.ToString(_objDataReader_DCR_Info["noitacol"]), Convert.ToString(_objDataReader_DCR_Info["tnemtraped"]), Convert.ToString(_objDataReader_DCR_Info["epyt_cod_qer"]), Convert.ToString(_objDataReader_DCR_Info["lebal6"]), "LEBAL6", Convert.ToString(_objDataReader_DCR_Info["lebal4"]), Convert.ToString(_objDataReader_DCR_Info["lebal5"])) });

                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Change_Control_Number", PropertyValue = _objDataReader_DCR_Info["on_codd"] is DBNull ? "Change_Control_Number" : Convert.ToString(_objDataReader_DCR_Info["on_codd"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Company_Desc", PropertyValue = Get_Code_Description(Convert.ToString(_objDataReader_DCR_Info["ynapmoc"]), "COMP") });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Location_Desc", PropertyValue = Get_Code_Description(Convert.ToString(_objDataReader_DCR_Info["noitacol"]), "LOC") });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Department_Desc", PropertyValue = Get_Code_Description(Convert.ToString(_objDataReader_DCR_Info["tnemtraped"]), "DPT") });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "DocType_Desc", PropertyValue = Get_Code_Description(Convert.ToString(_objDataReader_DCR_Info["epyt_cod_qer"]), "DOC") });

                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Manager", PropertyValue = Get_Full_Name_of_User(Convert.ToString(_objDataReader_DCR_Info["renwo_cod"])) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "FileName", PropertyValue = "FileName" });

                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Document Number", PropertyValue = _objDataReader_DCR_Info["on_cod_qer"] is DBNull ? "Document Number" : Convert.ToString(_objDataReader_DCR_Info["on_cod_qer"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Document Type", PropertyValue = _objDataReader_DCR_Info["epyt_cod_qer"] is DBNull ? "Document Type" : Convert.ToString(_objDataReader_DCR_Info["epyt_cod_qer"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Version Number", PropertyValue = _objDataReader_DCR_Info["rev_wen_cod_qer"] is DBNull ? "Version Number" : Convert.ToString(_objDataReader_DCR_Info["rev_wen_cod_qer"]) });


                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Company", PropertyValue = _objDataReader_DCR_Info["ynapmoc"] is DBNull ? "Company" : Convert.ToString(_objDataReader_DCR_Info["ynapmoc"]) });

                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Location", PropertyValue = _objDataReader_DCR_Info["noitacol"] is DBNull ? "Location" : Convert.ToString(_objDataReader_DCR_Info["noitacol"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Loc_Name", PropertyValue = Get_Code_Description(Convert.ToString(_objDataReader_DCR_Info["noitacol"]), "LOC") });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Department", PropertyValue = Convert.ToString(_objDataReader_DCR_Info["tnemtraped"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Label4", PropertyValue = Convert.ToString(_objDataReader_DCR_Info["lebal4"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Label5", PropertyValue = Convert.ToString(_objDataReader_DCR_Info["lebal5"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Label6", PropertyValue = Convert.ToString(_objDataReader_DCR_Info["lebal6"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "ARF", PropertyValue = _objDataReader_DCR_Info["on_frc"] is DBNull ? "ARF" : Convert.ToString(_objDataReader_DCR_Info["on_frc"]) });

                        switch (Convert.ToString(_objDataReader_Doc_Info["sutats"]))
                        {
                            case "18":
                                _objINI = new clsReadAppXml(_szAppXmlPath);
                                objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Status", PropertyValue = _objDataReader_Doc_Info["sutats"] is DBNull ? "Status" : _objINI.GetApplicationVariable("DraftApprovedStatus") });
                                _objINI = null;
                                break;
                            default:
                                _objINI = new clsReadAppXml(_szAppXmlPath);
                                if (Document_Status.Equals(_objINI.GetApplicationVariable("DraftApprovedStatus")))
                                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Status", PropertyValue = _objDataReader_Doc_Info["sutats"] is DBNull ? "Status" : _objINI.GetApplicationVariable("DraftApprovedStatus") });
                                else
                                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Status", PropertyValue = _objDataReader_Doc_Info["sutats"] is DBNull ? "Status" : _objINI.GetApplicationVariable("InProcessDocStatus") });
                                _objINI = null;
                                objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Author", PropertyValue = Get_Full_Name_of_User(Convert.ToString(_objDataReader_DCR_Info["rohtua_cod"])) });
                                objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Author_Dt", PropertyValue = szAuthDate });
                                objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Author_Desig", PropertyValue = _objDataReader_Doc_Info["on_frc"] is DBNull ? "Author_Desig" : Get_Code_Description_For_Author(Convert.ToString(_objDataReader_DCR_Info["on_frc"]), "DESG") });
                                objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Author_Dept", PropertyValue = _objDataReader_Doc_Info["on_frc"] is DBNull ? "Author_Dept" : Get_Code_Description_For_Author(Convert.ToString(_objDataReader_DCR_Info["on_frc"]), "DPT") });
                                break;
                        }
                        #endregion

                        if (_bisRevApproverExist)
                        {
                            while (_objDataReader_Doc_Dist.Read())
                            {
                                if (Convert.ToBoolean(_objDataReader_Doc_Dist["codno_tnirp"]))
                                {
                                    if (_objDataReader_Doc_Dist["sutats"] is DBNull || !_objDataReader_Doc_Dist["sutats"].ToString().Equals("8"))
                                    {
                                        if (_objDataReader_Doc_Dist["epyt_resu"].ToString() == "A")
                                        {
                                            iApp++;
                                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = _objDataReader_Doc_Dist["epyt_resu"].ToString() + iApp.ToString() + "_Desig", PropertyValue = "XXXX XXXX XXXX" });
                                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = _objDataReader_Doc_Dist["epyt_resu"].ToString() + iApp.ToString() + "_Dept", PropertyValue = "XXXX XXXX XXXX" });
                                        }
                                        else
                                        {
                                            iRev++;
                                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = _objDataReader_Doc_Dist["epyt_resu"].ToString() + iApp.ToString() + "_Desig", PropertyValue = "XXXX XXXX XXXX" });
                                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = _objDataReader_Doc_Dist["epyt_resu"].ToString() + iApp.ToString() + "_Dept", PropertyValue = "XXXX XXXX XXXX" });
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            for (iCount = 1; iCount <= 10; iCount++)
                            {
                                objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "R" + iCount + "_Desig", PropertyValue = "XXX XXX XXX XXX XXX XXX" });
                                objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "R" + iCount + "_Dept", PropertyValue = "XXX XXX XXX XXX XXX XXX" });

                                if (iCount <= 5)
                                {
                                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "A" + iCount + "_Desig", PropertyValue = "XXX XXX XXX XXX XXX XXX" });
                                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "A" + iCount + "_Dept", PropertyValue = "XXX XXX XXX XXX XXX XXX" });
                                }
                            }
                        }


                        //_objINI = new clsReadAppXml(_szAppXmlPath);
                        //if (!Convert.ToString(_objDataReader_Doc_Info["sutats"]).Equals("18") || !Document_Status.Equals(_objINI.GetApplicationVariable("DraftApprovedStatus")))
                        //{
                        //    for (iCount = 1; iCount <= 10; iCount++)
                        //    {

                        //        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "R" + iCount + "_Desig", PropertyValue = "XXX XXX XXX XXX XXX XXX" });
                        //        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "R" + iCount + "_Dept", PropertyValue = "XXX XXX XXX XXX XXX XXX" });

                        //        if (iCount <= 5)
                        //        {
                        //            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "A" + iCount + "_Desig", PropertyValue = "XXX XXX XXX XXX XXX XXX" });
                        //            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "A" + iCount + "_Dept", PropertyValue = "XXX XXX XXX XXX XXX XXX" });

                        //        }
                        //    }
                        //}
                        //_objINI = null;
                        break;

                    case Documents_Process.Document_Issuance:

                        if (_objDataReader_Issueance_Details.Read())
                        {
                            objDate = new eDocsDN_BaseFunctions.ClsBaseFunctions(_objDal, AppXmlPath);

                            #region ... Get Date Format ....
                            //Get Date Format
                            szDateFormat = GetDateFormat(Convert.ToString(_objDataReader_Issueance_Details["ynapmoc"]), Convert.ToString(_objDataReader_Issueance_Details["noitacol"]), Convert.ToString(_objDataReader_Issueance_Details["tnemtraped"]), Convert.ToString(_objDataReader_Issueance_Details["epyt_cod"]));
                            if (szDateFormat != "")
                                szDateFormat = objDate.GetCodeDescription(szDateFormat.ToString(), "DT");
                            else
                                szDateFormat = objDate.GetCodeDescription("1", "DT");


                            if (!(_objDataReader_Issueance_Details["no_eussi"] is DBNull))
                                szCreatedAT = Convert.ToDateTime(_objDataReader_Issueance_Details["no_eussi"]).ToString("T");

                            #endregion

                            if (!string.IsNullOrEmpty(Document_Status))
                                objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Status", PropertyValue = Document_Status });
                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Issue_Date", PropertyValue = ConvertToDateTime("Document expiry date", Convert.ToString(_objDataReader_Issueance_Details["etad_eussi"])).ToString(szDateFormat) });
                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Issue_By", PropertyValue = Get_Full_Name_of_User(_objDataReader_Issueance_Details["yb_eussi"].ToString()) });
                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Issue_On", PropertyValue = ConvertToDateTime("Document expiry date", Convert.ToString(_objDataReader_Issueance_Details["no_eussi"])).ToString(szDateFormat) + " " + szCreatedAT });
                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Issue_Location", PropertyValue = Convert.ToString(_objDataReader_Issueance_Details["noitacol_eussi"]) });
                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Issue_Department", PropertyValue = Convert.ToString(_objDataReader_Issueance_Details["tnemtraped_eussi"]) });
                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Issued_From_Department", PropertyValue = Convert.ToString(_objDataReader_Issueance_Details["tnemtraped"]) });
                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Issued_From_Location", PropertyValue = Convert.ToString(_objDataReader_Issueance_Details["noitacol"]) });
                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Loc_Eff_Date", PropertyValue = _objDataReader_Issueance_Details["td_ffe"] is DBNull ? "Loc_Eff_Date" : ConvertToDateTime("Document effective date", Convert.ToString(_objDataReader_Issueance_Details["td_ffe"])).ToString(szDateFormat) });

                        }
                        _objDataReader_Issueance_Details.Close();
                        _objDataReader_Issueance_Details.Dispose();
                        _objDataReader_Issueance_Details = null;
                        szCreatedAT = string.Empty;

                        break;
                    case Documents_Process.Controller_Publish:

                        string szOwnersName = string.Empty;
                        string szTime = string.Empty;
                        int iReviewer_Approver = 1;
                        string szApproversName = null;
                        StringBuilder sbTime = null;

                        objDate = new eDocsDN_BaseFunctions.ClsBaseFunctions(_objDal, AppXmlPath);

                        _objDataReader_DCR_Info.Read();
                        _objDataReader_Doc_Info.Read();

                        #region ... Get Date Format ....
                        //Get Date Format
                        szDateFormat = GetDateFormat(Convert.ToString(_objDataReader_Doc_Info["ynapmoc"]), Convert.ToString(_objDataReader_Doc_Info["noitacol"]), Convert.ToString(_objDataReader_Doc_Info["tnemtraped"]), Convert.ToString(_objDataReader_Doc_Info["epyt_cod"]));
                        if (szDateFormat != "")
                            szDateFormat = objDate.GetCodeDescription(szDateFormat.ToString(), "DT");
                        else
                            szDateFormat = objDate.GetCodeDescription("1", "DT");

                        #endregion

                        #region .... Get Author Date .....

                        if (!(_objDataReader_Doc_Info["ta_detaerc"] is DBNull))
                            szCreatedAT = Convert.ToDateTime(_objDataReader_Doc_Info["ta_detaerc"]).ToString("T");

                        if (!(_objDataReader_Doc_Info["ta_degnahc"] is DBNull))
                            szChangedAT = Convert.ToDateTime(_objDataReader_Doc_Info["ta_degnahc"]).ToString("T");

                        szAuthDate = _objDataReader_Doc_Info["no_degnahc"] is DBNull ? Convert.ToString(_objDataReader_Doc_Info["no_detaerc"]) : Convert.ToString(_objDataReader_Doc_Info["no_degnahc"]);

                        szAuthDate = Convert.ToDateTime(szAuthDate).ToString(szDateFormat);

                        if (szChangedAT != "")
                            szAuthDate = szAuthDate + " " + szChangedAT;
                        else
                            szAuthDate = szAuthDate + " " + szCreatedAT;

                        #endregion

                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Doc_exp_Date", PropertyValue = _objDataReader_Doc_Info["td_pxe"] is DBNull ? "Doc_exp_Date" : ConvertToDateTime("Document expiry date", Convert.ToString(_objDataReader_Doc_Info["td_pxe"])).ToString(szDateFormat) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Doc_eff_Date", PropertyValue = _objDataReader_Doc_Info["td_ffe"] is DBNull ? "Doc_eff_Date" : ConvertToDateTime("Document effective date", Convert.ToString(_objDataReader_Doc_Info["td_ffe"])).ToString(szDateFormat) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Target Date", PropertyValue = _objDataReader_Doc_Info["td_tegrat"] is DBNull ? "Target Date" : ConvertToDateTime("Document target date", Convert.ToString(_objDataReader_Doc_Info["td_tegrat"])).ToString(szDateFormat) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Review Date", PropertyValue = _objDataReader_Doc_Info["td_pxe"] is DBNull ? "Review Date" : ConvertToDateTime("Document expiry date", Convert.ToString(_objDataReader_Doc_Info["td_pxe"])).ToString(szDateFormat) });

                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Title", PropertyValue = _objDataReader_Doc_Info["eltit_cod"] is DBNull ? "Title" : Convert.ToString(_objDataReader_Doc_Info["eltit_cod"]) });

                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Comments", PropertyValue = _objDataReader_Doc_Info["mmoc_giro"] is DBNull ? "Comments" : Convert.ToString(_objDataReader_Doc_Info["mmoc_giro"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "ARF", PropertyValue = _objDataReader_DCR_Info["on_frc"] is DBNull ? "ARF" : Convert.ToString(_objDataReader_DCR_Info["on_frc"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Change_Control_Number", PropertyValue = _objDataReader_DCR_Info["on_codd"] is DBNull ? "Change_Control_Number" : Convert.ToString(_objDataReader_DCR_Info["on_codd"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Keywords", PropertyValue = _objDataReader_Doc_Info["sdrowyek_cod"] is DBNull ? "Keywords" : Convert.ToString(_objDataReader_Doc_Info["sdrowyek_cod"]) });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Manager", PropertyValue = Get_Full_Name_of_User(Convert.ToString(_objDataReader_DCR_Info["renwo_cod"])) });



                        #region ... Last Reviewer/Approver

                        //...  Last R/A
                        _szQuery = "select U.di_resu,U.eltit ,U.emanf ,U.emanm ,U.emanl ,D.td_sutats,D.mt_sutats,D.ot_detageled,U.tamrof_etad_resu,D.noitangised " +
                      "from zespl_tsm_resu U,zespl_tsid_cod D " +
                      "where D.epyt_resu='A' And D.yek_cod = " + DCRNo + " and U.di_resu = D.di_resu order by D.lvl_qes desc";

                        if (!_objDal.IsRecordExist(_szQuery))
                        {
                            if (_objDal.msgError != "")
                            {
                                msgError = "Error while Last Approvers Approvers List  :" + _objDal.msgError;
                                throw new Exception(msgError);
                            }
                            if (_objDataReader_DCR_Info["sutats_t"].ToString() != "D" && isMigrated == false)
                            {
                                msgError = "No Last Approvers Found";
                                //throw new Exception(msgError);
                            }
                        }
                        _objDataReader = _objDal.DecideDatabaseQDR(_szQuery);
                        if (_objDal.msgError != "")
                        {
                            msgError = "Error while Last Approvers Approvers List  :" + _objDal.msgError;
                            throw new Exception(msgError);
                        }
                        if (_objDataReader.Read())
                        {
                            if (!(_objDataReader["mt_sutats"] is DBNull))
                            {
                                szTime = Convert.ToDateTime(_objDataReader["mt_sutats"]).ToString("T");
                            }
                            if (_objDataReader["ot_detageled"] is DBNull)
                            {
                                szApproversName = Get_Full_Name_of_User(Convert.ToString(_objDataReader["di_resu"]));
                                objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Checked By", PropertyValue = szApproversName });
                            }
                            else
                            {
                                szOwnersName = Get_Full_Name_of_User(Convert.ToString(_objDataReader["ot_detageled"]));
                                if (szOwnersName == null)
                                {
                                    msgError = "Error while getting Delegators information from User master for checked-by   :" + msgError;
                                    throw new Exception(msgError);
                                }
                                objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "A" + iReviewer_Approver, PropertyValue = szOwnersName });
                                if (!(_objDataReader["td_sutats"] is DBNull))
                                {

                                    sbTime = new StringBuilder(_objDataReader["td_sutats"] is DBNull ? "" : Convert.ToDateTime(_objDataReader["td_sutats"]).ToString(szDateFormat) + " " + szTime);
                                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "A" + iReviewer_Approver + "_Dt", PropertyValue = sbTime.ToString() });
                                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "A" + iReviewer_Approver + "_Desig", PropertyValue = Get_Code_Description(_objDataReader["noitangised"].ToString(), "DESG") });
                                    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "A" + iReviewer_Approver + "_Dept", PropertyValue = Get_User_Info_Code_Description(szApproversName, "DPT") });
                                }

                            }

                        }

                        if (_objDataReader != null)
                        {
                            _objDataReader.Close();
                            _objDataReader.Dispose();
                            _objDataReader = null;
                        }

                        #endregion

                        #region .... Add Loc ....
                        //...
                        //...AddLoc
                        IDataReader drLocations;
                        iReviewer_Approver = 1;
                        drLocations = GetLocations(Convert.ToString(DCRNo));
                        while (drLocations.Read())
                        {
                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "AddLoc" + iReviewer_Approver, PropertyValue = Convert.ToString(drLocations["noitacol_eussi"]) });
                            iReviewer_Approver = iReviewer_Approver + 1;
                        }
                        if (drLocations != null)
                        {
                            drLocations.Close();
                            drLocations.Dispose();
                            drLocations = null;
                        }
                        iReviewer_Approver = 1;

                        #endregion

                        #region ..... Issue Date ....
                        #endregion

                        #region ...Print Variables ....
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Print_Purpose", PropertyValue = "Print_Purpose" });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Receipent_Name", PropertyValue = "Receipent_Name" });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Terminal_Name", PropertyValue = "Terminal_Name" });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Printed_By", PropertyValue = "Printed_By" });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Print_Copy_As", PropertyValue = "Print_Copy_As" });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Print_Dt", PropertyValue = "Print_Dt" });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Copy_No", PropertyValue = "Copy_No" });


                        #endregion



                        break;
                    case Documents_Process.Document_Recall:

                        _objDataReader_DCR_Info.Read();
                        _objDataReader_Doc_Info.Read();

                        if (Convert.ToString(_objDataReader_Doc_Info["sutats"]) != "18")
                        {
                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Author", PropertyValue = _objDataReader_DCR_Info["rohtua_cod"] is DBNull ? "Author" : Convert.ToString(_objDataReader_DCR_Info["rohtua_cod"]) });
                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Author_Dt", PropertyValue = "Author_Dt" });
                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Author_Desig", PropertyValue = "Author_Desig" });
                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Author_Dept", PropertyValue = "Author_Dept" });
                        }

                        break;
                    case Documents_Process.TR4:

                        _dicUserName = new Dictionary<string, string>();
                        IDataReader objdrUserMaster;
                        string szUserName = string.Empty;
                        iApp = 0; iRev = 0;

                        #region ... Get Date Format ....
                        _objDataReader_Doc_Info.Read();
                        //Get Date Format
                        objDate = new eDocsDN_BaseFunctions.ClsBaseFunctions(_objDal, AppXmlPath);
                        szDateFormat = GetDateFormat(Convert.ToString(_objDataReader_Doc_Info["ynapmoc"]), Convert.ToString(_objDataReader_Doc_Info["noitacol"]), Convert.ToString(_objDataReader_Doc_Info["tnemtraped"]), Convert.ToString(_objDataReader_Doc_Info["epyt_cod"]));
                        if (szDateFormat != "")
                            szDateFormat = objDate.GetCodeDescription(szDateFormat.ToString(), "DT");
                        else
                            szDateFormat = objDate.GetCodeDescription("1", "DT");

                        #endregion

                        //.. Update Author Department and Designation
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Author_Desig", PropertyValue = _objDataReader_Doc_Info["on_frc"] is DBNull ? "Author_Desig" : Get_Code_Description_For_Author(Convert.ToString(_objDataReader_Doc_Info["on_frc"]), "DESG") });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Author_Dept", PropertyValue = _objDataReader_Doc_Info["on_frc"] is DBNull ? "Author_Dept" : Get_Code_Description_For_Author(Convert.ToString(_objDataReader_Doc_Info["on_frc"]), "DPT") });



                        while (_objDataReader_Doc_Dist.Read())
                        {
                            if (Convert.ToBoolean(_objDataReader_Doc_Dist["codno_tnirp"]))
                            {
                                if (_objDataReader_Doc_Dist["ot_detageled"] is DBNull)
                                    _szQuery = "Select eltit, emanf, emanm, emanl, ngise,txe_ngise, tamrof_etad_resu from zespl_tsm_resu where upper(di_resu) = '" + _objDataReader_Doc_Dist["di_resu"].ToString().ToUpper() + "'";
                                else
                                    _szQuery = "select eltit, emanf, emanm, emanl, ngise,txe_ngise,tamrof_etad_resu from zespl_tsm_resu where upper(di_resu) ='" + _objDataReader_Doc_Dist["ot_detageled"].ToString().ToUpper() + "'";
                                objdrUserMaster = _objDal.DecideDatabaseQDR(_szQuery);
                                if (_objDal.msgError != "")
                                    throw new Exception(_objDal.msgError);

                                if (objdrUserMaster.Read())
                                {
                                    szUserName = Get_Full_Name_of_User(_objDataReader_Doc_Dist["di_resu"].ToString());
                                    if (!(_objDataReader_Doc_Dist["mt_sutats"] is DBNull))
                                    {
                                        szRevTime = Convert.ToDateTime(_objDataReader_Doc_Dist["mt_sutats"]).ToString("T");
                                    }
                                    if (!(_objDataReader_Doc_Dist["td_sutats"] is DBNull))
                                    {
                                        szRevDate = Convert.ToDateTime(_objDataReader_Doc_Dist["td_sutats"]).ToString(szDateFormat) + " " + szRevTime;
                                    }

                                    if (_objDataReader_Doc_Dist["epyt_resu"].ToString() == "A")
                                    {
                                        if (szUserName != "" && szRevDate != "")
                                        {
                                            iApp++;
                                            _dicUserName.Add(_objDataReader_Doc_Dist["epyt_resu"].ToString() + iApp.ToString(), szUserName);
                                            _dicUserName.Add(_objDataReader_Doc_Dist["epyt_resu"].ToString() + iApp.ToString() + "_Dt", szRevDate);
                                            _dicUserName.Add(_objDataReader_Doc_Dist["epyt_resu"].ToString() + iApp.ToString() + "_Desig", Get_Code_Description_For_RevApp(_objDataReader_Doc_Dist["di_resu"].ToString(), "DESG"));
                                            _dicUserName.Add(_objDataReader_Doc_Dist["epyt_resu"].ToString() + iApp.ToString() + "_Dept", Get_User_Info_Code_Description(_objDataReader_Doc_Dist["di_resu"].ToString(), "DPT"));
                                        }
                                    }
                                    else
                                    {
                                        if (szUserName != "" && szRevDate != "")
                                        {
                                            iRev++;
                                            _dicUserName.Add(_objDataReader_Doc_Dist["epyt_resu"].ToString() + iRev.ToString(), szUserName);
                                            _dicUserName.Add(_objDataReader_Doc_Dist["epyt_resu"].ToString() + iRev.ToString() + "_Dt", szRevDate);
                                            _dicUserName.Add(_objDataReader_Doc_Dist["epyt_resu"].ToString() + iRev.ToString() + "_Desig", Get_Code_Description_For_RevApp(_objDataReader_Doc_Dist["di_resu"].ToString(), "DESG"));
                                            _dicUserName.Add(_objDataReader_Doc_Dist["epyt_resu"].ToString() + iRev.ToString() + "_Dept", Get_User_Info_Code_Description(_objDataReader_Doc_Dist["di_resu"].ToString(), "DPT"));
                                        }
                                    }
                                }
                                if (objdrUserMaster != null)
                                {
                                    objdrUserMaster.Close();
                                    objdrUserMaster.Dispose();
                                }
                                objdrUserMaster = null;
                            }
                        }
                        break;
                    default:
                        break;
                }

                switch (eProcess)
                {
                    case Documents_Process.Controller_Live:
                    case Documents_Process.Transfer_Document:


                        #region .... Review /Approver

                        for (iCount = 1; iCount <= 10; iCount++)
                        {
                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "R" + iCount, PropertyValue = "R" + iCount });
                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "R" + iCount + "_Dt", PropertyValue = "R" + iCount + "_Dt" });
                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "R" + iCount + "_Desig", PropertyValue = "R" + iCount + "_Desig" });
                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "R" + iCount + "_Dept", PropertyValue = "R" + iCount + "_Dept" });

                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "AddLoc" + iCount, PropertyValue = "AddLoc" + iCount });
                            if (iCount <= 5)
                            {
                                objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "A" + iCount, PropertyValue = "A" + iCount });
                                objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "A" + iCount + "_Dt", PropertyValue = "A" + iCount + "_Dt" });
                                objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "A" + iCount + "_Desig", PropertyValue = "A" + iCount + "_Desig" });
                                objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "A" + iCount + "_Dept", PropertyValue = "A" + iCount + "_Dept" });

                            }
                        }

                        #endregion

                        #region .... Info Variables ....


                        iCount = 0;
                        Info = string.Empty;
                        for (iCount = 1; iCount <= 20; iCount++)
                        {
                            Info = "info" + iCount;
                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = Info, PropertyValue = Info });
                        }


                        #endregion

                        #region .... Info Values ....

                        _szQuery = "select eulav_lebal,elbairav_motsuc from zespl_atad_lebal where on_frc =" + Convert.ToInt32(_objDataReader_DCR_Info["on_frc"]) + " order by di_cer";
                        _objDataReader = _objDal.DecideDatabaseQDR(_szQuery);
                        if (_objDataReader == null)
                        {
                            throw new Exception(_objDal.msgError);
                        }
                        while (_objDataReader.Read())
                        {
                            if (!(_objDataReader["eulav_lebal"] is DBNull))
                            {
                                Info = _objDataReader["elbairav_motsuc"].ToString().ToLower();
                                //Code Commented And Added for DR:920149
                                //objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = Info, PropertyValue = Convert.ToString(_objDataReader["eulav_lebal"]) });
                                objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = Info, PropertyValue = string.IsNullOrEmpty(Convert.ToString(_objDataReader["eulav_lebal"])) ? "_____" : Convert.ToString(_objDataReader["eulav_lebal"]) });
                            }
                            else
                            {
                                Info = _objDataReader["elbairav_motsuc"].ToString().ToLower();
                                objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = Info, PropertyValue = "_____" });
                            }

                        }


                        if (_objDataReader != null)
                        {
                            _objDataReader.Close();
                            _objDataReader.Dispose();
                            _objDataReader = null;
                        }

                        #endregion

                        #region .... Other Custom Variables ....

                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Checked By", PropertyValue = "Checked By" });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Review Date", PropertyValue = "Review Date" });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "AddIn_Changes", PropertyValue = "AddIn_Changes" });

                        if (Document_Status == "")
                            Document_Status = "Draft";

                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Status", PropertyValue = Document_Status });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Ref No", PropertyValue = string.IsNullOrEmpty(Convert.ToString(_objDataReader_DCR_Info["on_fer_ruoy"])) ? "_____" : Convert.ToString(_objDataReader_DCR_Info["on_fer_ruoy"]) });

                        #endregion

                        break;
                    case Documents_Process.Preview:

                        #region .... Info Variables ....


                        iCount = 0;
                        Info = string.Empty;
                        for (iCount = 1; iCount <= 20; iCount++)
                        {
                            Info = "info" + iCount;
                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = Info, PropertyValue = Info });
                        }


                        #endregion
                        #region .... Info Values ....

                        _szQuery = "select eulav_lebal,elbairav_motsuc from zespl_atad_lebal where on_frc =" + Convert.ToInt32(_objDataReader_DCR_Info["on_frc"]) + " order by di_cer";
                        _objDataReader = _objDal.DecideDatabaseQDR(_szQuery);
                        if (_objDataReader == null)
                        {
                            throw new Exception(_objDal.msgError);
                        }
                        while (_objDataReader.Read())
                        {
                            if (!(_objDataReader["eulav_lebal"] is DBNull))
                            {
                                Info = _objDataReader["elbairav_motsuc"].ToString().ToLower();
                                //Code Commented And Added for DR:920149
                                //objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = Info, PropertyValue = Convert.ToString(_objDataReader["eulav_lebal"]) });
                                objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = Info, PropertyValue = string.IsNullOrEmpty(Convert.ToString(_objDataReader["eulav_lebal"])) ? "_____" : Convert.ToString(_objDataReader["eulav_lebal"]) });
                            }
                            else
                            {
                                Info = _objDataReader["elbairav_motsuc"].ToString().ToLower();
                                objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = Info, PropertyValue = "_____" });
                            }

                        }


                        if (_objDataReader != null)
                        {
                            _objDataReader.Close();
                            _objDataReader.Dispose();
                            _objDataReader = null;
                        }

                        #endregion

                        #region .... Other Custom Variables ....

                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Checked By", PropertyValue = "Checked By" });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Review Date", PropertyValue = "Review Date" });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "AddIn_Changes", PropertyValue = "AddIn_Changes" });


                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Ref No", PropertyValue = string.IsNullOrEmpty(Convert.ToString(_objDataReader_DCR_Info["on_fer_ruoy"])) ? "_____" : Convert.ToString(_objDataReader_DCR_Info["on_fer_ruoy"]) });

                        #endregion

                        #region ...Print Variables ....
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Print_Purpose", PropertyValue = "xxxxxxxxxx xxxxxx" });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Receipent_Name", PropertyValue = "xxxxxxxxxx xxxxxxxxxx xxxxxxxx" });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Terminal_Name", PropertyValue = "xxxxxxxxxx" });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Printed_By", PropertyValue = "xxxxxxxxxx xxxxxxxxxx xxxxxxxx" });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Print_Copy_As", PropertyValue = "xxxxxxxxxx xxxxxx" });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Print_Dt", PropertyValue = "99/99/9999  HH:MM:SS" });
                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Copy_No", PropertyValue = "99" });


                        #endregion

                        break;

                    case Documents_Process.Controller_Publish:

                        if (!string.IsNullOrEmpty(Document_Status))
                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Status", PropertyValue = Document_Status });




                        break;
                    case Documents_Process.Document_Recall:

                        //objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Status", PropertyValue = InProcessDocStatus });

                        #region .... Review / Approver ....

                        //for (iCount = 1; iCount <= 10; iCount++)
                        //{
                        //    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "R" + iCount, PropertyValue = "R" + iCount });
                        //    objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "R" + iCount + "_Dt", PropertyValue = "R" + iCount + "_Dt" });
                        //    if (iCount <= 5)
                        //    {
                        //        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "A" + iCount, PropertyValue = "A" + iCount });
                        //        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "A" + iCount + "_Dt", PropertyValue = "A" + iCount + "_Dt" });
                        //    }
                        //}

                        #endregion

                        break;

                    case Documents_Process.TR4:

                        if (_dicUserName != null)
                        {
                            foreach (KeyValuePair<string, string> kv in _dicUserName)
                                objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = kv.Key, PropertyValue = kv.Value });
                        }
                        if (!string.IsNullOrEmpty(Document_Status))
                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Status", PropertyValue = Document_Status });

                        break;
                    case Documents_Process.obsolete_Document:
                        if (!string.IsNullOrEmpty(Document_Status))
                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Status", PropertyValue = Document_Status });

                        break;
                    case Documents_Process.Expired_Document:

                        _objDataReader_Doc_Info.Read();

                        #region ... Get Date Format ....
                        objDate = new eDocsDN_BaseFunctions.ClsBaseFunctions(_objDal, AppXmlPath);
                        //Get Date Format
                        szDateFormat = GetDateFormat(Convert.ToString(_objDataReader_Doc_Info["ynapmoc"]), Convert.ToString(_objDataReader_Doc_Info["noitacol"]), Convert.ToString(_objDataReader_Doc_Info["tnemtraped"]), Convert.ToString(_objDataReader_Doc_Info["epyt_cod"]));
                        if (szDateFormat != "")
                            szDateFormat = objDate.GetCodeDescription(szDateFormat.ToString(), "DT");
                        else
                            szDateFormat = objDate.GetCodeDescription("1", "DT");

                        #endregion


                        objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Doc_exp_Date", PropertyValue = ConvertToDateTime("Document expiry date", Convert.ToString(DateTime.Now)).ToString(szDateFormat) });

                        if (!string.IsNullOrEmpty(Document_Status))
                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Status", PropertyValue = Document_Status });

                        break;
                    case Documents_Process.Attach_Custom_Variables_To_Template:
                        break;
                    default:
                        if (!string.IsNullOrEmpty(Document_Status))
                            objUpdate_Cust.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Status", PropertyValue = Document_Status });
                        break;
                }


                _strmDocument = objUpdate_Cust.Attach_Document_Custom_Property(objUpdate_Cust.lstCustom_Properties);
                if (!string.IsNullOrEmpty(objUpdate_Cust.msgError))
                    throw new Exception(objUpdate_Cust.msgError);



            }
            finally
            {
                if (objUpdate_Cust != null)
                    objUpdate_Cust.Dispose();
                objUpdate_Cust = null;
            }
            return _strmDocument;
        }

        private IDataReader GetLocations(string szDCRNo)
        {

            IDataReader objdr_Locations = null;
            _szQuery = "select noitacol_eussi from zespl_col_lppa where on_frc = " + szDCRNo + " order by on_rs";
            objdr_Locations = _objDal.DecideDatabaseQDR(_szQuery);
            if (_objDal.msgError != "")
            {
                throw new Exception("Error while getting Locations(ADD_LOC)" + _objDal.msgError);
            }
            return objdr_Locations;
        }

        private string Get_Full_Name_of_User(string szUserID)
        {

            string szOwnerName = null;
            IDataReader objDataReader = null;
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
                    szOwnerName = Convert.ToString(objDataReader["eltit"]) + " " +
                        Convert.ToString(objDataReader["emanf"]) + " ";
                    if (objDataReader["emanm"].ToString() != "")
                        szOwnerName = szOwnerName + Convert.ToString(objDataReader["emanm"]) + " ";
                    szOwnerName = szOwnerName + Convert.ToString(objDataReader["emanl"]);

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
            _objDataReader = null;
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

        private string Get_Code_Description_For_Author(string szDCR_Number, string szType)
        {
            _objDataReader = null;
            string szCodeDescription = string.Empty;
            try
            {
                switch (szType)
                {
                    case "DPT":
                        _szQuery = "select csed_dc from zespl_tsm_edoc where edoc=(select dc_tped from zespl_tsm_resu where di_resu = (select di_rohtua From zespl_ofni_cod where on_frc = " + szDCR_Number + ")) and epyt = '" + szType + "'";
                        break;
                    case "DESG":
                        _szQuery = "select csed_dc from zespl_tsm_edoc where edoc=(select noitangised from zespl_tsm_resu where di_resu = (select di_rohtua From zespl_ofni_cod where on_frc = " + szDCR_Number + ")) and epyt = '" + szType + "'";
                        break;
                    default:
                        break;

                }
                if (!string.IsNullOrEmpty(_szQuery))
                {
                    _objReturnVal = _objDal.GetFirstColumnValue(_szQuery);
                    if (_objDal.msgError != "")
                        throw new Exception("Error while Gettting Code Desc : " + _objDal.msgError);
                    if (_objReturnVal != null)
                    {
                        szCodeDescription = Convert.ToString(_objReturnVal);
                    }
                    _objReturnVal = null;
                }

            }
            finally
            {
                _objReturnVal = null;
            }
            return szCodeDescription;
        }

        private string Get_Code_Description_For_RevApp(string szUserID, string szType)
        {
            _objDataReader = null;
            string szCodeDescription = string.Empty;
            try
            {
                switch (szType)
                {
                    case "DPT":
                        _szQuery = "select csed_dc from zespl_tsm_edoc where edoc=(select dc_tped from zespl_tsm_resu where di_resu='" + szUserID + "')";
                        break;
                    case "DESG":
                        _szQuery = "select csed_dc from zespl_tsm_edoc where edoc=(select noitangised from zespl_tsm_resu where di_resu ='" + szUserID + "')";
                        break;
                    default:
                        break;

                }
                if (!string.IsNullOrEmpty(_szQuery))
                {
                    _objReturnVal = _objDal.GetFirstColumnValue(_szQuery);
                    if (_objDal.msgError != "")
                        throw new Exception("Error while Gettting Code Desc : " + _objDal.msgError);
                    if (_objReturnVal != null)
                    {
                        szCodeDescription = Convert.ToString(_objReturnVal);
                    }
                    _objReturnVal = null;
                }

            }
            finally
            {
                _objReturnVal = null;
            }
            return szCodeDescription;
        }


        private string Get_User_Info_Code_Description(string szUserID, string szCodeType)
        {
            _objDataReader = null;
            string szCodeDescription = string.Empty;
            try
            {

                _szQuery = "select C.csed_dc from zespl_tsm_resu U inner join zespl_tsm_edoc C on c.edoc = U.dc_tped and C.epyt = '" + szCodeType + "'  where di_resu = '" + szUserID + "'";
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


        private string Get_HRK_Code_Desc(string szComp, string szLoc, string szDept, string szDocType, string szCode, string szType, string szLabel4, string szLabel5)
        {

            string szDesc = szCode;
            string szHdrSurrKey = "";
            _objReturnVal = null;
            try
            {
                if (Is_Labels_ExistIn_Mapping(szComp, szLoc, szDept, szDocType))
                {
                    switch (szType)
                    {
                        case "LEBAL4":
                            //_szQuery = "Select lebal_csed From zespl_tsm_edoc_lebal_tsuc Where lebal_type in ('Label4','LABEL4','LEBAL4','lebal4') And lebal_edoc='" + szCode.Trim() + "'";
                            //_szQuery = "Select csed From zespl_pam_code_4lebal_tsuc Where lebal='LABEL4' And edoc='" + szCode + "'";

                            _szQuery = " select csed from zespl_pam_code_4lebal_tsuc C inner join zespl_redaeh_lebal_tsuc H on C.yek_rrus_rdh=H.yek_rrus_rdh" +
                                     " where h.ynapmoc='" + szComp + "' AND h.noitacol='" + szLoc + "' AND h.tnemtraped='" + szDept + "' AND h.epyt_cod in ('" + szDocType + "','All')" +
                                     " AND edoc='" + szCode + "'";

                            break;
                        case "LEBAL5":
                            //_szQuery = "Select lebal_csed From zespl_tsm_edoc_lebal_tsuc Where lebal_type in ('Label5','LABEL5','LEBAL5','lebal5') And lebal_edoc='" + szCode.Trim() + "'";
                            //_szQuery = "Select csed From zespl_pam_code_5lebal_tsuc Where lebal='LABEL5' And edoc='" + szCode + "'";
                            _szQuery = " select l5.csed from zespl_pam_code_4lebal_tsuc C inner join zespl_redaeh_lebal_tsuc H on C.yek_rrus_rdh=H.yek_rrus_rdh" +
                                      " inner join zespl_pam_code_5lebal_tsuc l5 on l5.yek_rrus_rdh=c.yek_rrus_ltd " +
                                      " where h.ynapmoc='" + szComp + "' AND h.noitacol='" + szLoc + "' AND h.tnemtraped='" + szDept + "' AND h.epyt_cod in ('" + szDocType + "','All')" +
                                      " AND l5.edoc='" + szCode + "' AND c.edoc='" + szLabel4 + "'";

                            break;
                        case "LEBAL6":
                            //_szQuery = "Select lebal_csed From zespl_tsm_edoc_lebal_tsuc Where lebal_type in ('Label6','LABEL6','LEBAL6','lebal6') And lebal_edoc='" + szCode.Trim() + "'";
                            //_szQuery = "Select csed From zespl_pam_code_6lebal_tsuc Where lebal='LABEL6' And edoc='" + szCode + "'";

                            _szQuery = "   select l6.csed from zespl_pam_code_4lebal_tsuc C inner join zespl_redaeh_lebal_tsuc H on C.yek_rrus_rdh=H.yek_rrus_rdh" +
                                     " inner join zespl_pam_code_5lebal_tsuc l5 on l5.yek_rrus_rdh=c.yek_rrus_ltd " +
                                     " inner join zespl_pam_code_6lebal_tsuc l6 on l6.yek_rrus_rdh=l5.yek_rrus_ltd" +
                                     " where h.ynapmoc='" + szComp + "' AND h.noitacol='" + szLoc + "' AND h.tnemtraped='" + szDept + "' AND h.epyt_cod in ('" + szDocType + "','All')" +
                                     "   AND l6.edoc='" + szCode + "' AND l5.edoc='" + szLabel5 + "' AND c.edoc='" + szLabel4 + "'";

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

        private bool Is_Labels_ExistIn_Mapping(string szCompany, string szLocation, string szDept, string szDocType)
        {
            msgError = "";
            bool bReturn = false;
            object objReturnVal = null;
            try
            {

                //_szQuery = "Select yek_rrus_rdh From zespl_redaeh_lebal_tsuc";
                //bReturn = _objDal.IsRecordExist(_szQuery);
                //if (_objDal.msgError != "")
                //    throw new Exception("Error while selecting record from zespl_redaeh_lebal_tsuc table" + _objDal.msgError);

                _szQuery = "select gnippam_elbane from  zespl_setalpmet_cod_txe " +
                        " where ynapmoc = '" + szCompany + "' and noitacol = '" + szLocation + "' and tnemtraped = '" + szDept + "'" +
                        " and epyt_cod = '" + szDocType + "'";
                objReturnVal = _objDal.GetFirstColumnValue(_szQuery);
                if (_objDal.msgError != "")
                    throw new Exception(_objDal.msgError);
                if (objReturnVal != null)
                    bReturn = Convert.ToBoolean(objReturnVal);
                objReturnVal = null;


            }
            finally
            {

            }
            return bReturn;
        }

        private string GetDateFormat(string szCompany, string szLocation, string szDepartment, string szDocType)
        {
            string szConfigDate = string.Empty;
            try
            {
                _szQuery = "SELECT dt_ngis_ele FROM zespl_setalpmet_cod WHERE UPPER(ynapmoc) = '" + szCompany.ToUpper() +
                    "' AND UPPER(noitacol) = '" + szLocation.ToUpper() + "' AND UPPER(tnemtraped) = '" + szDepartment.ToUpper() +
                    "' AND UPPER(epyt_cod) = '" + szDocType.ToUpper() + "'";
                _objReturnVal = _objDal.GetFirstColumnValue(_szQuery);
                if (_objDal.msgError != "")
                    throw new Exception(_objDal.msgError);

                if (_objReturnVal != null)
                {
                    szConfigDate = Convert.ToString(_objReturnVal);
                }
                _objReturnVal = null;
            }
            finally
            {
                _objReturnVal = null;
            }
            return szConfigDate;
        }

        private DateTime ConvertToDateTime(string szDateName, string szDatetimeToConvert)
        {
            if (objComm == null)
                objComm = new ClsCommonFunction();

            DateTime sdtConvertedDateTime = objComm.ConvertToDateTime(szDatetimeToConvert);
            if (objComm.msgError != "")
            {
                msgError = "Error While Converting " + szDateName + " (" + Convert.ToString(szDatetimeToConvert) + ") -> " + objComm.msgError;
                throw new Exception(msgError);
            }
            return sdtConvertedDateTime;
        }

        internal static Stream Convert_Document_To_Stream(byte[] arrDocument)
        {
            MemoryStream strmDocument = new MemoryStream();
            strmDocument.Write(arrDocument, 0, (int)arrDocument.Length);
            return strmDocument;
        }

        #endregion

        #region .... IDISPOSABLE ....

        public new void Dispose()
        {
            Dispose(true);
        }
        protected new virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_objDataReader_DCR_Info != null)
                {
                    _objDataReader_DCR_Info.Close();
                    _objDataReader_DCR_Info.Dispose();
                    _objDataReader_DCR_Info = null;
                }
                if (_objDataReader_Doc_Info != null)
                {
                    _objDataReader_Doc_Info.Close();
                    _objDataReader_Doc_Info.Dispose();
                    _objDataReader_Doc_Info = null;
                }
                if (_objDataReader_Issueance_Details != null)
                {
                    _objDataReader_Issueance_Details.Close();
                    _objDataReader_Issueance_Details.Dispose();
                    _objDataReader_Issueance_Details = null;
                }
                if (_objCustomVariables != null)
                {
                    _objCustomVariables.Close();
                    _objCustomVariables.Dispose();
                }
                if (_objDataReader_Doc_Dist != null)
                {
                    _objDataReader_Doc_Dist.Close();
                    _objDataReader_Doc_Dist.Dispose();
                }
                _objDataReader_Doc_Dist = null;

                _objCustomVariables = null;

                if (_objDataReader != null)
                {
                    _objDataReader.Close();
                    _objDataReader.Dispose();
                    _objDataReader = null;
                }

                if (_objDal != null)
                {
                    _objDal.CloseConnection();
                    _objDal.Dispose();
                }
                _objDal = null;

                if (_objINI != null)
                    _objINI.Dispose();
                _objINI = null;


                _dicUserName = null;
                szCreatedAT = string.Empty;
                szChangedAT = string.Empty;
                szAuthDate = string.Empty;
                szRevDate = string.Empty;
                szRevTime = string.Empty;
                szDateFormat = string.Empty;
                Info = string.Empty;

            }
            else
            {

            }
        }

        ~ClsGet_Custom_Properties()
        {
            Dispose(false);
        }


        #endregion


    }
}