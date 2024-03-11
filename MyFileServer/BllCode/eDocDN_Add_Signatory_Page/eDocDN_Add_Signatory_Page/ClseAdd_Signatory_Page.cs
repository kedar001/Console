
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DDLLCS;
using System.Data;
using eDocsDN_Get_Directory_Info;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;



namespace eDocDN_Add_Signatory_Page
{
    public class ClseAdd_Signatory_Page : IDisposable
    {
        #region .... Variable Declaration .....
        string _szAppXmlPath = string.Empty;
        string _szDBName = string.Empty;
        string _szFilePath = string.Empty;
        string _szSqlQuery = string.Empty;
        string _szDateFormat = string.Empty;
        ClsBuildQuery _objDal = null;
        public Stream _strmDocument;
        clsPossitionOfSignature _oConfig = null;
        bool _bGenerateTable = true;
        string _szLogFileName = string.Empty;
        bool _bGenerateLog = false;

        #endregion
        public string msgError { get; set; }
        public Documents_Process eDocumentProcess { get; set; }

        public ClseAdd_Signatory_Page(string szAppXmlPath, string szFileName)
        {
            msgError = string.Empty;
            _szAppXmlPath = szAppXmlPath;
            _szFilePath = szFileName;
            _szLogFileName = AppDomain.CurrentDomain.BaseDirectory + @"\Log\eDetailLog.txt";
            if (!Directory.Exists(AppDomain.CurrentDomain.BaseDirectory + @"\Log\"))
                Directory.CreateDirectory(AppDomain.CurrentDomain.BaseDirectory + @"\Log\");
        }
        public ClseAdd_Signatory_Page(ClsBuildQuery objDal, string szFileName)
        {
            msgError = string.Empty;
            _objDal = objDal;
            _szFilePath = szFileName;
            _szLogFileName = AppDomain.CurrentDomain.BaseDirectory + @"\Log\eDetailLog.txt";
            if (!Directory.Exists(AppDomain.CurrentDomain.BaseDirectory + @"\Log\"))
                Directory.CreateDirectory(AppDomain.CurrentDomain.BaseDirectory + @"\Log\");
        }
        public ClseAdd_Signatory_Page(ClsBuildQuery objDal, Stream strmDocument)
        {
            msgError = string.Empty;
            _objDal = objDal;
            _strmDocument = strmDocument;
            _szLogFileName = AppDomain.CurrentDomain.BaseDirectory + @"\Log\eDetailLog.txt";
            if (!Directory.Exists(AppDomain.CurrentDomain.BaseDirectory + @"\Log\"))
                Directory.CreateDirectory(AppDomain.CurrentDomain.BaseDirectory + @"\Log\");
        }

        public ClseAdd_Signatory_Page(string szDBName, string szAppXmlPath, Stream strmDocument)
        {
            msgError = "";
            _szAppXmlPath = szAppXmlPath;
            _szDBName = szDBName;
            _objDal = new ClsBuildQuery(_szDBName, _szAppXmlPath);
            _objDal.OpenConnection(ConnectionFor.DocsExecutive, ConnectionType.NewConnection);
            _szFilePath = null;
            _strmDocument = strmDocument;
            _szLogFileName = AppDomain.CurrentDomain.BaseDirectory + @"\Log\eDetailLog.txt";
            if (!Directory.Exists(AppDomain.CurrentDomain.BaseDirectory + @"\Log\"))
                Directory.CreateDirectory(AppDomain.CurrentDomain.BaseDirectory + @"\Log\");
        }
        public ClseAdd_Signatory_Page(string szDBName, string szAppXmlPath, string szFilePath)
        {
            msgError = "";
            _szAppXmlPath = szAppXmlPath;
            _szDBName = szDBName;
            _objDal = new ClsBuildQuery(_szDBName, _szAppXmlPath);
            _objDal.OpenConnection(ConnectionFor.DocsExecutive, ConnectionType.NewConnection);
            _strmDocument = null;
            _szFilePath = szFilePath;
            _szLogFileName = AppDomain.CurrentDomain.BaseDirectory + @"\Log\eDetailLog.txt";
            if (!Directory.Exists(AppDomain.CurrentDomain.BaseDirectory + @"\Log\"))
                Directory.CreateDirectory(AppDomain.CurrentDomain.BaseDirectory + @"\Log\");
        }


        #region ... Public Functions ....
        public bool Add_Signatory_Page(int iDCR_Number)
        {
            bool bResult = true;
            IDataReader objDataReader = null;
            //bool bAutoGenerateSignatoryPage = false;

            List<clsSignatore> lstSignatures = new List<clsSignatore>();



            try
            {
                switch (eDocumentProcess)
                {

                    case Documents_Process.Preview:
                        if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " =================Preview===============" + Environment.NewLine);
                        _szSqlQuery = " select yek_cod from zespl_tsid_cod where yek_cod=" + iDCR_Number;
                        _bGenerateTable = _objDal.IsRecordExist(_szSqlQuery);
                        if (_objDal.msgError != "")
                            throw new Exception(_objDal.msgError);
                        if (_bGenerateTable)
                        {
                            _szSqlQuery = " select yek_cod from zespl_tsid_cod where yek_cod=" + iDCR_Number + " and (sutats != 8 OR sutats IS NULL)";
                            _bGenerateTable = !(_objDal.IsRecordExist(_szSqlQuery));
                            if (_objDal.msgError != "")
                                throw new Exception(_objDal.msgError);
                        }

                        _szSqlQuery = " select otua_erutangis_etareneg,btratstadneppA from zespl_setalpmet_cod Doc inner join  zespl_ofni_cod D " +
                            " on d.ynapmoc = doc.ynapmoc and d.noitacol = doc.noitacol and D.tnemtraped = Doc.tnemtraped and D.epyt_cod = doc.epyt_cod" +
                            " where on_frc =" + iDCR_Number + " and  otua_erutangis_etareneg=1";
                        objDataReader = _objDal.DecideDatabaseQDR(_szSqlQuery);
                        if (_objDal.msgError != "")
                            throw new Exception(_objDal.msgError);
                        if (objDataReader.Read())
                        {
                            _oConfig = new clsPossitionOfSignature();
                            _oConfig.AutoGenerateSignature = Convert.ToBoolean(objDataReader["otua_erutangis_etareneg"]);
                            _oConfig.isFirstPageSignature = Convert.ToBoolean(objDataReader["btratstadneppA"]);
                        }
                        if (objDataReader != null)
                        {
                            objDataReader.Close();
                            objDataReader.Dispose();
                        }
                        objDataReader = null;

                        break;

                    default:
                        _szSqlQuery = " select otua_erutangis_etareneg,btratstadneppA from zespl_setalpmet_cod Doc inner join  zespl_ofni_cod D " +
                            " on d.ynapmoc = doc.ynapmoc and d.noitacol = doc.noitacol and D.tnemtraped = Doc.tnemtraped and D.epyt_cod = doc.epyt_cod" +
                            " where on_frc =" + iDCR_Number + " and  otua_erutangis_etareneg=1";
                        objDataReader = _objDal.DecideDatabaseQDR(_szSqlQuery);
                        if (_objDal.msgError != "")
                            throw new Exception(_objDal.msgError);

                        if (objDataReader.Read())
                        {
                            _oConfig = new clsPossitionOfSignature();
                            _oConfig.AutoGenerateSignature = Convert.ToBoolean(objDataReader["otua_erutangis_etareneg"]);
                            _oConfig.isFirstPageSignature = Convert.ToBoolean(objDataReader["btratstadneppA"]);
                        }
                        if (objDataReader != null)
                        {
                            objDataReader.Close();
                            objDataReader.Dispose();
                        }
                        objDataReader = null;

                        break;
                }
                if (_oConfig != null)
                    if (_oConfig.AutoGenerateSignature)
                    {

                        eDocsDN_BaseFunctions.ClsBaseFunctions objDate = new eDocsDN_BaseFunctions.ClsBaseFunctions(_objDal, _szAppXmlPath);
                        _szDateFormat = Get_Date_Format(iDCR_Number);
                        if (_szDateFormat != "")
                            _szDateFormat = objDate.GetCodeDescription(_szDateFormat.ToString(), "DT");
                        else
                            _szDateFormat = objDate.GetCodeDescription("1", "DT");

                        //_szSqlQuery = "select epyt_resu, lvl_qes, d.di_resu, td_sutats, mt_sutats, ot_detageled, codno_tnirp,U.noitangised,U.dc_tped,u.eltit+' '+  u.emanf+' '+u.emanm+' '+u.emanl as FullName " +
                        //            " from zespl_tsid_cod d inner join zespl_tsm_resu U on d.di_resu = U.di_resu where epyt_resu in('R', 'A') And yek_cod =" + iDCR_Number + " and d.codno_tnirp=1  Union " +
                        //            "  select  'AU' as epyt_resu, 0 as lvl_qes,di_rohtua as di_resu ,ISNULL(no_degnahc ,d.no_detaerc) as td_sutats," +
                        //            " ISNULL(ta_degnahc ,d.ta_detaerc) as mt_sutats, null as ot_detageled, 1 as codno_tnirp,U.noitangised,U.dc_tped,u.eltit+' '+  u.emanf+' '+u.emanm+' '+u.emanl as FullName " +
                        //            " from zespl_ofni_cod d inner join zespl_tsm_resu U on d.di_rohtua = U.di_resu where on_frc = " + iDCR_Number + "   order by td_sutats ,mt_sutats";

                        _szSqlQuery = "select epyt_resu, lvl_qes, d.di_resu, td_sutats, mt_sutats, ot_detageled, codno_tnirp,C.csed_dc as noitangised ,Dept.csed_dc as dc_tped ,u.eltit + ' ' + u.emanf + ' ' + u.emanm + ' ' + u.emanl as FullName" +
                                      " from zespl_tsid_cod d inner join zespl_tsm_resu U on d.di_resu = U.di_resu left join zespl_tsm_edoc C on C.edoc = U.noitangised and c.epyt='DESG'" +
                                      " left join zespl_tsm_edoc Dept on Dept.edoc = U.dc_tped and Dept.epyt='DPT' where epyt_resu in('R', 'A') And yek_cod = " + iDCR_Number + " and d.codno_tnirp = 1  Union " +
                                      " select  'AU' as epyt_resu, 0 as lvl_qes,di_rohtua as di_resu ,ISNULL(no_degnahc, d.no_detaerc) as td_sutats, " +
                                     " ISNULL(ta_degnahc, d.ta_detaerc) as mt_sutats, null as ot_detageled, 1 as codno_tnirp,C.csed_dc as noitangised ,Dept.csed_dc as dc_tped,u.eltit + ' ' + u.emanf + ' ' + u.emanm + ' ' + u.emanl as FullName" +
                                     " from zespl_ofni_cod d inner  join zespl_tsm_resu U on d.di_rohtua = U.di_resu left join zespl_tsm_edoc C on C.edoc = U.noitangised and c.epyt='DESG' left join zespl_tsm_edoc Dept on Dept.edoc = U.dc_tped and Dept.epyt='DPT'" +
                                     " where on_frc = " + iDCR_Number + "   order by td_sutats ,mt_sutats ";



                        if (!_objDal.IsRecordExist(_szSqlQuery))
                            throw new Exception("Record Not Exists");

                        objDataReader = _objDal.DecideDatabaseQDR(_szSqlQuery);
                        if (_objDal.msgError != "")
                            throw new Exception(_objDal.msgError);

                        while (objDataReader.Read())
                        {
                            clsSignatore oSignature = new clsSignatore();
                            oSignature.UserID = objDataReader["di_resu"].ToString();
                            if (DBNull.Value != objDataReader["td_sutats"])
                                oSignature.UserDate = Convert.ToDateTime(objDataReader["td_sutats"]);
                            if (DBNull.Value != objDataReader["mt_sutats"])
                                oSignature.UserTime = Convert.ToDateTime(objDataReader["mt_sutats"]);
                            oSignature.UserDesignation = DBNull.Value.Equals(objDataReader["noitangised"]) ? "" : objDataReader["noitangised"].ToString();
                            oSignature.UserType = objDataReader["epyt_resu"].ToString();
                            oSignature.Sequence = objDataReader["lvl_qes"].ToString();
                            oSignature.UserDepartment = DBNull.Value.Equals(objDataReader["dc_tped"]) ? "" : objDataReader["dc_tped"].ToString();
                            oSignature.UserFullName = objDataReader["FullName"].ToString();
                            lstSignatures.Add(oSignature);
                            oSignature = null;
                        }
                        if (objDataReader.Read())
                        {
                            objDataReader.Close();
                            objDataReader.Dispose();
                        }
                        objDataReader = null;

                        if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " END OF DB Activity " + Environment.NewLine);

                        if (_oConfig.isFirstPageSignature)
                        {
                            if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " isFirstPageSignature " + Environment.NewLine);
                            if (_strmDocument != null)
                                Add_Signatory_Page_Stream(lstSignatures);
                            else
                                Add_Signatory_Page(lstSignatures);
                        }
                        else
                        {
                            if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " LastPageSignature " + Environment.NewLine);
                            if (_strmDocument != null)
                                AppendToLastPage_Stream(lstSignatures);
                            else
                                AppendToLastPage(lstSignatures);
                        }
                    }
            }
            catch (Exception ex)
            {
                bResult = false;
                msgError = ex.Message;
                if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " Error :Add_Signatory_Page : " + ex.Message + Environment.NewLine);
                if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " Error :Add_Signatory_Page : " + ex.StackTrace + Environment.NewLine);
            }
            finally
            {
                if (objDataReader != null)
                {
                    objDataReader.Close();
                    objDataReader.Dispose();
                }
                objDataReader = null;
                if (lstSignatures != null)
                    lstSignatures.Clear();
                lstSignatures = null;
                _oConfig = null;
            }
            return bResult;
        }

        public Stream Delete_Signatures(int iDCR_Number)
        {
            bool bResult = true;
            IDataReader objDataReader = null;
            clsPossitionOfSignature _oConfig = null;
            try
            {
                if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " Delete_Signatures : " + Environment.NewLine);
                _szSqlQuery = " select otua_erutangis_etareneg,btratstadneppA from zespl_setalpmet_cod Doc inner join  zespl_ofni_cod D " +
                           " on d.ynapmoc = doc.ynapmoc and d.noitacol = doc.noitacol and D.tnemtraped = Doc.tnemtraped and D.epyt_cod = doc.epyt_cod" +
                           " where on_frc =" + iDCR_Number + " and  otua_erutangis_etareneg=1";

                if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " _szSqlQuery : " + _szSqlQuery + Environment.NewLine);

                objDataReader = _objDal.DecideDatabaseQDR(_szSqlQuery);
                if (_objDal.msgError != "")
                    throw new Exception(_objDal.msgError);

                if (objDataReader.Read())
                {
                    _oConfig = new clsPossitionOfSignature();
                    _oConfig.AutoGenerateSignature = Convert.ToBoolean(objDataReader["otua_erutangis_etareneg"]);
                    _oConfig.isFirstPageSignature = Convert.ToBoolean(objDataReader["btratstadneppA"]);
                }
                if (objDataReader != null)
                {
                    objDataReader.Close();
                    objDataReader.Dispose();
                }
                objDataReader = null;
                if (_oConfig != null)
                {
                    if (_oConfig.isFirstPageSignature)
                    {
                        if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " _oConfig.isFirstPageSignature : " + _oConfig.isFirstPageSignature + Environment.NewLine);
                        if (_strmDocument != null)
                            _strmDocument = Delete_Signature_Stream(0);
                        else
                            Delete_Signature(0);
                    }
                    else
                    {
                        if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " _oConfig.isFirstPageSignature : " + _oConfig.isFirstPageSignature + Environment.NewLine);
                        if (_strmDocument != null)
                            _strmDocument = Delete_Signature_from_LastPage_Stream();
                        else
                            Delete_Signature_from_LastPage();
                    }
                }
                else
                {
                    if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " _oConfig is NULL :" + Environment.NewLine);
                    if (_strmDocument != null)
                        _strmDocument = Delete_Signature_Stream(0);
                    else
                        Delete_Signature(0);

                    if (_strmDocument != null)
                        _strmDocument = Delete_Signature_from_LastPage_Stream();
                    else
                        Delete_Signature_from_LastPage();

                }
            }
            catch (Exception ex)
            {
                if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " Error :Delete_Signatures : " + ex.Message + Environment.NewLine);
                if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " Error :Delete_Signatures : " + ex.StackTrace + Environment.NewLine);
            }

            finally
            {
                if (objDataReader != null)
                {
                    objDataReader.Close();
                    objDataReader.Dispose();
                }
                objDataReader = null;
                _oConfig = null;
            }
            return _strmDocument;
        }
        #endregion

        #region .... Private Fuctions ...

        private void Add_Signatory_Page(List<clsSignatore> lstSignatures)
        {

            try
            {
                Delete_Signature(0);
                using (WordprocessingDocument doc =
               WordprocessingDocument.Open(_szFilePath, true))
                {
                    var body = doc.MainDocumentPart.Document.Body;

                    //..Check BookMark Already Exists
                    List<BookmarkStart> obreak = doc.MainDocumentPart.Document.Body.
                         Descendants<BookmarkStart>().
                         ToList();

                    if (obreak.Exists(c => c.Name == "Sign_F_Table"))
                    {
                        BookmarkStart oBook = obreak.FindAll(c => c.Name == "Sign_F_Table").FirstOrDefault();
                        if (oBook != null)
                        {
                            OpenXmlElement oPara = oBook.Parent;
                            oPara.InsertAfterSelf(Generate_Signatory_Table(lstSignatures));

                        }
                    }
                    else
                    {
                        body.PrependChild<Paragraph>(
                            new Paragraph(new Run(new Break() { Type = BreakValues.Page }))
                            );
                        body.PrependChild(Generate_Signatory_Table(lstSignatures));
                        body.PrependChild<Paragraph>(GenerateTableHeaderParagraph());
                    }
                }
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
        }
        private Stream Add_Signatory_Page_Stream(List<clsSignatore> lstSignatures)
        {

            try
            {
                Delete_Signature_Stream(0);
                using (WordprocessingDocument doc =
               WordprocessingDocument.Open(_strmDocument, true))
                {
                    var body = doc.MainDocumentPart.Document.Body;

                    //..Check BookMark Already Exists
                    List<BookmarkStart> obreak = doc.MainDocumentPart.Document.Body.
                         Descendants<BookmarkStart>().
                         ToList();

                    if (obreak.Exists(c => c.Name == "Sign_F_Table"))
                    {
                        BookmarkStart oBook = obreak.FindAll(c => c.Name == "Sign_F_Table").FirstOrDefault();
                        if (oBook != null)
                        {
                            OpenXmlElement oPara = oBook.Parent;
                            oPara.InsertAfterSelf(Generate_Signatory_Table(lstSignatures));

                        }
                    }
                    else
                    {
                        body.PrependChild<Paragraph>(
                            new Paragraph(new Run(new Break() { Type = BreakValues.Page }))
                            );
                        body.PrependChild(Generate_Signatory_Table(lstSignatures));
                        body.PrependChild<Paragraph>(GenerateTableHeaderParagraph());
                    }
                }
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            return _strmDocument;
        }

        public void AppendToLastPage(List<clsSignatore> lstSignatures)
        {

            try
            {
                if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " AppendToLastPage:" + Environment.NewLine);
                string bookmarkName = "Sign_L_Table";
                Delete_Signature_from_LastPage();
                using (WordprocessingDocument doc =
                    WordprocessingDocument.Open(_szFilePath, true))
                {
                    var body = doc.MainDocumentPart.Document.Body;
                    var paras = body.Elements<Paragraph>().LastOrDefault();
                    //.. body.AppendChild(new Paragraph(new Run(new RunProperties(new Text()))));


                    List<BookmarkStart> obreak = doc.MainDocumentPart.Document.Body.
                            Descendants<BookmarkStart>().
                            ToList();

                    if (obreak.Exists(c => c.Name == "Sign_L_Table"))
                    {
                        BookmarkStart oBook = obreak.FindAll(c => c.Name == "Sign_L_Table").FirstOrDefault();
                        if (oBook != null)
                        {
                            OpenXmlElement oPara = oBook.Parent;
                            oPara.InsertAfterSelf(Generate_Signatory_Table(lstSignatures));

                        }
                    }
                    else
                    {
                        body.AppendChild<Paragraph>(
                      new Paragraph(new Run(new Break() { Type = BreakValues.Page }))
                      );
                        body.AppendChild<Paragraph>(GenerateTableHeaderParagraph());

                        paras = body.Elements<Paragraph>().LastOrDefault();
                        Add_BookMark_For_LastPage(paras, bookmarkName);
                        body.AppendChild(Generate_Signatory_Table(lstSignatures));
                    }


                }
            }
            catch (Exception ex)
            {
                if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " Error :AppendToLastPage : " + ex.Message + Environment.NewLine);
                if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " Error :AppendToLastPage : " + ex.StackTrace + Environment.NewLine);

            }


        }

        public void AppendToLastPage_Stream(List<clsSignatore> lstSignatures)
        {

            try
            {
                if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " =================AppendToLastPage_Stream===============" + Environment.NewLine);
                string bookmarkName = "Sign_L_Table";
                Delete_Signature_from_LastPage_Stream();
                using (WordprocessingDocument doc =
                    WordprocessingDocument.Open(_strmDocument, true))
                {
                    var body = doc.MainDocumentPart.Document.Body;
                    var paras = body.Elements<Paragraph>().LastOrDefault();
                    //.. body.AppendChild(new Paragraph(new Run(new RunProperties(new Text()))));


                    List<BookmarkStart> obreak = doc.MainDocumentPart.Document.Body.
                            Descendants<BookmarkStart>().
                            ToList();

                    if (obreak.Exists(c => c.Name == "Sign_L_Table"))
                    {
                        if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " =================BookMark Found===============" + Environment.NewLine);
                        BookmarkStart oBook = obreak.FindAll(c => c.Name == "Sign_L_Table").FirstOrDefault();
                        if (oBook != null)
                        {
                            OpenXmlElement oPara = oBook.Parent;
                            if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " =================oPara Found===============" + Environment.NewLine);
                            oPara.InsertAfterSelf(Generate_Signatory_Table(lstSignatures));

                        }
                    }
                    else
                    {
                        body.AppendChild<Paragraph>(
                      new Paragraph(new Run(new Break() { Type = BreakValues.Page }))
                      );
                        body.AppendChild<Paragraph>(GenerateTableHeaderParagraph());

                        paras = body.Elements<Paragraph>().LastOrDefault();
                        Add_BookMark_For_LastPage(paras, bookmarkName);
                        body.AppendChild(Generate_Signatory_Table(lstSignatures));
                    }


                }
            }
            catch (Exception ex)
            {
                if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " Error :AppendToLastPage_Stream : " + ex.Message + Environment.NewLine);
                if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " Error :AppendToLastPage_Stream : " + ex.StackTrace + Environment.NewLine);

            }
        }


        private string Get_Date_Format(int iDCR_Number)
        {
            string szConfigDate = string.Empty;
            object _objReturnVal = null;
            try
            {
                _szSqlQuery = " SELECT dt_ngis_ele FROM zespl_setalpmet_cod doc inner join zespl_ofni_cod d on doc.ynapmoc=d.ynapmoc and " +
                              " doc.noitacol = d.noitacol and doc.tnemtraped = d.tnemtraped and doc.epyt_cod = d.epyt_cod" +
                              "  WHERE d.on_frc =" + iDCR_Number;

                _objReturnVal = _objDal.GetFirstColumnValue(_szSqlQuery);
                if (_objDal.msgError != "")
                    throw new Exception(_objDal.msgError);

                if (_objReturnVal != null)
                {
                    szConfigDate = Convert.ToString(_objReturnVal);
                }
                _objReturnVal = null;
            }

            finally { }
            return szConfigDate;
        }

        #endregion


        #region ... Delete  ... 
        public void Delete_Signature(int iDCRNumber)
        {
            if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " Delete_Signature : " + Environment.NewLine);
            try
            {
                Break table = null;
                BookmarkStart bookmark = null;
                bool bPageBreakFound = false;
                bool bPageFound = false;
                bool bConsiderForDeletion = false;
                List<OpenXmlElement> lstParatoRemove = new List<OpenXmlElement>();
                using (WordprocessingDocument myDoc = WordprocessingDocument.Open(_szFilePath, true))
                {
                    MainDocumentPart mainPart = myDoc.MainDocumentPart;

                    List<OpenXmlElement> obreak = myDoc.MainDocumentPart.Document.Body.
                        Elements<OpenXmlElement>().
                        ToList();

                    foreach (var item in obreak)
                    {
                        if (item.Descendants<BookmarkStart>().Count() > 0)
                        {
                            bookmark = item.Descendants<BookmarkStart>().First();
                            if (bookmark.Name == "Sign_F_Table")
                            {
                                bPageFound = true;
                                if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " bPageFound :" + bPageFound + Environment.NewLine);

                            }
                        }
                        if (item.Descendants<Break>().Count() > 0)
                        {
                            table = item.Descendants<Break>().First();
                            if (table.Type.InnerText == "page")
                            {
                                bPageBreakFound = true;
                                if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " bPageBreakFound :" + bPageBreakFound + Environment.NewLine);
                                break;
                            }
                        }
                        if (bConsiderForDeletion && !bPageBreakFound)
                            lstParatoRemove.Add(item);

                        if (bPageFound)
                            bConsiderForDeletion = true;
                    }
                    if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " bPageFound :" + bPageFound + "  bPageBreakFound" + bPageBreakFound + Environment.NewLine);

                    if (bPageFound && bPageBreakFound)
                        foreach (var item in lstParatoRemove)
                        {
                            if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " deleting :" + item.ToString() + Environment.NewLine);
                            item.RemoveAllChildren();
                            item.Remove();
                        }
                }
            }
            catch (Exception ex)
            {
                if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " Error :Delete_Signature : " + ex.Message + Environment.NewLine);
                if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " Error :Delete_Signature : " + ex.StackTrace + Environment.NewLine);
            }
        }
        public Stream Delete_Signature_Stream(int iDCR_Number)
        {
            if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " Delete_Signature_Stream : " + Environment.NewLine);
            try
            {
                Break table = null;
                BookmarkStart bookmark = null;
                bool bPageBreakFound = false;
                bool bPageFound = false;
                bool bConsiderForDeletion = false;
                List<OpenXmlElement> lstParatoRemove = new List<OpenXmlElement>();
                using (WordprocessingDocument myDoc = WordprocessingDocument.Open(_strmDocument, true))
                {
                    MainDocumentPart mainPart = myDoc.MainDocumentPart;

                    List<OpenXmlElement> obreak = myDoc.MainDocumentPart.Document.Body.
                        Elements<OpenXmlElement>().
                        ToList();

                    foreach (var item in obreak)
                    {
                        if (item.Descendants<BookmarkStart>().Count() > 0)
                        {
                            bookmark = item.Descendants<BookmarkStart>().First();
                            if (bookmark.Name == "Sign_F_Table")
                            {
                                bPageFound = true;
                                if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + "  bPageBreakFound" + bPageBreakFound + Environment.NewLine);

                            }
                        }
                        if (item.Descendants<Break>().Count() > 0)
                        {
                            table = item.Descendants<Break>().First();
                            if (table.Type.InnerText == "page")
                            {
                                if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + "   bPageBreakFound" + bPageBreakFound + Environment.NewLine);

                                bPageBreakFound = true;
                                break;
                            }
                        }
                        if (bConsiderForDeletion && !bPageBreakFound)
                            lstParatoRemove.Add(item);

                        if (bPageFound)
                            bConsiderForDeletion = true;
                    }
                    if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " bPageFound :" + bPageFound + "  bPageBreakFound" + bPageBreakFound + Environment.NewLine);
                    if (bPageFound && bPageBreakFound)
                        foreach (var item in lstParatoRemove)
                        {
                            if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " deleting :" + item.ToString() + Environment.NewLine);
                            item.RemoveAllChildren();
                            item.Remove();
                        }
                }
            }
            catch (Exception ex)
            {
                if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " Error :Delete_Signature_Stream : " + ex.Message + Environment.NewLine);
                if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " Error :Delete_Signature_Stream : " + ex.StackTrace + Environment.NewLine);
            }
            return _strmDocument;
        }


        public void Delete_Signature_from_LastPage()
        {
            try
            {
                if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " =================Delete_Signature_from_LastPage===============" + Environment.NewLine);
                Break table = null;
                BookmarkStart bookmark = null;
                bool bPageBreakFound = false;
                bool bConsiderForDeletion = false;
                bool bPageFound = false;
                List<OpenXmlElement> lstParatoRemove = new List<OpenXmlElement>();
                using (WordprocessingDocument myDoc = WordprocessingDocument.Open(_szFilePath, true))
                {
                    MainDocumentPart mainPart = myDoc.MainDocumentPart;

                    List<OpenXmlElement> obreak = myDoc.MainDocumentPart.Document.Body.
                        Elements<OpenXmlElement>().
                        ToList();

                    foreach (var item in obreak)
                    {
                        if (item.Descendants<Break>().Count() > 0)
                        {
                            table = item.Descendants<Break>().First();
                            if (table.Type.InnerText == "page")
                            {
                                bPageBreakFound = true;
                                if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " Page Break :" + item.ToString() + Environment.NewLine);
                            }
                        }

                        if (item.Descendants<BookmarkStart>().Count() > 0)
                        {
                            bookmark = item.Descendants<BookmarkStart>().First();
                            if (bookmark.Name == "Sign_L_Table")
                            {
                                if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " BookMark :" + bookmark.Name + Environment.NewLine);
                                bPageFound = true;
                            }
                        }

                        if (bConsiderForDeletion)
                        {
                            lstParatoRemove.Add(item);
                        }
                        if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " bPageFound :" + bPageFound.ToString() + Environment.NewLine);
                        if (bPageFound)
                            bConsiderForDeletion = true;
                    }
                    if (bPageFound)
                        foreach (var item in lstParatoRemove)
                        {
                            if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " deleting :" + item.ToString() + Environment.NewLine);
                            item.RemoveAllChildren();
                            item.Remove();
                        }
                }
            }
            catch (Exception ex)
            {
                if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " Error :Delete_Signature_from_LastPage : " + ex.Message + Environment.NewLine);
                if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " Error :Delete_Signature_from_LastPage : " + ex.StackTrace + Environment.NewLine);

            }
        }
        public Stream Delete_Signature_from_LastPage_Stream()
        {
            try
            {
                if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " =================Delete_Signature_from_LastPage_Stream===============" + Environment.NewLine);
                Break table = null;
                BookmarkStart bookmark = null;
                bool bPageBreakFound = false;
                bool bConsiderForDeletion = false;
                bool bPageFound = false;
                List<OpenXmlElement> lstParatoRemove = new List<OpenXmlElement>();
                using (WordprocessingDocument myDoc = WordprocessingDocument.Open(_strmDocument, true))
                {
                    MainDocumentPart mainPart = myDoc.MainDocumentPart;

                    List<OpenXmlElement> obreak = myDoc.MainDocumentPart.Document.Body.
                        Elements<OpenXmlElement>().
                        ToList();

                    foreach (var item in obreak)
                    {
                        if (item.Descendants<Break>().Count() > 0)
                        {
                            table = item.Descendants<Break>().First();
                            if (table.Type.InnerText == "page")
                            {
                                if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " table.Type.InnerText :" + table.Type.InnerText + Environment.NewLine);
                                bPageBreakFound = true;
                            }
                        }

                        if (item.Descendants<BookmarkStart>().Count() > 0)
                        {
                            bookmark = item.Descendants<BookmarkStart>().First();
                            if (bookmark.Name == "Sign_L_Table")
                            {
                                if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " BookMark :" + bookmark.Name + Environment.NewLine);
                                bPageFound = true;
                            }
                        }

                        if (bConsiderForDeletion)
                        {
                            lstParatoRemove.Add(item);
                        }
                        if (bPageFound)
                            bConsiderForDeletion = true;
                    }

                    if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " bPageFound :" + bPageFound.ToString() + Environment.NewLine);
                    if (bPageFound)
                        foreach (var item in lstParatoRemove)
                        {
                            if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " deleting :" + item.ToString() + Environment.NewLine);
                            item.RemoveAllChildren();
                            item.Remove();
                        }
                }
            }
            catch (Exception ex)
            {
                if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " Error :Delete_Signature_from_LastPage_Stream : " + ex.Message + Environment.NewLine);
                if (_bGenerateLog) File.AppendAllText(_szLogFileName, DateTime.Now + " Error :Delete_Signature_from_LastPage_Stream : " + ex.StackTrace + Environment.NewLine);

            }
            return _strmDocument;
        }


        #endregion 


        #region ..... Table Structure  ...


        public Table Generate_Signatory_Table(List<clsSignatore> lstSignatory)
        {

            Table table1 = new Table();
            if (_bGenerateTable)
            {

                table1.AppendChild(GenerateTableProperties());
                table1.AppendChild(GenerateTableGrid());
                table1.AppendChild(GenerateTableHeaderRow());

                foreach (clsSignatore item in lstSignatory)
                {
                    if (item.UserType == "AU")
                        table1.AppendChild(Generate_Author_TableRow(item));
                    else
                        table1.AppendChild(Generate_Reviewer_Approver_TableRow(item));
                }

            }
            return table1;
        }
        public Paragraph GenerateTableHeaderParagraph()
        {
            Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "008840CE", RsidParagraphAddition = "00600789", RsidParagraphProperties = "00B73025", RsidRunAdditionDefault = "00791546" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Heading1" };
            Justification justification1 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            RunFonts runFonts1 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };

            paragraphMarkRunProperties1.Append(runFonts1);

            paragraphProperties1.Append(paragraphStyleId1);
            paragraphProperties1.Append(justification1);
            paragraphProperties1.Append(paragraphMarkRunProperties1);
            Run run1 = new Run() { RsidRunProperties = "008840CE" };
            RunProperties runProperties1 = new RunProperties();
            RunFonts runFonts2 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            Bold bold2 = new Bold();
            FontSize fontSize2 = new FontSize() { Val = "40" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "40" };


            runProperties1.Append(runFonts2);
            runProperties1.Append(bold2);
            runProperties1.Append(fontSize2);
            runProperties1.Append(fontSizeComplexScript2);

            Text text1 = new Text();
            text1.Text = "Signature Page";

            run1.Append(runProperties1);
            run1.Append(text1);
            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            Add_BookMark(paragraph1, "Sign_F_Table");

            return paragraph1;

        }

        public TableProperties GenerateTableProperties()
        {
            TableProperties tableProperties1 = new TableProperties();
            TableWidth tableWidth1 = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };

            TableCellMarginDefault tableCellMarginDefault1 = new TableCellMarginDefault();
            TopMargin topMargin1 = new TopMargin() { Width = "15", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin1 = new TableCellLeftMargin() { Width = 15, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin1 = new BottomMargin() { Width = "15", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin1 = new TableCellRightMargin() { Width = 15, Type = TableWidthValues.Dxa };

            tableCellMarginDefault1.Append(topMargin1);
            tableCellMarginDefault1.Append(tableCellLeftMargin1);
            tableCellMarginDefault1.Append(bottomMargin1);
            tableCellMarginDefault1.Append(tableCellRightMargin1);
            TableLook tableLook1 = new TableLook() { Val = "04A0", FirstRow = true, LastRow = false, FirstColumn = true, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = true };

            tableProperties1.Append(tableWidth1);
            tableProperties1.Append(tableCellMarginDefault1);
            tableProperties1.Append(tableLook1);
            return tableProperties1;
        }

        public TableGrid GenerateTableGrid()
        {
            TableGrid tableGrid1 = new TableGrid();
            GridColumn gridColumn1 = new GridColumn() { Width = "3379" };
            GridColumn gridColumn2 = new GridColumn() { Width = "3679" };
            GridColumn gridColumn3 = new GridColumn() { Width = "3509" };
            GridColumn gridColumn4 = new GridColumn() { Width = "3381" };

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);
            tableGrid1.Append(gridColumn3);
            tableGrid1.Append(gridColumn4);
            return tableGrid1;
        }

        public TableRow GenerateTableHeaderRow()
        {

            //            TableRowProperties tblHeaderRowProps = new TableRowProperties(
            //    new CantSplit() { Val = OnOffOnlyValues.On },
            //    new TableHeader() { Val = OnOffOnlyValues.On }
            //);

            //            tblHeaderRow.AppendChild<TableRowProperties>(tblHeaderRowProps);

            TableRow tableRow1 = new TableRow() { RsidTableRowMarkRevision = "00600789", RsidTableRowAddition = "00600789", RsidTableRowProperties = "00A6633F" };

            TableRowProperties tableRowProperties1 = new TableRowProperties();
            CantSplit cantSplit1 = new CantSplit();
            TableHeader tableHeader1 = new TableHeader();

            tableRowProperties1.Append(cantSplit1);
            tableRowProperties1.Append(tableHeader1);

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "1211", Type = TableWidthUnitValues.Pct };

            TableCellBorders tableCellBorders1 = new TableCellBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableCellBorders1.Append(topBorder1);
            tableCellBorders1.Append(leftBorder1);
            tableCellBorders1.Append(bottomBorder1);
            tableCellBorders1.Append(rightBorder1);

            TableCellMargin tableCellMargin1 = new TableCellMargin();
            TopMargin topMargin1 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin1 = new LeftMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin1 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin1 = new RightMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };

            tableCellMargin1.Append(topMargin1);
            tableCellMargin1.Append(leftMargin1);
            tableCellMargin1.Append(bottomMargin1);
            tableCellMargin1.Append(rightMargin1);
            TableCellVerticalAlignment tableCellVerticalAlignment1 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };
            HideMark hideMark1 = new HideMark();

            tableCellProperties1.Append(tableCellWidth1);
            tableCellProperties1.Append(tableCellBorders1);
            tableCellProperties1.Append(tableCellMargin1);
            tableCellProperties1.Append(tableCellVerticalAlignment1);
            tableCellProperties1.Append(hideMark1);

            Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "00600789", RsidParagraphAddition = "00600789", RsidParagraphProperties = "00A461D4", RsidRunAdditionDefault = "00600789" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Justification justification1 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            RunFonts runFonts1 = new RunFonts() { Ascii = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            FontSize fontSize1 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "24" };
            Languages languages1 = new Languages() { EastAsia = "en-IN" };

            paragraphMarkRunProperties1.Append(runFonts1);
            paragraphMarkRunProperties1.Append(fontSize1);
            paragraphMarkRunProperties1.Append(fontSizeComplexScript1);
            paragraphMarkRunProperties1.Append(languages1);

            paragraphProperties1.Append(spacingBetweenLines1);
            paragraphProperties1.Append(justification1);
            paragraphProperties1.Append(paragraphMarkRunProperties1);
            BookmarkStart bookmarkStart1 = new BookmarkStart() { Name = "_GoBack", ColumnFirst = 1, ColumnLast = 1, Id = "0" };

            Run run1 = new Run() { RsidRunProperties = "00600789" };

            RunProperties runProperties1 = new RunProperties();
            RunFonts runFonts2 = new RunFonts() { Ascii = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            Color color1 = new Color() { Val = "000000" };
            FontSize fontSize2 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "24" };
            Languages languages2 = new Languages() { EastAsia = "en-IN" };

            runProperties1.Append(runFonts2);
            runProperties1.Append(bold1);
            runProperties1.Append(boldComplexScript1);
            runProperties1.Append(color1);
            runProperties1.Append(fontSize2);
            runProperties1.Append(fontSizeComplexScript2);
            runProperties1.Append(languages2);
            LastRenderedPageBreak lastRenderedPageBreak1 = new LastRenderedPageBreak();
            Text text1 = new Text();
            text1.Text = "Signatories";

            run1.Append(runProperties1);
            run1.Append(lastRenderedPageBreak1);
            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(bookmarkStart1);
            paragraph1.Append(run1);

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph1);

            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "1319", Type = TableWidthUnitValues.Pct };

            TableCellBorders tableCellBorders2 = new TableCellBorders();
            TopBorder topBorder2 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
            LeftBorder leftBorder2 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder2 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
            RightBorder rightBorder2 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };

            tableCellBorders2.Append(topBorder2);
            tableCellBorders2.Append(leftBorder2);
            tableCellBorders2.Append(bottomBorder2);
            tableCellBorders2.Append(rightBorder2);

            TableCellMargin tableCellMargin2 = new TableCellMargin();
            TopMargin topMargin2 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin2 = new LeftMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin2 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin2 = new RightMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };

            tableCellMargin2.Append(topMargin2);
            tableCellMargin2.Append(leftMargin2);
            tableCellMargin2.Append(bottomMargin2);
            tableCellMargin2.Append(rightMargin2);
            HideMark hideMark2 = new HideMark();

            tableCellProperties2.Append(tableCellWidth2);
            tableCellProperties2.Append(tableCellBorders2);
            tableCellProperties2.Append(tableCellMargin2);
            tableCellProperties2.Append(hideMark2);

            Paragraph paragraph2 = new Paragraph() { RsidParagraphMarkRevision = "00600789", RsidParagraphAddition = "00600789", RsidParagraphProperties = "00A461D4", RsidRunAdditionDefault = "00600789" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Justification justification2 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
            RunFonts runFonts3 = new RunFonts() { Ascii = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            FontSize fontSize3 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "24" };
            Languages languages3 = new Languages() { EastAsia = "en-IN" };

            paragraphMarkRunProperties2.Append(runFonts3);
            paragraphMarkRunProperties2.Append(fontSize3);
            paragraphMarkRunProperties2.Append(fontSizeComplexScript3);
            paragraphMarkRunProperties2.Append(languages3);

            paragraphProperties2.Append(spacingBetweenLines2);
            paragraphProperties2.Append(justification2);
            paragraphProperties2.Append(paragraphMarkRunProperties2);

            Run run2 = new Run() { RsidRunProperties = "00600789" };

            RunProperties runProperties2 = new RunProperties();
            RunFonts runFonts4 = new RunFonts() { Ascii = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            Bold bold2 = new Bold();
            BoldComplexScript boldComplexScript2 = new BoldComplexScript();
            Color color2 = new Color() { Val = "000000" };
            FontSize fontSize4 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "24" };
            Languages languages4 = new Languages() { EastAsia = "en-IN" };

            runProperties2.Append(runFonts4);
            runProperties2.Append(bold2);
            runProperties2.Append(boldComplexScript2);
            runProperties2.Append(color2);
            runProperties2.Append(fontSize4);
            runProperties2.Append(fontSizeComplexScript4);
            runProperties2.Append(languages4);
            Text text2 = new Text();
            text2.Text = "Name ";

            run2.Append(runProperties2);
            run2.Append(text2);

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(run2);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphMarkRevision = "00600789", RsidParagraphAddition = "00600789", RsidParagraphProperties = "00A461D4", RsidRunAdditionDefault = "00600789" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Justification justification3 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
            RunFonts runFonts5 = new RunFonts() { Ascii = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            FontSize fontSize5 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "24" };
            Languages languages5 = new Languages() { EastAsia = "en-IN" };

            paragraphMarkRunProperties3.Append(runFonts5);
            paragraphMarkRunProperties3.Append(fontSize5);
            paragraphMarkRunProperties3.Append(fontSizeComplexScript5);
            paragraphMarkRunProperties3.Append(languages5);

            paragraphProperties3.Append(spacingBetweenLines3);
            paragraphProperties3.Append(justification3);
            paragraphProperties3.Append(paragraphMarkRunProperties3);

            Run run3 = new Run() { RsidRunProperties = "00600789" };

            RunProperties runProperties3 = new RunProperties();
            RunFonts runFonts6 = new RunFonts() { Ascii = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            Bold bold3 = new Bold();
            BoldComplexScript boldComplexScript3 = new BoldComplexScript();
            Color color3 = new Color() { Val = "000000" };
            FontSize fontSize6 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "24" };
            Languages languages6 = new Languages() { EastAsia = "en-IN" };

            runProperties3.Append(runFonts6);
            runProperties3.Append(bold3);
            runProperties3.Append(boldComplexScript3);
            runProperties3.Append(color3);
            runProperties3.Append(fontSize6);
            runProperties3.Append(fontSizeComplexScript6);
            runProperties3.Append(languages6);
            Text text3 = new Text();
            text3.Text = "(Department)";

            run3.Append(runProperties3);
            run3.Append(text3);

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(run3);

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph2);
            tableCell2.Append(paragraph3);

            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "1258", Type = TableWidthUnitValues.Pct };

            TableCellBorders tableCellBorders3 = new TableCellBorders();
            TopBorder topBorder3 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder3 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
            RightBorder rightBorder3 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };

            tableCellBorders3.Append(topBorder3);
            tableCellBorders3.Append(bottomBorder3);
            tableCellBorders3.Append(rightBorder3);

            TableCellMargin tableCellMargin3 = new TableCellMargin();
            TopMargin topMargin3 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin3 = new LeftMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin3 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin3 = new RightMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };

            tableCellMargin3.Append(topMargin3);
            tableCellMargin3.Append(leftMargin3);
            tableCellMargin3.Append(bottomMargin3);
            tableCellMargin3.Append(rightMargin3);
            HideMark hideMark3 = new HideMark();

            tableCellProperties3.Append(tableCellWidth3);
            tableCellProperties3.Append(tableCellBorders3);
            tableCellProperties3.Append(tableCellMargin3);
            tableCellProperties3.Append(hideMark3);

            Paragraph paragraph4 = new Paragraph() { RsidParagraphMarkRevision = "00600789", RsidParagraphAddition = "00600789", RsidParagraphProperties = "00A461D4", RsidRunAdditionDefault = "00600789" };

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines4 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Justification justification4 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
            RunFonts runFonts7 = new RunFonts() { Ascii = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            FontSize fontSize7 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "24" };
            Languages languages7 = new Languages() { EastAsia = "en-IN" };

            paragraphMarkRunProperties4.Append(runFonts7);
            paragraphMarkRunProperties4.Append(fontSize7);
            paragraphMarkRunProperties4.Append(fontSizeComplexScript7);
            paragraphMarkRunProperties4.Append(languages7);

            paragraphProperties4.Append(spacingBetweenLines4);
            paragraphProperties4.Append(justification4);
            paragraphProperties4.Append(paragraphMarkRunProperties4);

            Run run4 = new Run() { RsidRunProperties = "00600789" };

            RunProperties runProperties4 = new RunProperties();
            RunFonts runFonts8 = new RunFonts() { Ascii = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            Bold bold4 = new Bold();
            BoldComplexScript boldComplexScript4 = new BoldComplexScript();
            Color color4 = new Color() { Val = "000000" };
            FontSize fontSize8 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "24" };
            Languages languages8 = new Languages() { EastAsia = "en-IN" };

            runProperties4.Append(runFonts8);
            runProperties4.Append(bold4);
            runProperties4.Append(boldComplexScript4);
            runProperties4.Append(color4);
            runProperties4.Append(fontSize8);
            runProperties4.Append(fontSizeComplexScript8);
            runProperties4.Append(languages8);
            Text text4 = new Text();
            text4.Text = "Designation";

            run4.Append(runProperties4);
            run4.Append(text4);

            paragraph4.Append(paragraphProperties4);
            paragraph4.Append(run4);

            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph4);

            TableCell tableCell4 = new TableCell();

            TableCellProperties tableCellProperties4 = new TableCellProperties();
            TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "1212", Type = TableWidthUnitValues.Pct };

            TableCellBorders tableCellBorders4 = new TableCellBorders();
            TopBorder topBorder4 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder4 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
            RightBorder rightBorder4 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };

            tableCellBorders4.Append(topBorder4);
            tableCellBorders4.Append(bottomBorder4);
            tableCellBorders4.Append(rightBorder4);

            TableCellMargin tableCellMargin4 = new TableCellMargin();
            TopMargin topMargin4 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin4 = new LeftMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin4 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin4 = new RightMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };

            tableCellMargin4.Append(topMargin4);
            tableCellMargin4.Append(leftMargin4);
            tableCellMargin4.Append(bottomMargin4);
            tableCellMargin4.Append(rightMargin4);
            HideMark hideMark4 = new HideMark();

            tableCellProperties4.Append(tableCellWidth4);
            tableCellProperties4.Append(tableCellBorders4);
            tableCellProperties4.Append(tableCellMargin4);
            tableCellProperties4.Append(hideMark4);

            Paragraph paragraph5 = new Paragraph() { RsidParagraphMarkRevision = "00600789", RsidParagraphAddition = "00600789", RsidParagraphProperties = "00A461D4", RsidRunAdditionDefault = "00600789" };

            ParagraphProperties paragraphProperties5 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines5 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Justification justification5 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
            RunFonts runFonts9 = new RunFonts() { Ascii = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            FontSize fontSize9 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "24" };
            Languages languages9 = new Languages() { EastAsia = "en-IN" };

            paragraphMarkRunProperties5.Append(runFonts9);
            paragraphMarkRunProperties5.Append(fontSize9);
            paragraphMarkRunProperties5.Append(fontSizeComplexScript9);
            paragraphMarkRunProperties5.Append(languages9);

            paragraphProperties5.Append(spacingBetweenLines5);
            paragraphProperties5.Append(justification5);
            paragraphProperties5.Append(paragraphMarkRunProperties5);

            Run run5 = new Run() { RsidRunProperties = "00600789" };

            RunProperties runProperties5 = new RunProperties();
            RunFonts runFonts10 = new RunFonts() { Ascii = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            Bold bold5 = new Bold();
            BoldComplexScript boldComplexScript5 = new BoldComplexScript();
            Color color5 = new Color() { Val = "000000" };
            FontSize fontSize10 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "24" };
            Languages languages10 = new Languages() { EastAsia = "en-IN" };

            runProperties5.Append(runFonts10);
            runProperties5.Append(bold5);
            runProperties5.Append(boldComplexScript5);
            runProperties5.Append(color5);
            runProperties5.Append(fontSize10);
            runProperties5.Append(fontSizeComplexScript10);
            runProperties5.Append(languages10);
            Text text5 = new Text();
            text5.Text = "Date time";

            run5.Append(runProperties5);
            run5.Append(text5);

            paragraph5.Append(paragraphProperties5);
            paragraph5.Append(run5);

            tableCell4.Append(tableCellProperties4);
            tableCell4.Append(paragraph5);

            tableRow1.Append(tableRowProperties1);
            tableRow1.Append(tableCell1);
            tableRow1.Append(tableCell2);
            tableRow1.Append(tableCell3);
            tableRow1.Append(tableCell4);
            return tableRow1;
        }

        public TableRow Generate_Author_TableRow(clsSignatore oSign)
        {
            TableRow tableRow1 = new TableRow() { RsidTableRowMarkRevision = "00600789", RsidTableRowAddition = "00600789", RsidTableRowProperties = "00A6633F" };
            TableRowProperties tableRowProperties1 = new TableRowProperties();
            TableHeader tableHeader1 = new TableHeader();
            tableRowProperties1.Append(tableHeader1);
            TableCell tableCell1 = new TableCell();
            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "1211", Type = TableWidthUnitValues.Pct };
            TableCellBorders tableCellBorders1 = new TableCellBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
            tableCellBorders1.Append(topBorder1);
            tableCellBorders1.Append(leftBorder1);
            tableCellBorders1.Append(bottomBorder1);
            tableCellBorders1.Append(rightBorder1);
            TableCellMargin tableCellMargin1 = new TableCellMargin();
            TopMargin topMargin1 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin1 = new LeftMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin1 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin1 = new RightMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };
            tableCellMargin1.Append(topMargin1);
            tableCellMargin1.Append(leftMargin1);
            tableCellMargin1.Append(bottomMargin1);
            tableCellMargin1.Append(rightMargin1);
            HideMark hideMark1 = new HideMark();
            tableCellProperties1.Append(tableCellWidth1);
            tableCellProperties1.Append(tableCellBorders1);
            tableCellProperties1.Append(tableCellMargin1);
            tableCellProperties1.Append(hideMark1);
            Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "00600789", RsidParagraphAddition = "00600789", RsidParagraphProperties = "00A461D4", RsidRunAdditionDefault = "00600789" };
            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Justification justification1 = new Justification() { Val = JustificationValues.Center };
            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            RunFonts runFonts1 = new RunFonts() { Ascii = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            FontSize fontSize1 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "24" };
            Languages languages1 = new Languages() { EastAsia = "en-IN" };
            paragraphMarkRunProperties1.Append(runFonts1);
            paragraphMarkRunProperties1.Append(fontSize1);
            paragraphMarkRunProperties1.Append(fontSizeComplexScript1);
            paragraphMarkRunProperties1.Append(languages1);
            paragraphProperties1.Append(spacingBetweenLines1);
            paragraphProperties1.Append(justification1);
            paragraphProperties1.Append(paragraphMarkRunProperties1);
            Run run1 = new Run() { RsidRunProperties = "00600789" };
            RunProperties runProperties1 = new RunProperties();
            RunFonts runFonts2 = new RunFonts() { Ascii = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            Color color1 = new Color() { Val = "000000" };
            FontSize fontSize2 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "24" };
            Languages languages2 = new Languages() { EastAsia = "en-IN" };
            runProperties1.Append(runFonts2);
            runProperties1.Append(bold1);
            runProperties1.Append(boldComplexScript1);
            runProperties1.Append(color1);
            runProperties1.Append(fontSize2);
            runProperties1.Append(fontSizeComplexScript2);
            runProperties1.Append(languages2);
            Text text1 = new Text();
            text1.Text = "Prepared By";
            run1.Append(runProperties1);
            run1.Append(text1);
            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);
            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph1);
            TableCell tableCell2 = new TableCell();
            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "1319", Type = TableWidthUnitValues.Pct };
            TableCellBorders tableCellBorders2 = new TableCellBorders();
            BottomBorder bottomBorder2 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
            RightBorder rightBorder2 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };

            tableCellBorders2.Append(bottomBorder2);
            tableCellBorders2.Append(rightBorder2);

            TableCellMargin tableCellMargin2 = new TableCellMargin();
            TopMargin topMargin2 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin2 = new LeftMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin2 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin2 = new RightMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };

            tableCellMargin2.Append(topMargin2);
            tableCellMargin2.Append(leftMargin2);
            tableCellMargin2.Append(bottomMargin2);
            tableCellMargin2.Append(rightMargin2);
            HideMark hideMark2 = new HideMark();

            tableCellProperties2.Append(tableCellWidth2);
            tableCellProperties2.Append(tableCellBorders2);
            tableCellProperties2.Append(tableCellMargin2);
            tableCellProperties2.Append(hideMark2);

            Paragraph paragraph2 = new Paragraph() { RsidParagraphMarkRevision = "00600789", RsidParagraphAddition = "00600789", RsidParagraphProperties = "00A461D4", RsidRunAdditionDefault = "00600789" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Justification justification2 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
            RunFonts runFonts3 = new RunFonts() { Ascii = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            FontSize fontSize3 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "24" };
            Languages languages3 = new Languages() { EastAsia = "en-IN" };

            paragraphMarkRunProperties2.Append(runFonts3);
            paragraphMarkRunProperties2.Append(fontSize3);
            paragraphMarkRunProperties2.Append(fontSizeComplexScript3);
            paragraphMarkRunProperties2.Append(languages3);

            paragraphProperties2.Append(spacingBetweenLines2);
            paragraphProperties2.Append(justification2);
            paragraphProperties2.Append(paragraphMarkRunProperties2);

            Run run2 = new Run() { RsidRunProperties = "00600789" };

            RunProperties runProperties2 = new RunProperties();
            RunFonts runFonts4 = new RunFonts() { Ascii = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            Color color2 = new Color() { Val = "000000" };
            FontSize fontSize4 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "24" };
            Languages languages4 = new Languages() { EastAsia = "en-IN" };

            runProperties2.Append(runFonts4);
            runProperties2.Append(color2);
            runProperties2.Append(fontSize4);
            runProperties2.Append(fontSizeComplexScript4);
            runProperties2.Append(languages4);
            Text text2 = new Text();
            text2.Text = oSign.UserFullName;

            run2.Append(runProperties2);
            run2.Append(text2);

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(run2);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphMarkRevision = "00600789", RsidParagraphAddition = "00600789", RsidParagraphProperties = "00A461D4", RsidRunAdditionDefault = "00600789" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Justification justification3 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
            RunFonts runFonts5 = new RunFonts() { Ascii = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            FontSize fontSize5 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "24" };
            Languages languages5 = new Languages() { EastAsia = "en-IN" };

            paragraphMarkRunProperties3.Append(runFonts5);
            paragraphMarkRunProperties3.Append(fontSize5);
            paragraphMarkRunProperties3.Append(fontSizeComplexScript5);
            paragraphMarkRunProperties3.Append(languages5);

            paragraphProperties3.Append(spacingBetweenLines3);
            paragraphProperties3.Append(justification3);
            paragraphProperties3.Append(paragraphMarkRunProperties3);

            Run run3 = new Run() { RsidRunProperties = "00600789" };

            RunProperties runProperties3 = new RunProperties();
            RunFonts runFonts6 = new RunFonts() { Ascii = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            Color color3 = new Color() { Val = "000000" };
            FontSize fontSize6 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "24" };
            Languages languages6 = new Languages() { EastAsia = "en-IN" };

            runProperties3.Append(runFonts6);
            runProperties3.Append(color3);
            runProperties3.Append(fontSize6);
            runProperties3.Append(fontSizeComplexScript6);
            runProperties3.Append(languages6);
            Text text3 = new Text();
            text3.Text = "(" + oSign.UserDepartment + ")";

            run3.Append(runProperties3);
            run3.Append(text3);

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(run3);

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph2);
            tableCell2.Append(paragraph3);

            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "1258", Type = TableWidthUnitValues.Pct };

            TableCellBorders tableCellBorders3 = new TableCellBorders();
            BottomBorder bottomBorder3 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
            RightBorder rightBorder3 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };

            tableCellBorders3.Append(bottomBorder3);
            tableCellBorders3.Append(rightBorder3);

            TableCellMargin tableCellMargin3 = new TableCellMargin();
            TopMargin topMargin3 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin3 = new LeftMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin3 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin3 = new RightMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };

            tableCellMargin3.Append(topMargin3);
            tableCellMargin3.Append(leftMargin3);
            tableCellMargin3.Append(bottomMargin3);
            tableCellMargin3.Append(rightMargin3);
            HideMark hideMark3 = new HideMark();

            tableCellProperties3.Append(tableCellWidth3);
            tableCellProperties3.Append(tableCellBorders3);
            tableCellProperties3.Append(tableCellMargin3);
            tableCellProperties3.Append(hideMark3);

            Paragraph paragraph4 = new Paragraph() { RsidParagraphMarkRevision = "00600789", RsidParagraphAddition = "00600789", RsidParagraphProperties = "00A461D4", RsidRunAdditionDefault = "00600789" };

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines4 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Justification justification4 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
            RunFonts runFonts7 = new RunFonts() { Ascii = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            FontSize fontSize7 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "24" };
            Languages languages7 = new Languages() { EastAsia = "en-IN" };

            paragraphMarkRunProperties4.Append(runFonts7);
            paragraphMarkRunProperties4.Append(fontSize7);
            paragraphMarkRunProperties4.Append(fontSizeComplexScript7);
            paragraphMarkRunProperties4.Append(languages7);

            paragraphProperties4.Append(spacingBetweenLines4);
            paragraphProperties4.Append(justification4);
            paragraphProperties4.Append(paragraphMarkRunProperties4);

            Run run4 = new Run() { RsidRunProperties = "00600789" };

            RunProperties runProperties4 = new RunProperties();
            RunFonts runFonts8 = new RunFonts() { Ascii = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            Color color4 = new Color() { Val = "000000" };
            FontSize fontSize8 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "24" };
            Languages languages8 = new Languages() { EastAsia = "en-IN" };

            runProperties4.Append(runFonts8);
            runProperties4.Append(color4);
            runProperties4.Append(fontSize8);
            runProperties4.Append(fontSizeComplexScript8);
            runProperties4.Append(languages8);
            Text text4 = new Text();
            text4.Text = oSign.UserDesignation;

            run4.Append(runProperties4);
            run4.Append(text4);

            paragraph4.Append(paragraphProperties4);
            paragraph4.Append(run4);

            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph4);

            TableCell tableCell4 = new TableCell();

            TableCellProperties tableCellProperties4 = new TableCellProperties();
            TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "1212", Type = TableWidthUnitValues.Pct };

            TableCellBorders tableCellBorders4 = new TableCellBorders();
            BottomBorder bottomBorder4 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
            RightBorder rightBorder4 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };

            tableCellBorders4.Append(bottomBorder4);
            tableCellBorders4.Append(rightBorder4);

            TableCellMargin tableCellMargin4 = new TableCellMargin();
            TopMargin topMargin4 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin4 = new LeftMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin4 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin4 = new RightMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };

            tableCellMargin4.Append(topMargin4);
            tableCellMargin4.Append(leftMargin4);
            tableCellMargin4.Append(bottomMargin4);
            tableCellMargin4.Append(rightMargin4);
            HideMark hideMark4 = new HideMark();

            tableCellProperties4.Append(tableCellWidth4);
            tableCellProperties4.Append(tableCellBorders4);
            tableCellProperties4.Append(tableCellMargin4);
            tableCellProperties4.Append(hideMark4);

            Paragraph paragraph5 = new Paragraph() { RsidParagraphMarkRevision = "00600789", RsidParagraphAddition = "00600789", RsidParagraphProperties = "00A461D4", RsidRunAdditionDefault = "00600789" };

            ParagraphProperties paragraphProperties5 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines5 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Justification justification5 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
            RunFonts runFonts9 = new RunFonts() { Ascii = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            FontSize fontSize9 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "24" };
            Languages languages9 = new Languages() { EastAsia = "en-IN" };

            paragraphMarkRunProperties5.Append(runFonts9);
            paragraphMarkRunProperties5.Append(fontSize9);
            paragraphMarkRunProperties5.Append(fontSizeComplexScript9);
            paragraphMarkRunProperties5.Append(languages9);

            paragraphProperties5.Append(spacingBetweenLines5);
            paragraphProperties5.Append(justification5);
            paragraphProperties5.Append(paragraphMarkRunProperties5);

            Run run5 = new Run() { RsidRunProperties = "00600789" };

            RunProperties runProperties5 = new RunProperties();
            RunFonts runFonts10 = new RunFonts() { Ascii = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            Color color5 = new Color() { Val = "000000" };
            FontSize fontSize10 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "24" };
            Languages languages10 = new Languages() { EastAsia = "en-IN" };

            runProperties5.Append(runFonts10);
            runProperties5.Append(color5);
            runProperties5.Append(fontSize10);
            runProperties5.Append(fontSizeComplexScript10);
            runProperties5.Append(languages10);
            Text text5 = new Text();
            text5.Text = oSign.UserDate.ToString(_szDateFormat) + " " + oSign.UserTime.ToString("T");
            //..oReviewr[i].UserDate.ToString(_szDateFormat) +" " + oReviewr[i].UserTime.ToString("T"))

            run5.Append(runProperties5);
            run5.Append(text5);

            paragraph5.Append(paragraphProperties5);
            paragraph5.Append(run5);

            tableCell4.Append(tableCellProperties4);
            tableCell4.Append(paragraph5);

            tableRow1.Append(tableRowProperties1);
            tableRow1.Append(tableCell1);
            tableRow1.Append(tableCell2);
            tableRow1.Append(tableCell3);
            tableRow1.Append(tableCell4);
            return tableRow1;
        }
        public TableRow Generate_Reviewer_Approver_TableRow(clsSignatore oSign)
        {
            TableRow tableRow1 = new TableRow() { RsidTableRowMarkRevision = "00600789", RsidTableRowAddition = "00600789", RsidTableRowProperties = "00A6633F" };
            TableRowProperties tableRowProperties1 = new TableRowProperties();
            TableHeader tableHeader1 = new TableHeader();
            tableRowProperties1.Append(tableHeader1);
            TableCell tableCell1 = new TableCell();
            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "1211", Type = TableWidthUnitValues.Pct };
            TableCellBorders tableCellBorders1 = new TableCellBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
            tableCellBorders1.Append(topBorder1);
            tableCellBorders1.Append(leftBorder1);
            tableCellBorders1.Append(bottomBorder1);
            tableCellBorders1.Append(rightBorder1);
            TableCellMargin tableCellMargin1 = new TableCellMargin();
            TopMargin topMargin1 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin1 = new LeftMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin1 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin1 = new RightMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };
            tableCellMargin1.Append(topMargin1);
            tableCellMargin1.Append(leftMargin1);
            tableCellMargin1.Append(bottomMargin1);
            tableCellMargin1.Append(rightMargin1);
            HideMark hideMark1 = new HideMark();
            tableCellProperties1.Append(tableCellWidth1);
            tableCellProperties1.Append(tableCellBorders1);
            tableCellProperties1.Append(tableCellMargin1);
            tableCellProperties1.Append(hideMark1);
            Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "00600789", RsidParagraphAddition = "00600789", RsidParagraphProperties = "00A461D4", RsidRunAdditionDefault = "00600789" };
            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Justification justification1 = new Justification() { Val = JustificationValues.Center };
            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            RunFonts runFonts1 = new RunFonts() { Ascii = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            FontSize fontSize1 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "24" };
            Languages languages1 = new Languages() { EastAsia = "en-IN" };
            paragraphMarkRunProperties1.Append(runFonts1);
            paragraphMarkRunProperties1.Append(fontSize1);
            paragraphMarkRunProperties1.Append(fontSizeComplexScript1);
            paragraphMarkRunProperties1.Append(languages1);
            paragraphProperties1.Append(spacingBetweenLines1);
            paragraphProperties1.Append(justification1);
            paragraphProperties1.Append(paragraphMarkRunProperties1);
            Run run1 = new Run() { RsidRunProperties = "00600789" };
            RunProperties runProperties1 = new RunProperties();
            RunFonts runFonts2 = new RunFonts() { Ascii = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            Color color1 = new Color() { Val = "000000" };
            FontSize fontSize2 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "24" };
            Languages languages2 = new Languages() { EastAsia = "en-IN" };
            runProperties1.Append(runFonts2);
            runProperties1.Append(bold1);
            runProperties1.Append(boldComplexScript1);
            runProperties1.Append(color1);
            runProperties1.Append(fontSize2);
            runProperties1.Append(fontSizeComplexScript2);
            runProperties1.Append(languages2);
            Text text1 = new Text();
            if (oSign.UserType == "R")
            {
                text1.Text = "Reviewed By";
            }
            else
            {
                text1.Text = "Approved By";
            }
            run1.Append(runProperties1);
            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);
            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph1);
            TableCell tableCell2 = new TableCell();
            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "1319", Type = TableWidthUnitValues.Pct };
            TableCellBorders tableCellBorders2 = new TableCellBorders();
            BottomBorder bottomBorder2 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
            RightBorder rightBorder2 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };

            tableCellBorders2.Append(bottomBorder2);
            tableCellBorders2.Append(rightBorder2);

            TableCellMargin tableCellMargin2 = new TableCellMargin();
            TopMargin topMargin2 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin2 = new LeftMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin2 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin2 = new RightMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };

            tableCellMargin2.Append(topMargin2);
            tableCellMargin2.Append(leftMargin2);
            tableCellMargin2.Append(bottomMargin2);
            tableCellMargin2.Append(rightMargin2);
            HideMark hideMark2 = new HideMark();

            tableCellProperties2.Append(tableCellWidth2);
            tableCellProperties2.Append(tableCellBorders2);
            tableCellProperties2.Append(tableCellMargin2);
            tableCellProperties2.Append(hideMark2);

            Paragraph paragraph2 = new Paragraph() { RsidParagraphMarkRevision = "00600789", RsidParagraphAddition = "00600789", RsidParagraphProperties = "00A461D4", RsidRunAdditionDefault = "00600789" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Justification justification2 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
            RunFonts runFonts3 = new RunFonts() { Ascii = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            FontSize fontSize3 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "24" };
            Languages languages3 = new Languages() { EastAsia = "en-IN" };

            paragraphMarkRunProperties2.Append(runFonts3);
            paragraphMarkRunProperties2.Append(fontSize3);
            paragraphMarkRunProperties2.Append(fontSizeComplexScript3);
            paragraphMarkRunProperties2.Append(languages3);

            paragraphProperties2.Append(spacingBetweenLines2);
            paragraphProperties2.Append(justification2);
            paragraphProperties2.Append(paragraphMarkRunProperties2);

            Run run2 = new Run() { RsidRunProperties = "00600789" };

            RunProperties runProperties2 = new RunProperties();
            RunFonts runFonts4 = new RunFonts() { Ascii = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            Color color2 = new Color() { Val = "000000" };
            FontSize fontSize4 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "24" };
            Languages languages4 = new Languages() { EastAsia = "en-IN" };

            runProperties2.Append(runFonts4);
            runProperties2.Append(color2);
            runProperties2.Append(fontSize4);
            runProperties2.Append(fontSizeComplexScript4);
            runProperties2.Append(languages4);
            Text text2 = new Text();
            text2.Text = oSign.UserFullName;

            run2.Append(runProperties2);
            run2.Append(text2);

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(run2);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphMarkRevision = "00600789", RsidParagraphAddition = "00600789", RsidParagraphProperties = "00A461D4", RsidRunAdditionDefault = "00600789" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Justification justification3 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
            RunFonts runFonts5 = new RunFonts() { Ascii = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            FontSize fontSize5 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "24" };
            Languages languages5 = new Languages() { EastAsia = "en-IN" };

            paragraphMarkRunProperties3.Append(runFonts5);
            paragraphMarkRunProperties3.Append(fontSize5);
            paragraphMarkRunProperties3.Append(fontSizeComplexScript5);
            paragraphMarkRunProperties3.Append(languages5);

            paragraphProperties3.Append(spacingBetweenLines3);
            paragraphProperties3.Append(justification3);
            paragraphProperties3.Append(paragraphMarkRunProperties3);

            Run run3 = new Run() { RsidRunProperties = "00600789" };

            RunProperties runProperties3 = new RunProperties();
            RunFonts runFonts6 = new RunFonts() { Ascii = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            Color color3 = new Color() { Val = "000000" };
            FontSize fontSize6 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "24" };
            Languages languages6 = new Languages() { EastAsia = "en-IN" };

            runProperties3.Append(runFonts6);
            runProperties3.Append(color3);
            runProperties3.Append(fontSize6);
            runProperties3.Append(fontSizeComplexScript6);
            runProperties3.Append(languages6);
            Text text3 = new Text();
            text3.Text = "(" + oSign.UserDepartment + ")";

            run3.Append(runProperties3);
            run3.Append(text3);

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(run3);

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph2);
            tableCell2.Append(paragraph3);

            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "1258", Type = TableWidthUnitValues.Pct };

            TableCellBorders tableCellBorders3 = new TableCellBorders();
            BottomBorder bottomBorder3 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
            RightBorder rightBorder3 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };

            tableCellBorders3.Append(bottomBorder3);
            tableCellBorders3.Append(rightBorder3);

            TableCellMargin tableCellMargin3 = new TableCellMargin();
            TopMargin topMargin3 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin3 = new LeftMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin3 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin3 = new RightMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };

            tableCellMargin3.Append(topMargin3);
            tableCellMargin3.Append(leftMargin3);
            tableCellMargin3.Append(bottomMargin3);
            tableCellMargin3.Append(rightMargin3);
            HideMark hideMark3 = new HideMark();

            tableCellProperties3.Append(tableCellWidth3);
            tableCellProperties3.Append(tableCellBorders3);
            tableCellProperties3.Append(tableCellMargin3);
            tableCellProperties3.Append(hideMark3);

            Paragraph paragraph4 = new Paragraph() { RsidParagraphMarkRevision = "00600789", RsidParagraphAddition = "00600789", RsidParagraphProperties = "00A461D4", RsidRunAdditionDefault = "00600789" };

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines4 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Justification justification4 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
            RunFonts runFonts7 = new RunFonts() { Ascii = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            FontSize fontSize7 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "24" };
            Languages languages7 = new Languages() { EastAsia = "en-IN" };

            paragraphMarkRunProperties4.Append(runFonts7);
            paragraphMarkRunProperties4.Append(fontSize7);
            paragraphMarkRunProperties4.Append(fontSizeComplexScript7);
            paragraphMarkRunProperties4.Append(languages7);

            paragraphProperties4.Append(spacingBetweenLines4);
            paragraphProperties4.Append(justification4);
            paragraphProperties4.Append(paragraphMarkRunProperties4);

            Run run4 = new Run() { RsidRunProperties = "00600789" };

            RunProperties runProperties4 = new RunProperties();
            RunFonts runFonts8 = new RunFonts() { Ascii = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            Color color4 = new Color() { Val = "000000" };
            FontSize fontSize8 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "24" };
            Languages languages8 = new Languages() { EastAsia = "en-IN" };

            runProperties4.Append(runFonts8);
            runProperties4.Append(color4);
            runProperties4.Append(fontSize8);
            runProperties4.Append(fontSizeComplexScript8);
            runProperties4.Append(languages8);
            Text text4 = new Text();
            text4.Text = oSign.UserDesignation;

            run4.Append(runProperties4);
            run4.Append(text4);

            paragraph4.Append(paragraphProperties4);
            paragraph4.Append(run4);

            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph4);

            TableCell tableCell4 = new TableCell();

            TableCellProperties tableCellProperties4 = new TableCellProperties();
            TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "1212", Type = TableWidthUnitValues.Pct };

            TableCellBorders tableCellBorders4 = new TableCellBorders();
            BottomBorder bottomBorder4 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
            RightBorder rightBorder4 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };

            tableCellBorders4.Append(bottomBorder4);
            tableCellBorders4.Append(rightBorder4);

            TableCellMargin tableCellMargin4 = new TableCellMargin();
            TopMargin topMargin4 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin4 = new LeftMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin4 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin4 = new RightMargin() { Width = "108", Type = TableWidthUnitValues.Dxa };

            tableCellMargin4.Append(topMargin4);
            tableCellMargin4.Append(leftMargin4);
            tableCellMargin4.Append(bottomMargin4);
            tableCellMargin4.Append(rightMargin4);
            HideMark hideMark4 = new HideMark();

            tableCellProperties4.Append(tableCellWidth4);
            tableCellProperties4.Append(tableCellBorders4);
            tableCellProperties4.Append(tableCellMargin4);
            tableCellProperties4.Append(hideMark4);

            Paragraph paragraph5 = new Paragraph() { RsidParagraphMarkRevision = "00600789", RsidParagraphAddition = "00600789", RsidParagraphProperties = "00A461D4", RsidRunAdditionDefault = "00600789" };

            ParagraphProperties paragraphProperties5 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines5 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Justification justification5 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
            RunFonts runFonts9 = new RunFonts() { Ascii = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            FontSize fontSize9 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "24" };
            Languages languages9 = new Languages() { EastAsia = "en-IN" };

            paragraphMarkRunProperties5.Append(runFonts9);
            paragraphMarkRunProperties5.Append(fontSize9);
            paragraphMarkRunProperties5.Append(fontSizeComplexScript9);
            paragraphMarkRunProperties5.Append(languages9);

            paragraphProperties5.Append(spacingBetweenLines5);
            paragraphProperties5.Append(justification5);
            paragraphProperties5.Append(paragraphMarkRunProperties5);

            Run run5 = new Run() { RsidRunProperties = "00600789" };

            RunProperties runProperties5 = new RunProperties();
            RunFonts runFonts10 = new RunFonts() { Ascii = "Calibri", ComplexScriptTheme = ThemeFontValues.MinorHighAnsi };
            Color color5 = new Color() { Val = "000000" };
            FontSize fontSize10 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "24" };
            Languages languages10 = new Languages() { EastAsia = "en-IN" };

            runProperties5.Append(runFonts10);
            runProperties5.Append(color5);
            runProperties5.Append(fontSize10);
            runProperties5.Append(fontSizeComplexScript10);
            runProperties5.Append(languages10);
            Text text5 = new Text();
            text5.Text = oSign.UserDate.ToString(_szDateFormat) + " " + oSign.UserTime.ToString("T");

            run5.Append(runProperties5);
            run5.Append(text5);

            paragraph5.Append(paragraphProperties5);
            paragraph5.Append(run5);

            tableCell4.Append(tableCellProperties4);
            tableCell4.Append(paragraph5);

            tableRow1.Append(tableRowProperties1);
            tableRow1.Append(tableCell1);
            tableRow1.Append(tableCell2);
            tableRow1.Append(tableCell3);
            tableRow1.Append(tableCell4);
            return tableRow1;
        }


        private static void Add_BookMark(Paragraph firstParagraph, string szBookMarkName)
        {
            string szFirstWordOfPara = string.Empty;
            string id = "0";
            BookmarkStart bookmarkStart1 = new BookmarkStart() { Name = szBookMarkName, Id = id };
            BookmarkEnd bookmarkEnd1 = new BookmarkEnd() { Id = id };
            firstParagraph.PrependChild(bookmarkStart1);
            firstParagraph.Append(bookmarkEnd1);
            //firstParagraph.Append(oPrevProp);
            id = id + 1;

        }

        private static void Add_BookMark_For_LastPage(Paragraph firstParagraph, string szBookMarkName)
        {
            string szFirstWordOfPara = string.Empty;
            string id = "0";
            var properties = firstParagraph
          .Descendants<Run>()
          .First()
          .Clone();
            BookmarkStart bookmarkStart1 = new BookmarkStart() { Name = szBookMarkName, Id = id };
            BookmarkEnd bookmarkEnd1 = new BookmarkEnd() { Id = id };
            firstParagraph.PrependChild(bookmarkStart1);
            firstParagraph.Append(bookmarkEnd1);
            id = id + 1;

        }

        #endregion


        #region .... IDISPOSABLE ....

        public void Dispose()
        {
            Dispose(true);
        }
        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
            }
            else
            {

            }
        }

        ~ClseAdd_Signatory_Page()
        {
            Dispose(false);
        }


        #endregion


    }
}
