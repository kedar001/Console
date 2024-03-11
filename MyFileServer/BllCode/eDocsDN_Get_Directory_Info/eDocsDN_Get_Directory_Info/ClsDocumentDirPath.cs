using eDocsDN_ReadAppXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace eDocsDN_Get_Directory_Info
{
    public class ClsDocumentDirPath : IDisposable
    {
        #region .... Variables Declaration ....

        clsReadAppXml _objReadXml = null;
        Directory_Attributes _oDirInfo = null;

        string _szAppXmlpath = string.Empty;
        string _szLocation = string.Empty;
        string _szDepartment = string.Empty;
        string _szPath = string.Empty;

        bool _bResult = false;

        #endregion

        #region .... Constructor ....

        public ClsDocumentDirPath(string szAppXmlPath, string szDBName, string szLocation, string szDepartment)
        {
            SetVariables(szAppXmlPath, szLocation, szDepartment, true);
        }

        #endregion

        #region .... Functions Definition ....

        private void SetVariables(string szAppXmlPath, string szLocation, string szDepartment, bool isBackEnd)
        {
            _szLocation = "";
            _szDepartment = "";
            _szAppXmlpath = szAppXmlPath;
            _objReadXml = new clsReadAppXml(_szAppXmlpath);
            _objReadXml.IsBackEnd = isBackEnd;
            _szLocation = szLocation;
            _szDepartment = szDepartment;
        }


        #endregion

        #region .... Properties ....

        public string WorkingDoc
        {
            get
            {
                _szPath = _objReadXml.GetAppDirPath(_szLocation, _szDepartment, "WorkingDoc");
                return _szPath;
            }
        }

        public string WorkingHtm
        {
            get
            {
                _szPath = _objReadXml.GetAppDirPath(_szLocation, _szDepartment, "WorkingHtm");
                return _szPath;
            }
        }

        public string PublishDoc
        {
            get
            {
                _szPath = _objReadXml.GetAppDirPath(_szLocation, _szDepartment, "PublishDoc");
                return _szPath;
            }
        }

        public string PublishHtm
        {
            get
            {
                _szPath = _objReadXml.GetAppDirPath(_szLocation, _szDepartment, "PublishHtm");
                return _szPath;
            }
        }

        public string HistoryDoc
        {
            get
            {
                _szPath = _objReadXml.GetAppDirPath(_szLocation, _szDepartment, "HistoryDoc");
                return _szPath;
            }
        }

        public string HistoryHtm
        {
            get
            {
                _szPath = _objReadXml.GetAppDirPath(_szLocation, _szDepartment, "HistoryHtm");
                return _szPath;
            }
        }

        public string ChangeHistoryDoc
        {
            get
            {
                _szPath = _objReadXml.GetAppDirPath(_szLocation, _szDepartment, "ChangeHistoryDoc");
                return _szPath;
            }
        }

        public string ChangeHistoryHtm
        {
            get
            {
                _szPath = _objReadXml.GetAppDirPath(_szLocation, _szDepartment, "ChangeHistoryHtm");
                return _szPath;
            }
        }

        public string TempSharedPhysical
        {
            get
            {
                _bResult = _objReadXml.IsDefaultFs(_szLocation);
                if (_bResult)
                    _szPath = _objReadXml.GetPhysicalPath("PhysicalTempShared");
                else
                    _szPath = _objReadXml.GetAppDirPath(_szLocation, _szDepartment, "PhysicalTempShared");
                return _szPath;
            }
        }

        public string TempShared
        {
            get
            {
                _szPath = _objReadXml.GetLocationVariable(_szLocation, _szDepartment, "TempShared");
                return _szPath;
            }
        }

        public string TempDir
        {
            get
            {
                _bResult = _objReadXml.IsDefaultFs(_szLocation);
                if (_bResult)
                    _szPath = _objReadXml.GetPhysicalPath("TempDir");
                else
                    _szPath = _objReadXml.GetAppDirPath(_szLocation, _szDepartment, "TempDir");
                return _szPath;
            }
        }

        public string VirtualTempDir
        {
            get
            {
                _szPath = _objReadXml.GetLocationVariable(_szLocation, _szDepartment, "VirtualTempDir");
                return _szPath;

            }
        }

        public string VirtualDir
        {
            get
            {
                _szPath = _objReadXml.GetLocationVariable(_szLocation, _szDepartment, "VirtualDir");
                return _szPath;
            }
        }

        public string VirtualShared
        {
            get
            {
                _szPath = _objReadXml.GetLocationVariable(_szLocation, _szDepartment, "VirtualSharedDir");
                return _szPath;
            }
        }

        public string TemplateWord
        {
            get
            {
                _szPath = _objReadXml.GetAppDirPath(_szLocation, _szDepartment, "TemplateWord");
                return _szPath;
            }
        }

        public string TemplateExcel
        {
            get
            {
                _szPath = _objReadXml.GetAppDirPath(_szLocation, _szDepartment, "TemplateExcel");
                return _szPath;
            }
        }

        public string TemplateOther
        {
            get
            {
                _szPath = _objReadXml.GetAppDirPath(_szLocation, _szDepartment, "TemplateOther");
                return _szPath;
            }
        }

        public string DcrFiles
        {
            get
            {
                _szPath = _objReadXml.GetAppDirPath(_szLocation, _szDepartment, "DcrFiles");
                return _szPath;
            }
        }

        public string Migration
        {
            get
            {
                _szPath = _objReadXml.GetAppDirPath(_szLocation, _szDepartment, "Migration");
                return _szPath;
            }
        }

        public string Download
        {
            get
            {
                _szPath = _objReadXml.GetAppDirPath(_szLocation, _szDepartment, "Download");
                return _szPath;
            }
        }

        public string PrintTemplate
        {
            get
            {
                _szPath = _objReadXml.GetAppDirPath(_szLocation, _szDepartment, "PrintTemplate");
                return _szPath;
            }
        }
        public string DraftVersion
        {
            get
            {
                _szPath = _objReadXml.GetAppDirPath(_szLocation, _szDepartment, "DraftVersion");
                return _szPath;
            }
        }
        public string Preview
        {
            get
            {
                _szPath = _objReadXml.GetAppDirPath(_szLocation, _szDepartment, "Preview");
                return _szPath;
            }
        }
        public string Deleted
        {
            get
            {
                _szPath = _objReadXml.GetAppDirPath(_szLocation, _szDepartment, "Deleted");
                return _szPath;
            }
        }
        public string SupportingFiles
        {
            get
            {
                _szPath = _objReadXml.GetAppDirPath(_szLocation, _szDepartment, "SupportingFiles");
                return _szPath;
            }
        }
        public string Draft_Template_Version
        {
            get
            {
                _szPath = _objReadXml.GetAppDirPath(_szLocation, _szDepartment, "DraftTemplateVersion");
                return _szPath;
            }
        }
        public string Deleted_DCR
        {
            get
            {
                _szPath = _objReadXml.GetAppDirPath(_szLocation, _szDepartment, "Deleted_DCR_Files");
                return _szPath;
            }
        }
        public string PublishPDF
        {
            get
            {
                _szPath = _objReadXml.GetAppDirPath(_szLocation, _szDepartment, "PublishPdf");
                return _szPath;
            }
        }

        public string HistoryPDF
        {
            get
            {
                _szPath = _objReadXml.GetAppDirPath(_szLocation, _szDepartment, "HistoryPdf");
                return _szPath;
            }
        }

        public string WorkingPDF
        {
            get
            {
                _szPath = _objReadXml.GetAppDirPath(_szLocation, _szDepartment, "WorkingPdf");
                return _szPath;
            }
        }

        public string ChangeHistoryHTMPdf
        {
            get
            {
                _szPath = _objReadXml.GetAppDirPath(_szLocation, _szDepartment, "ChangeHistoryHtmPdf");
                return _szPath;
            }
        }









        public Directory_Attributes GetDirPath(string szType, bool bStoreFilesinBlob, bool bisPhysicalStorage, bool bIsEcryptionEnabled)
        {

            switch (szType.ToUpper())
            {
                case "WD":
                    _oDirInfo = new Directory_Attributes();
                    _oDirInfo.Directory_Path = WorkingDoc;

                    if (bStoreFilesinBlob)
                    {
                        _oDirInfo.Database_Storage = true;
                        _oDirInfo.Table_Name = "zespl_cod_gnikrow";
                    }
                    if (bisPhysicalStorage)
                    {
                        if (bIsEcryptionEnabled)
                            _oDirInfo.Files_To_Be_Encrypted = false;
                        else
                            _oDirInfo.Files_To_Be_Encrypted = false;

                        _oDirInfo.Physical_Directory = true;
                    }
                    break;
                //case "WH":
                //    _oDirInfo = new Directory_Attributes();
                //    _oDirInfo.Directory_Path = WorkingHtm;

                //    if (bStoreFilesinBlob)
                //    {
                //        _oDirInfo.Physical_Directory = false;
                //        _oDirInfo.Database_Storage = true;
                //        _oDirInfo.Table_Name = "zespl_mth_gnikrow";
                //    }
                //    if (bisPhysicalStorage)
                //    {
                //        if (bIsEcryptionEnabled)
                //            _oDirInfo.Files_To_Be_Encrypted = false;
                //        else
                //            _oDirInfo.Files_To_Be_Encrypted = false;

                //        _oDirInfo.Physical_Directory = true;
                //        _oDirInfo.Database_Storage = false;
                //    }

                //break;
                case "PD":
                    _oDirInfo = new Directory_Attributes();
                    _oDirInfo.Directory_Path = PublishDoc;

                    if (bStoreFilesinBlob)
                    {
                        _oDirInfo.Database_Storage = true;
                        _oDirInfo.Table_Name = "zespl_cod_hsilbup";
                    }
                    if (bisPhysicalStorage)
                    {
                        _oDirInfo.Physical_Directory = true;
                        if (bIsEcryptionEnabled)
                            _oDirInfo.Files_To_Be_Encrypted = true;
                        else
                            _oDirInfo.Files_To_Be_Encrypted = false;
                    }
                    break;
                //case "PH":

                //    _oDirInfo = new Directory_Attributes();
                //    _oDirInfo.Directory_Path = PublishHtm;
                //    //_oDirInfo.Files_To_Be_Encrypted = true;
                //    //_oDirInfo.Physical_Directory = true;
                //    //_oDirInfo.Database_Storage = false;
                //    _oDirInfo.Table_Name = "NA";

                //    break;
                case "HD":
                    _oDirInfo = new Directory_Attributes();
                    _oDirInfo.Directory_Path = HistoryDoc;



                    if (bStoreFilesinBlob)
                    {
                        _oDirInfo.Database_Storage = true;
                        _oDirInfo.Table_Name = "zespl_cod_yrotsih";
                    }
                    if (bisPhysicalStorage)
                    {
                        _oDirInfo.Physical_Directory = true;
                        if (bIsEcryptionEnabled)
                            _oDirInfo.Files_To_Be_Encrypted = true;
                        else
                            _oDirInfo.Files_To_Be_Encrypted = false;

                    }

                    break;
                //case "HH":
                //    _oDirInfo = new Directory_Attributes();
                //    _oDirInfo.Directory_Path = HistoryHtm;
                //    if (bIsEcryptionEnabled)
                //        _oDirInfo.Files_To_Be_Encrypted = true;
                //    else
                //        _oDirInfo.Files_To_Be_Encrypted = false;

                //    if (bStoreFilesinBlob)
                //    {
                //        _oDirInfo.Database_Storage = true;
                //        _oDirInfo.Table_Name = "zespl_mth_yrotsih";
                //    }
                //    if (bisPhysicalStorage)
                //    {
                //        _oDirInfo.Physical_Directory = true;
                //    }
                //    break;
                case "CD":
                    _oDirInfo = new Directory_Attributes();
                    _oDirInfo.Directory_Path = ChangeHistoryDoc;


                    if (bStoreFilesinBlob)
                    {
                        _oDirInfo.Database_Storage = true;
                        _oDirInfo.Table_Name = "zespl_cod_yrotsih_degnahc";
                    }
                    if (bisPhysicalStorage)
                    {
                        _oDirInfo.Physical_Directory = true;
                        if (bIsEcryptionEnabled)
                            _oDirInfo.Files_To_Be_Encrypted = true;
                        else
                            _oDirInfo.Files_To_Be_Encrypted = false;
                    }

                    break;
                case "CH":
                    _oDirInfo = new Directory_Attributes();
                    _oDirInfo.Directory_Path = ChangeHistoryHtm;

                    if (bStoreFilesinBlob)
                    {
                        _oDirInfo.Database_Storage = true;
                        _oDirInfo.Table_Name = "zespl_mth_yrotsih_degnahc";
                    }
                    if (bisPhysicalStorage)
                    {
                        _oDirInfo.Physical_Directory = true;
                        if (bIsEcryptionEnabled)
                            _oDirInfo.Files_To_Be_Encrypted = true;
                        else
                            _oDirInfo.Files_To_Be_Encrypted = false;
                    }
                    break;
                case "TS":
                    _oDirInfo = new Directory_Attributes();
                    _oDirInfo.Directory_Path = TempSharedPhysical;
                    _oDirInfo.Files_To_Be_Encrypted = false;
                    _oDirInfo.Physical_Directory = true;
                    _oDirInfo.Database_Storage = false;


                    break;
                case "SP":
                    _oDirInfo = new Directory_Attributes();
                    _oDirInfo.Directory_Path = TempShared;
                    _oDirInfo.Files_To_Be_Encrypted = false;
                    _oDirInfo.Physical_Directory = true;
                    _oDirInfo.Database_Storage = false;


                    break;
                case "TD":
                    _oDirInfo = new Directory_Attributes();
                    _oDirInfo.Directory_Path = TempDir;
                    _oDirInfo.Files_To_Be_Encrypted = false;
                    _oDirInfo.Physical_Directory = true;
                    _oDirInfo.Database_Storage = false;

                    break;
                case "TW":
                    _oDirInfo = new Directory_Attributes();
                    _oDirInfo.Directory_Path = TemplateWord;

                    if (bStoreFilesinBlob)
                    {
                        _oDirInfo.Database_Storage = true;
                        _oDirInfo.Table_Name = "zespl_drow_etalpmet";
                    }
                    if (bisPhysicalStorage)
                    {
                        _oDirInfo.Physical_Directory = true;
                        if (bIsEcryptionEnabled)
                            _oDirInfo.Files_To_Be_Encrypted = true;
                        else
                            _oDirInfo.Files_To_Be_Encrypted = false;
                    }

                    break;
                case "TE":

                    _oDirInfo = new Directory_Attributes();
                    _oDirInfo.Directory_Path = TemplateExcel;
                    if (bStoreFilesinBlob)
                    {
                        _oDirInfo.Database_Storage = true;
                        _oDirInfo.Table_Name = "zespl_lecxe_etalpmet";
                    }
                    if (bisPhysicalStorage)
                    {
                        _oDirInfo.Physical_Directory = true;
                        if (bIsEcryptionEnabled)
                            _oDirInfo.Files_To_Be_Encrypted = true;
                        else
                            _oDirInfo.Files_To_Be_Encrypted = false;
                    }
                    break;
                case "TO":
                    _oDirInfo = new Directory_Attributes();
                    _oDirInfo.Directory_Path = TemplateOther;

                    if (bStoreFilesinBlob)
                    {
                        _oDirInfo.Database_Storage = true;
                        _oDirInfo.Table_Name = "zespl_rehto_etalpmet";
                    }
                    if (bisPhysicalStorage)
                    {
                        _oDirInfo.Physical_Directory = true;
                        if (bIsEcryptionEnabled)
                            _oDirInfo.Files_To_Be_Encrypted = true;
                        else
                            _oDirInfo.Files_To_Be_Encrypted = false;
                    }

                    break;
                case "DCR":
                    _oDirInfo = new Directory_Attributes();
                    _oDirInfo.Directory_Path = DcrFiles;

                    if (bStoreFilesinBlob)
                    {
                        _oDirInfo.Database_Storage = true;
                        _oDirInfo.Table_Name = "zespl_selif_rcd";
                    }
                    if (bisPhysicalStorage)
                    {
                        _oDirInfo.Physical_Directory = true;
                        if (bIsEcryptionEnabled)
                            _oDirInfo.Files_To_Be_Encrypted = true;
                        else
                            _oDirInfo.Files_To_Be_Encrypted = false;
                    }

                    break;
                case "DF":
                    _oDirInfo = new Directory_Attributes();
                    _oDirInfo.Directory_Path = Download;

                    if (bStoreFilesinBlob)
                    {
                        _oDirInfo.Database_Storage = true;
                        _oDirInfo.Table_Name = "zespl_selif_daolnwod";
                    }
                    if (bisPhysicalStorage)
                    {
                        _oDirInfo.Physical_Directory = true;
                        if (bIsEcryptionEnabled)
                            _oDirInfo.Files_To_Be_Encrypted = true;
                        else
                            _oDirInfo.Files_To_Be_Encrypted = false;
                    }

                    break;
                case "PT":
                    _oDirInfo = new Directory_Attributes();
                    _oDirInfo.Directory_Path = PrintTemplate;

                    if (bStoreFilesinBlob)
                    {
                        _oDirInfo.Database_Storage = true;
                        _oDirInfo.Table_Name = "zespl_etelpmet_tnirp";
                    }
                    if (bisPhysicalStorage)
                    {
                        if (bIsEcryptionEnabled)
                            _oDirInfo.Files_To_Be_Encrypted = true;
                        else
                            _oDirInfo.Files_To_Be_Encrypted = false;
                        _oDirInfo.Physical_Directory = true;
                    }

                    break;
                case "PV":
                    _oDirInfo = new Directory_Attributes();
                    _oDirInfo.Directory_Path = Preview;

                    if (bStoreFilesinBlob)
                    {
                        _oDirInfo.Database_Storage = true;
                        _oDirInfo.Table_Name = "zespl_weiverp";
                    }
                    if (bisPhysicalStorage)
                    {
                        if (bIsEcryptionEnabled)
                            _oDirInfo.Files_To_Be_Encrypted = true;
                        else
                            _oDirInfo.Files_To_Be_Encrypted = false;

                        _oDirInfo.Physical_Directory = true;
                    }

                    break;
                case "DV":
                    _oDirInfo = new Directory_Attributes();
                    _oDirInfo.Directory_Path = DraftVersion;

                    if (bStoreFilesinBlob)
                    {
                        _oDirInfo.Database_Storage = true;
                        _oDirInfo.Table_Name = "zespl_noisrev_tfard";
                    }
                    if (bisPhysicalStorage)
                    {
                        if (bIsEcryptionEnabled)
                            _oDirInfo.Files_To_Be_Encrypted = true;
                        else
                            _oDirInfo.Files_To_Be_Encrypted = false;
                        _oDirInfo.Physical_Directory = true;
                    }

                    break;
                case "DD":
                    _oDirInfo = new Directory_Attributes();
                    _oDirInfo.Directory_Path = Deleted;

                    if (bStoreFilesinBlob)
                    {
                        _oDirInfo.Database_Storage = true;
                        _oDirInfo.Table_Name = "zespl_stnemucod_deteled";
                    }
                    if (bisPhysicalStorage)
                    {
                        if (bIsEcryptionEnabled)
                            _oDirInfo.Files_To_Be_Encrypted = true;
                        else
                            _oDirInfo.Files_To_Be_Encrypted = false;
                        _oDirInfo.Physical_Directory = true;
                    }
                    break;
                case "SF":
                    _oDirInfo = new Directory_Attributes();
                    _oDirInfo.Directory_Path = SupportingFiles;

                    if (bStoreFilesinBlob)
                    {
                        _oDirInfo.Database_Storage = true;
                        _oDirInfo.Table_Name = "zespl_selif_gnitroppus";
                    }
                    if (bisPhysicalStorage)
                    {
                        if (bIsEcryptionEnabled)
                            _oDirInfo.Files_To_Be_Encrypted = true;
                        else
                            _oDirInfo.Files_To_Be_Encrypted = false;
                        _oDirInfo.Physical_Directory = true;
                    }
                    break;
                case "DT":
                    _oDirInfo = new Directory_Attributes();
                    _oDirInfo.Directory_Path = Draft_Template_Version;

                    if (bStoreFilesinBlob)
                    {
                        _oDirInfo.Database_Storage = true;
                        _oDirInfo.Table_Name = "zespl_rev_estalpmet_tfard";
                    }
                    if (bisPhysicalStorage)
                    {
                        if (bIsEcryptionEnabled)
                            _oDirInfo.Files_To_Be_Encrypted = true;
                        else
                            _oDirInfo.Files_To_Be_Encrypted = false;
                        _oDirInfo.Physical_Directory = true;
                    }
                    break;

                case "DDCR":
                    _oDirInfo = new Directory_Attributes();
                    _oDirInfo.Directory_Path = Deleted_DCR;

                    if (bStoreFilesinBlob)
                    {
                        _oDirInfo.Database_Storage = true;
                        _oDirInfo.Table_Name = "zespl_rcd_deteled";
                    }
                    if (bisPhysicalStorage)
                    {
                        _oDirInfo.Physical_Directory = true;
                        if (bIsEcryptionEnabled)
                            _oDirInfo.Files_To_Be_Encrypted = true;
                        else
                            _oDirInfo.Files_To_Be_Encrypted = false;

                    }
                    break;
                case "MI":
                    _oDirInfo = new Directory_Attributes();
                    _oDirInfo.Directory_Path = Migration;
                    _oDirInfo.Files_To_Be_Encrypted = false;
                    _oDirInfo.Physical_Directory = true;
                    _oDirInfo.Database_Storage = false;
                    break;

                case "PP":
                    _oDirInfo = new Directory_Attributes();
                    _oDirInfo.Directory_Path = PublishPDF;

                    if (bStoreFilesinBlob)
                    {
                        _oDirInfo.Database_Storage = true;
                        _oDirInfo.Table_Name = "zespl_fdp_hsilbup";
                    }
                    if (bisPhysicalStorage)
                    {
                        if (bIsEcryptionEnabled)
                            _oDirInfo.Files_To_Be_Encrypted = true;
                        else
                            _oDirInfo.Files_To_Be_Encrypted = false;
                        _oDirInfo.Physical_Directory = true;
                    }

                    break;
                case "HP":
                    _oDirInfo = new Directory_Attributes();
                    _oDirInfo.Directory_Path = HistoryPDF;

                    if (bStoreFilesinBlob)
                    {
                        _oDirInfo.Database_Storage = true;
                        _oDirInfo.Table_Name = "zespl_fdp_yrotsih";
                    }
                    if (bisPhysicalStorage)
                    {
                        _oDirInfo.Physical_Directory = true;
                        if (bIsEcryptionEnabled)
                            _oDirInfo.Files_To_Be_Encrypted = true;
                        else
                            _oDirInfo.Files_To_Be_Encrypted = false;
                    }

                    break;
                case "WP":
                    _oDirInfo = new Directory_Attributes();
                    _oDirInfo.Directory_Path = WorkingPDF;


                    if (bStoreFilesinBlob)
                    {
                        _oDirInfo.Database_Storage = true;
                        _oDirInfo.Table_Name = "zespl_fdp_gnikrow";
                    }
                    if (bisPhysicalStorage)
                    {
                        _oDirInfo.Physical_Directory = true;
                        if (bIsEcryptionEnabled)
                            _oDirInfo.Files_To_Be_Encrypted = true;
                        else
                            _oDirInfo.Files_To_Be_Encrypted = false;

                    }

                    break;
                case "CHP":
                    _oDirInfo = new Directory_Attributes();
                    _oDirInfo.Directory_Path = ChangeHistoryHTMPdf;


                    if (bStoreFilesinBlob)
                    {
                        _oDirInfo.Database_Storage = true;
                        _oDirInfo.Table_Name = "zespl_fdp_mth_yrotsih_degnahc";
                    }
                    if (bisPhysicalStorage)
                    {
                        _oDirInfo.Physical_Directory = true;
                        if (bIsEcryptionEnabled)
                            _oDirInfo.Files_To_Be_Encrypted = true;
                        else
                            _oDirInfo.Files_To_Be_Encrypted = false;

                    }

                    break;





            }
            return _oDirInfo;
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
                if (_objReadXml != null)
                    _objReadXml.Dispose();
                _objReadXml = null;
                _oDirInfo = null;
                _szAppXmlpath = string.Empty;
                _szLocation = string.Empty;
                _szDepartment = string.Empty;
                _szPath = string.Empty;
            }
            else
            {

            }
        }

        ~ClsDocumentDirPath()
        {
            Dispose(false);
        }


        #endregion
    }
}
