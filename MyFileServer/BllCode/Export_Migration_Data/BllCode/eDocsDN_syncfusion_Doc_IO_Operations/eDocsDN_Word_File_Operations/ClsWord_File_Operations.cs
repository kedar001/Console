using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.Diagnostics;
using System.IO;

namespace eDocsDN_syncfusion_Operations
{
    public class ClsXml_Operations
    {
        #region ..... Variable Declaration .....
        string _szFilePath = string.Empty;
        string _szAppXmlPath = string.Empty;
        string _szLogFileName = string.Empty;
        #endregion

        #region ..... Properties .....
        public string msgError { get; set; }
        #endregion

        #region ..... Constroctor .....
        public ClsXml_Operations(string szAppXmlPath)
        {
            msgError = string.Empty;
            _szAppXmlPath = szAppXmlPath;
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("NTM1NjY5QDMxMzkyZTMzMmUzMFdKd3BlNHR0bTdzYmhQMUROUHVKbVpIK2pIeTdVMWdESEZ1TlJQL0tTVTg9");
            _szLogFileName = AppDomain.CurrentDomain.BaseDirectory + @"\Log\FileOperationLog.txt";
        }
        #endregion

        #region .....Public Functions .....
        //public bool Convert_Document_To_PDF(string szFilePath)
        //{

        //    bool flag = true;
        //    DocToPDFConverter converter = null;
        //    try
        //    {
        //        using (WordDocument wordDocument = new WordDocument(szFilePath, FormatType.Docx))
        //        {
        //            wordDocument.Comments.Clear();
        //            converter = new DocToPDFConverter();
        //            PdfDocument pdfDocument = converter.ConvertToPDF(wordDocument);
        //            wordDocument.Close();
        //            pdfDocument.Save(Path.GetDirectoryName(szFilePath) + "\\" + Path.GetFileNameWithoutExtension(szFilePath) + ".pdf");
        //            pdfDocument.Close(true);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Add_Log("Error : Convert_Document_To_PDF : " + ex.Message, EventLogEntryType.Error);
        //        msgError = ex.Message;
        //    }
        //    finally
        //    {
        //        converter = null;
        //    }
        //    return flag;
        //}
        //public Stream Convert_Document_To_PDF(Stream strmFile)
        //{

        //    Stream strmPdfFile = null;

        //    DocToPDFConverter converter = null;
        //    try
        //    {

        //        using (WordDocument wordDocument = new WordDocument(strmFile, FormatType.Docx))
        //        {
        //            wordDocument.Comments.Clear();
        //            converter = new DocToPDFConverter();
        //            PdfDocument pdfDocument = converter.ConvertToPDF(wordDocument);
        //            wordDocument.Close();
        //            pdfDocument.Save(strmFile);
        //            pdfDocument.Close(true);
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        Add_Log("Error : Convert_Document_To_PDF : " + ex.Message, EventLogEntryType.Error);
        //        msgError = ex.Message;
        //        strmPdfFile = null;

        //    }
        //    finally
        //    {
        //        converter = null;
        //    }
        //    return strmPdfFile;
        //}

        public bool Update_Document_Properties(string szFilePath)
        {
            bool bResult = true;
            try
            {

                using (WordDocument wordDocument = new WordDocument(szFilePath, FormatType.Docx))
                {
                    //wordDocument.Comments.Clear();
                    wordDocument.UpdateDocumentFields();
                    wordDocument.Save(szFilePath, FormatType.Docx);
                    wordDocument.Close();
                }
            }
            catch (Exception ex)
            {
                Add_Log("Error : Update_Document_Properties : " + ex.Message, EventLogEntryType.Error);
                bResult = false;
                msgError = ex.Message;

            }
            finally
            {

            }
            return bResult;
        }

        public bool Process_Word_Document(string szFilePath)
        {
            bool bReturnVal = true;
            WordDocument document = null;

            try
            {
                using (document = new WordDocument(szFilePath, FormatType.Docx))
                {
                    document.EnsureMinimal();
                    document.Comments.Clear();
                    if (document.HasChanges)
                    {
                        document.Revisions.AcceptAll();
                    }
                    document.SaveOptions.MaintainCompatibilityMode = true;
                    document.Save(szFilePath, FormatType.Docx);
                    document.Close();
                }
            }
            catch (Exception ex)
            {
                Add_Log("Error : Process_Word_Document : " + ex.Message, EventLogEntryType.Error);
                bReturnVal = false;
                msgError = ex.Message;
            }
            finally { document = null; }
            return bReturnVal;
        }
        public Stream Process_Word_Document(Stream strmFile)
        {
            WordDocument document = null;
            try
            {
                using (document = new WordDocument(strmFile, FormatType.Docx))
                {
                    document.Comments.Clear();
                    if (document.HasChanges)
                    {
                        document.Revisions.AcceptAll();
                    }
                    document.SaveOptions.MaintainCompatibilityMode = true;
                    document.Save(strmFile, FormatType.Docx);
                    document.Close();
                }
            }
            catch (Exception ex)
            {
                Add_Log("Error : Process_Word_Document : " + ex.Message, EventLogEntryType.Error);
                strmFile = null;
                msgError = ex.Message;
            }
            finally { document = null; }
            return strmFile;
        }

        private void Add_Log(string szMessage, EventLogEntryType e)
        {
            using (EventLog eventLog = new EventLog("FileServer"))
            {
                eventLog.Source = "FileServer";
                eventLog.WriteEntry(szMessage, e, 101, 1);
            }
        }
        #endregion
    }
}
