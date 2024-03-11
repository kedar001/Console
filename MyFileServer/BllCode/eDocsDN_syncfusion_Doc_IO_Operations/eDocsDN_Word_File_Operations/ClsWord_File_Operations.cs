using eDocsDN_Get_Directory_Info;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocToPDFConverter;
using Syncfusion.Pdf;
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
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("NRAiBiAaIQQuGjN/V0Z+XU9EaFtFVmJLYVB3WmpQdldgdVRMZVVbQX9PIiBoS35RdEVlWXZecHRcRmRdVkJ3");
            _szLogFileName = AppDomain.CurrentDomain.BaseDirectory + @"\Log\FileOperationLog.txt";
        }
        #endregion

        #region .....Public Functions .....
        public bool Convert_Document_To_PDF(string szFilePath)
        {

            File.AppendAllText(_szLogFileName, "Convert_Document_To_PDF");
            bool flag = true;
            DocToPDFConverter converter = null;
            PdfDocument pdfDocument = null;
            WordDocument wordDocument = null;
            try
            {
                using (wordDocument = new WordDocument(szFilePath, FormatType.Docx))
                {
                    wordDocument.Comments.Clear();
                    converter = new DocToPDFConverter();
                    using (pdfDocument = converter.ConvertToPDF(wordDocument))
                    {
                        wordDocument.Close();
                        pdfDocument.Save(Path.GetDirectoryName(szFilePath) + "\\" + Path.GetFileNameWithoutExtension(szFilePath) + ".pdf");
                        pdfDocument.Close(true);
                    }
                }
            }
            catch (Exception ex)
            {
                File.AppendAllText(_szLogFileName, "Error : Convert_Document_To_PDF : " + ex.Message + EventLogEntryType.Error);
                msgError = ex.Message;
            }
            finally
            {
                converter = null;
                pdfDocument = null;
                wordDocument = null;
            }
            return flag;
        }
        public Stream Convert_Document_To_PDF(Stream strmFile)
        {

            Stream strmPdfFile = null;
            File.AppendAllText(_szLogFileName, "Convert_Document_To_PDF");
            DocToPDFConverter converter = null;
            PdfDocument pdfDocument = null;
            WordDocument wordDocument = null;
            try
            {

                using (wordDocument = new WordDocument(strmFile, FormatType.Docx))
                {
                    wordDocument.Comments.Clear();
                    converter = new DocToPDFConverter();
                    using (pdfDocument = converter.ConvertToPDF(strmFile))
                    {
                        wordDocument.Close();
                        pdfDocument.Save(strmFile);
                        pdfDocument.Close(true);
                    }
                }

            }
            catch (Exception ex)
            {
                File.AppendAllText(_szLogFileName, "Error : Convert_Document_To_PDF : " + ex.Message + EventLogEntryType.Error);
                msgError = ex.Message;
                strmPdfFile = null;

            }
            finally
            {
                converter = null;
                pdfDocument = null;
                wordDocument = null;
            }
            return strmPdfFile;
        }


        public Stream Convert_Document_To_PDF(File_Data oDestinationFile)
        {

            DocToPDFConverter converter = null;
            File.AppendAllText(_szLogFileName, "Convert_Document_To_PDF");
            PdfDocument pdfDocument = null;
            WordDocument wordDocument = null;
            Stream oStream = Convert_Document_To_Stream(oDestinationFile.Data);
            try
            {

                using (wordDocument = new WordDocument(oStream, FormatType.Docx))
                {
                    wordDocument.Comments.Clear();
                    converter = new DocToPDFConverter();
                    using (pdfDocument = converter.ConvertToPDF(oStream))
                    {
                        wordDocument.Close();
                        pdfDocument.Save(oStream);
                        pdfDocument.Close(true);
                    }
                }
            }
            catch (Exception ex)
            {
                File.AppendAllText(_szLogFileName, "Error : Convert_Document_To_PDF : " + ex.Message + EventLogEntryType.Error);
                msgError = ex.Message;
            }
            finally
            {
                converter = null;
            }
            return oStream;
        }

        public bool Update_Document_Properties(string szFilePath)
        {
            bool bResult = true;
            try
            {

                //using (WordDocument wordDocument = new WordDocument(szFilePath, FormatType.Docx))
                //{
                //    //wordDocument.Comments.Clear();
                //    wordDocument.UpdateDocumentFields();
                //    wordDocument.Save(szFilePath, FormatType.Docx);
                //    wordDocument.Close();
                //}
            }
            catch (Exception ex)
            {
                File.AppendAllText(_szLogFileName, "Error : Convert_Document_To_PDF : " + ex.Message + EventLogEntryType.Error);
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
            File.AppendAllText(_szLogFileName, "Process_Word_Document");
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
                File.AppendAllText(_szLogFileName, "Error : Convert_Document_To_PDF : " + ex.Message + EventLogEntryType.Error);
                bReturnVal = false;
                msgError = ex.Message;
            }
            finally { document = null; }
            return bReturnVal;
        }
        public Stream Process_Word_Document(Stream strmFile)
        {
            WordDocument document = null;
            File.AppendAllText(_szLogFileName, "Process_Word_Document");
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
                File.AppendAllText(_szLogFileName, "Error : Convert_Document_To_PDF : " + ex.Message + EventLogEntryType.Error);
                strmFile = null;
                msgError = ex.Message;
            }
            finally { document = null; }
            return strmFile;
        }

        private void Add_Log(string szMessage, EventLogEntryType e)
        {
            //using (EventLog eventLog = new EventLog("FileServer"))
            //{
            //    eventLog.Source = "FileServer";
            //    eventLog.WriteEntry(szMessage, e, 101, 1);
            //}
        }

        internal static Stream Convert_Document_To_Stream(byte[] arrDocument)
        {
            MemoryStream strmDocument = new MemoryStream();
            strmDocument.Write(arrDocument, 0, (int)arrDocument.Length);
            return strmDocument;
        }

        #endregion
    }
}
