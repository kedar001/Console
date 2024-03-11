using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.ComponentModel;
//using Acrobat = ACRODISTXLib;
using word = Microsoft.Office.Interop.Word;
//using excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

using iTextSharp;
using iTextSharp;
using iTextSharp.text.pdf;
using iTextSharp.text;
using iTextSharp.text.pdf.parser;
//using PdfSharp;
//using PdfSharp.Pdf;
//using PdfSharp.Pdf.IO;


namespace eDocsDN_Common_LockContent
{
    internal class clsWord
    {
        #region .... Variable Declaration ....

        //Acrobat.PdfDistillerClass objDistiller;
        word.Application objApp = null;
        word.Document objDoc = null;
        object objMissing = Type.Missing;

        bool _bResult = true;
        string _szError;
        string _szLockPwd;

        #endregion

        #region .... Property ....

        private string LockPassword
        {
            get { return _szLockPwd; }
            set { _szLockPwd = value; }
        }

        public string msgError
        {
            get { return _szError; }
            set { _szError = value; }
        }

        #endregion

        #region .... Constructor ....

        public clsWord()
        {
            this.handle = new IntPtr();
            msgError = "";
            _szLockPwd = Decrypt("<'432mqtF").ToString();
        }

        public clsWord(IntPtr iPtrHandle)
        {
            this.handle = iPtrHandle;
            LockPassword = "";
            msgError = "";
            _szLockPwd = Decrypt("<'432mqtF").ToString();
        }

        #endregion

        #region .... Internal And Private Class Functions ....

        internal string ConvertWordToHtml(Object objSourceWordFilePath, Object objTargetHtmlFilePath)
        {
            #region ... Word default objects ...

            msgError = "";
            objMissing = Type.Missing;
            Object targetFileFormat = word.WdSaveFormat.wdFormatHTML;
            Object objReadOnly = false;

            #endregion

            try
            {
                objApp = new word.Application();
                objDoc = new word.Document();

                //Word.Document objDocuments = new Word.Document();
                objDoc = objApp.Documents.Open(ref objSourceWordFilePath, ref objMissing, ref objReadOnly, ref objMissing, ref objMissing, ref objMissing,
                                                                     ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing,

                                                                     ref objMissing, ref objMissing, ref objMissing, ref objMissing);
                objDoc.SaveAs(ref objTargetHtmlFilePath, ref targetFileFormat, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing,
                                            ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing);

                objDoc.Close(ref objMissing, ref objMissing, ref objMissing);
                objApp.Quit(ref objMissing, ref objMissing, ref objMissing);
                objDoc = null;
                objApp = null;

                Add_Security_Code(objTargetHtmlFilePath.ToString());
            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = ex.Message;
            }
            finally
            {
                #region ... Clean objects ...

                if (objDoc != null)
                    objDoc.Close(ref objMissing, ref objMissing, ref objMissing);

                if (objApp != null)
                    objApp.Quit(ref objMissing, ref objMissing, ref objMissing);

                objDoc = null;
                objApp = null;

                objSourceWordFilePath = null;
                objReadOnly = null;

                #endregion

                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            return msgError;
        }

        //internal string ConvertWordToPdf(Object objSourceWordFilePath, string strOutputPDFFilePath)
        //{
        //    #region ... Word default objects ...

        //    msgError = "";
        //    string szSrc_WordFilePath = Convert.ToString(objSourceWordFilePath);

        //    Object objReadOnly = false;
        //    Object objRange = word.WdPrintOutRange.wdPrintAllDocument;
        //    Object objItem = word.WdPrintOutItem.wdPrintDocumentContent;
        //    Object objCopies = 1;
        //    Object objPages = "";
        //    Object objPageType = word.WdPrintOutPages.wdPrintAllPages;
        //    Object objCollate = false;
        //    Object objBackground = false;
        //    Object objPrintZoomRow = 0;
        //    Object objPrintZoomColumn = 0;
        //    Object objPrintZoomPaperWidth = 0;
        //    Object objPrintZoomPaperHeight = 0;
        //    Object objOutPutFileName = strOutputPDFFilePath.Trim().Replace(".pdf", ".ps");
        //    Object objAppend = false;

        //    objMissing = Type.Missing;

        //    #endregion

        //    try
        //    {
        //        //... Unlock files and convert it to .pdf files and lock files with new password ...
        //        objApp = new Microsoft.Office.Interop.Word.Application();
        //        objDoc = new Microsoft.Office.Interop.Word.Document();
        //        objDoc = objApp.Documents.Open(ref objSourceWordFilePath, ref objMissing, ref objReadOnly, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing);

        //        //... Take printout i.e convert word files to .ps i.e. Photoshop ...
        //        objDoc.PrintOut(ref objBackground, ref objAppend, ref objRange, ref objOutPutFileName, ref objMissing, ref objMissing,
        //                        ref objItem, ref objCopies, ref objPages, ref objPageType, ref objMissing, ref objCollate, ref objMissing,
        //                        ref objMissing, ref objPrintZoomColumn, ref objPrintZoomRow, ref objPrintZoomPaperWidth, ref objPrintZoomPaperHeight);

        //        objDoc.Close(ref objMissing, ref objMissing, ref objMissing);
        //        objApp.Quit(ref objMissing, ref objMissing, ref objMissing);
        //        objDoc = null;
        //        objApp = null;

        //        //... Convert .ps to .pdf ...
        //        string strInputFile = strOutputPDFFilePath.Replace(".pdf", ".ps");
        //        objDistiller = new Acrobat.PdfDistillerClass();
        //        string strJobOption = ""; //strJobOption = "Print";
        //        objDistiller.FileToPDF(strInputFile, strOutputPDFFilePath, strJobOption);
        //        strInputFile = null;
        //        objDistiller = null;
        //    }
        //    catch (Exception ex)
        //    {
        //        _bResult = false;
        //        msgError = ex.Message;
        //    }
        //    finally
        //    {
        //        #region ... Delete extra files (.ps and .log) ...

        //        DeleteFile(strOutputPDFFilePath.Replace(".pdf", ".ps"));
        //        DeleteFile(strOutputPDFFilePath.Replace(".pdf", ".log"));

        //        #endregion

        //        #region ... Clean objects ...

        //        if (objDoc != null)
        //            objDoc.Close(ref objMissing, ref objMissing, ref objMissing);

        //        if (objApp != null)
        //            objApp.Quit(ref objMissing, ref objMissing, ref objMissing);

        //        objDoc = null;
        //        objApp = null;

        //        objSourceWordFilePath = null;
        //        //objMissing = null;
        //        objReadOnly = null;
        //        objRange = null;
        //        objItem = null;
        //        objCopies = null;
        //        objPages = null;
        //        objPageType = null;
        //        objCollate = null;
        //        objBackground = null;
        //        objPrintZoomRow = null;
        //        objPrintZoomColumn = null;
        //        objPrintZoomPaperWidth = null;
        //        objPrintZoomPaperHeight = null;
        //        objOutPutFileName = null;
        //        objAppend = null;
        //        objDistiller = null;

        //        Kill_DistillerObject();

        //        #endregion

        //        GC.WaitForPendingFinalizers();
        //        GC.Collect();
        //    }
        //    return msgError;
        //}

        //internal bool Convert_WordToPDF(string szSrc_WordFilePath, string szDest_PDF_FilePath , bool bLock_SrcDocument)
        //{
        //    #region ... Word default objects ...

        //    msgError = "";
        //    object objSrcWordFile = szSrc_WordFilePath;

        //    Object objReadOnly = false;
        //    Object objRange = word.WdPrintOutRange.wdPrintAllDocument;
        //    Object objItem = word.WdPrintOutItem.wdPrintDocumentContent;
        //    Object objCopies = 1;
        //    Object objPages = "";
        //    Object objPageType = word.WdPrintOutPages.wdPrintAllPages;
        //    Object objCollate = false;
        //    Object objBackground = false;
        //    Object objPrintZoomRow = 0;
        //    Object objPrintZoomColumn = 0;
        //    Object objPrintZoomPaperWidth = 0;
        //    Object objPrintZoomPaperHeight = 0;
        //    Object objOutPutFileName = szDest_PDF_FilePath.Replace(".pdf", ".ps");
        //    Object objAppend = false;

        //    objMissing = Type.Missing;

        //    #endregion

        //    try
        //    {
        //        //... Unlock files and convert it to .pdf files and lock files with new password ...
        //        objApp = new Microsoft.Office.Interop.Word.Application();
        //        objDoc = new Microsoft.Office.Interop.Word.Document();
        //        //objDoc = objApp.Documents.Open(ref objSrcWordFile, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
        //        objDoc = objApp.Documents.Open(ref objSrcWordFile, ref objMissing, ref objReadOnly, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing);

        //        //... Unlock documents i.e .doc/.docx files ...
        //        if (objDoc.ProtectionType == word.WdProtectionType.wdAllowOnlyReading)
        //        {
        //            //... Get password of document files to unlock .doc/.docx files ...
        //            object szUnlock_Password = LockPassword;//Decrypt("<'432mqtF");
        //            objDoc.Unprotect(ref szUnlock_Password);
        //            szUnlock_Password = null;
        //        }

        //        objDoc.Save();
        //        objDoc.PrintFormsData = false;

        //        //... Take printout i.e convert word files to .ps i.e. Photoshop ...
        //        objDoc.PrintOut(ref objBackground, ref objAppend, ref objRange, ref objOutPutFileName, ref objMissing, ref objMissing,
        //                        ref objItem, ref objCopies, ref objPages, ref objPageType, ref objMissing, ref objCollate, ref objMissing,
        //                        ref objMissing, ref objPrintZoomColumn, ref objPrintZoomRow, ref objPrintZoomPaperWidth, ref objPrintZoomPaperHeight);

        //        //... Convert .ps to .pdf ...
        //        string strInputFile = szDest_PDF_FilePath.Replace(".pdf", ".ps");
        //        objDistiller = new Acrobat.PdfDistillerClass();
        //        string strJobOption = "";
        //        objDistiller.FileToPDF(strInputFile, szDest_PDF_FilePath, strJobOption);
        //        strInputFile = null;
        //        objDistiller = null;

        //        //... Lock documents i.e .doc/.docx files ...
        //        if (bLock_SrcDocument)
        //        {
        //            if (objDoc.ProtectionType == word.WdProtectionType.wdNoProtection)
        //            {
        //                //... Password to Lock document files ...
        //                object szLock_Password = LockPassword;
        //                object objNoReset = false;
        //                object objUseIRM = false;
        //                object objStyleLock = null;

        //                objDoc.Protect(word.WdProtectionType.wdAllowOnlyReading, ref objNoReset, ref  szLock_Password, ref objUseIRM, ref objStyleLock);

        //                szLock_Password = null;
        //                objNoReset = null;
        //                objUseIRM = null;
        //                objStyleLock = null;
        //            }
        //        }
        //        //...

        //        objDoc.Save();
        //        objDoc.Close(ref objMissing, ref objMissing, ref objMissing);
        //        objApp.Quit(ref objMissing, ref objMissing, ref objMissing);
        //        objDoc = null;
        //        objApp = null;

        //        objDistiller = null;
        //    }
        //    catch (Exception ex)
        //    {
        //        _bResult = false;
        //        msgError = ex.Message;
        //    }
        //    finally
        //    {
        //        #region ... Delete extra files (.ps and .log) ...

        //        DeleteFile(szDest_PDF_FilePath.Replace(".pdf", ".ps"));
        //        DeleteFile(szDest_PDF_FilePath.Replace(".pdf", ".log"));

        //        #endregion

        //        #region ... Clean objects ...

        //        if (objDoc != null)
        //            objDoc.Close(ref objMissing, ref objMissing, ref objMissing);

        //        if (objApp != null)
        //            objApp.Quit(ref objMissing, ref objMissing, ref objMissing);

        //        objDoc = null;
        //        objApp = null;

        //        objDistiller = null;
        //        objSrcWordFile = null;
        //        //objMissing = null;
        //        objReadOnly = null;
        //        objRange = null;
        //        objItem = null;
        //        objCopies = null;
        //        objPages = null;
        //        objPageType = null;
        //        objCollate = null;
        //        objBackground = null;
        //        objPrintZoomRow = null;
        //        objPrintZoomColumn = null;
        //        objPrintZoomPaperWidth = null;
        //        objPrintZoomPaperHeight = null;
        //        objOutPutFileName = null;
        //        objAppend = null;

        //        Kill_DistillerObject();

        //        #endregion

        //        GC.WaitForPendingFinalizers();
        //        GC.Collect();
        //    }
        //    return _bResult;
        //}

        internal string LockContentDoc(string szFullFilePath, string szLockType, object szDocPwd, string szDocStatus, bool bPrint)
        {
            #region ... Word default objects ...

            msgError = "";
            object objSrcWordFile = szFullFilePath;
            object objUnlockPassword = LockPassword;
            objMissing = Type.Missing;
            object reset = true;
            string szFieldName = "";

            #endregion

            try
            {
                objApp = new Microsoft.Office.Interop.Word.Application();
                objDoc = new Microsoft.Office.Interop.Word.Document();

                if (szDocPwd.ToString().Equals(""))
                {
                    objDoc = objApp.Documents.Open(ref objSrcWordFile, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing);
                }
                else
                {
                    objDoc = objApp.Documents.Open(ref objSrcWordFile, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref szDocPwd, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing);
                }

                if (szDocStatus != "")
                {
                    Object oCustom;
                    string szStatus = "Status";
                    oCustom = objDoc.CustomDocumentProperties;
                    Type typeDocCustomProps = oCustom.GetType();
                    typeDocCustomProps.InvokeMember("Item", System.Reflection.BindingFlags.Default | System.Reflection.BindingFlags.SetProperty, null, oCustom, new object[] { szStatus, szDocStatus });

                    for (int i = 1; i <= objDoc.Fields.Count; i++)
                    {

                        if (objDoc.Fields[i].Code.Text != null)
                        {
                            szFieldName = objDoc.Fields[i].Code.Text.Replace("DOCPROPERTY", "");
                            if (szFieldName.Contains("MERGEFORMAT"))
                            {
                                szFieldName = szFieldName.Replace("\\* MERGEFORMAT", "");
                                szFieldName = szFieldName.Replace(" ", "");

                                if (szFieldName.Trim().ToUpper() == "STATUS")
                                {
                                    objDoc.Fields[i].DoClick();
                                    objDoc.Fields[i].Update();
                                }
                            }
                        }
                    }
                }

                foreach (word.Range oStory in objDoc.StoryRanges)
                {
                    oStory.Fields.Update();
                }

                //For Each oStory In objDoc.StoryRanges
                //    oStory.Fields.Update()
                //    If oStory.StoryType <> word.WdStoryType.wdMainTextStory Then
                //        While Not (oStory.NextStoryRange Is Nothing)
                //            oStory = oStory.NextStoryRange
                //            oStory.Fields.Update()
                //        End While
                //    End If
                //Next oStory
                //oStory = Nothing

                if (objDoc.ProtectionType == word.WdProtectionType.wdNoProtection)
                {
                    if (szLockType.ToUpper() == "COMM")
                    {
                        objDoc.Protect(word.WdProtectionType.wdAllowOnlyComments, ref reset, ref  objUnlockPassword, ref objMissing, ref objMissing);
                    }
                    else if (szLockType.ToUpper() == "FORM")
                    {
                        objDoc.Protect(word.WdProtectionType.wdAllowOnlyFormFields, ref reset, ref objUnlockPassword, ref objMissing, ref objMissing);
                    }
                    else
                    {
                        objDoc.Protect(word.WdProtectionType.wdAllowOnlyReading, ref reset, ref objUnlockPassword, ref objMissing, ref objMissing);
                    }

                } //if (oApp.ActiveDocument.ProtectionType == Word.WdProtectionType.wdNoProtection)

                if (bPrint == false)
                {
                    objDoc.PrintFormsData = true;
                }
                else
                {
                    objDoc.PrintFormsData = false;
                }

                objDoc.Save();
                objDoc.Close(ref objMissing, ref objMissing, ref objMissing);
                objApp.Quit(ref objMissing, ref objMissing, ref objMissing);
                objDoc = null;
                objApp = null;
                // }
            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = ex.Message;
            }
            finally
            {
                #region ... Clean objects ...

                objSrcWordFile = null;

                if (objDoc != null)
                    objDoc.Close(ref objMissing, ref objMissing, ref objMissing);

                if (objApp != null)
                    objApp.Quit(ref objMissing, ref objMissing, ref objMissing);

                //objMissing = null;
                objDoc = null;
                objApp = null;

                #endregion

                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            return msgError;
        }


        internal bool UnLock_WordDocument(string szSrc_WordFilePath, object objUnlockPassword)
        {
            #region ... Word default objects ...

            msgError = "";
            object objSrcWordFile = szSrc_WordFilePath;
            objMissing = Type.Missing;

            #endregion

            try
            {
                //... Unlock files and convert it to .pdf files and lock files with new password ...
                objApp = new Microsoft.Office.Interop.Word.Application();
                objDoc = new Microsoft.Office.Interop.Word.Document();

                if (objUnlockPassword.ToString().Equals(""))
                {
                    objDoc = objApp.Documents.Open(ref objSrcWordFile, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing);
                }
                else
                {
                    objDoc = objApp.Documents.Open(ref objSrcWordFile, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objUnlockPassword, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing);
                }

                //... Unlock documents i.e .doc/.docx files ...
                if (objDoc.ProtectionType == word.WdProtectionType.wdAllowOnlyComments || objDoc.ProtectionType == word.WdProtectionType.wdAllowOnlyFormFields || objDoc.ProtectionType == word.WdProtectionType.wdAllowOnlyReading)
                {
                    //... Get password of document files to unlock .doc/.docx files ...
                    objUnlockPassword = LockPassword;
                    objDoc.Unprotect(ref objUnlockPassword);
                    objUnlockPassword = null;
                }

                objDoc.PrintFormsData = false;
                objDoc.Save();
                objDoc.Close(ref objMissing, ref objMissing, ref objMissing);
                objApp.Quit(ref objMissing, ref objMissing, ref objMissing);
                objDoc = null;
                objApp = null;
            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = ex.Message;
            }
            finally
            {
                #region ... Clean objects ...

                objSrcWordFile = null;

                if (objDoc != null)
                    objDoc.Close(ref objMissing, ref objMissing, ref objMissing);

                if (objApp != null)
                    objApp.Quit(ref objMissing, ref objMissing, ref objMissing);

                //objMissing = null;
                objDoc = null;
                objApp = null;

                #endregion

                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            return _bResult;
        }

        internal string UnLockContentDoc(string szFullFilePath, object szDocPwd)
        {
            UnLock_WordDocument(szFullFilePath, szDocPwd);
            return msgError;
        }

        /// <summary>
        /// Developer Name :- Harshad Chavan.
        /// DRT:-4251
        /// </summary>
        /// <param name="szSourceFile"></param>
        /// <param name="szDestFilePath"></param>
        public void ConvertWordFileToPdf(string szSourceFile, string szDestFilePath, bool bLockFile)
        {
            word._Application objApp = null;
            word._Document objDoc = null;
            object objMissing = Type.Missing;
            object reset = false;
            string szTempFilePath, szFileName;

            try
            {
                #region .... Try Region ....

                if (bLockFile)
                {
                    szFileName = szSourceFile.Substring(szSourceFile.LastIndexOf('\\') + 1);
                    szTempFilePath = szSourceFile.Substring(0, szSourceFile.LastIndexOf('\\'));
                    szTempFilePath += "\\TempPDFLockingFolderH";
                    if (!Directory.Exists(szTempFilePath))
                        Directory.CreateDirectory(szTempFilePath);
                    szTempFilePath += "\\" + szFileName.Substring(0, szFileName.LastIndexOf('.')) + ".pdf";
                }
                else
                    szTempFilePath = szDestFilePath;

                objApp = new Microsoft.Office.Interop.Word.Application();
                objDoc = new Microsoft.Office.Interop.Word.Document();
                object missing = Type.Missing;
                object objFileName = szSourceFile;
                objDoc = objApp.Documents.Open(ref objFileName, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                //..... Convert to PDF ....
                //objDoc.ExportAsFixedFormat(szPDFFileName, word.WdExportFormat.wdExportFormatPDF, true, word.WdExportOptimizeFor.wdExportOptimizeForPrint, word.WdExportRange.wdExportAllDocument, 1, 1, word.WdExportItem.wdExportDocumentContent, true, true, word.WdExportCreateBookmarks.wdExportCreateNoBookmarks, true, true, false);

                objDoc.ExportAsFixedFormat(szTempFilePath, word.WdExportFormat.wdExportFormatPDF, false, word.WdExportOptimizeFor.wdExportOptimizeForOnScreen, word.WdExportRange.wdExportAllDocument, 1, 1, word.WdExportItem.wdExportDocumentContent, false, false, word.WdExportCreateBookmarks.wdExportCreateHeadingBookmarks, true, false, false);

                objDoc.Close(true, objMissing, objMissing);
                objApp.Quit(true, objMissing, objMissing);
                objApp = null;
                objDoc = null;

                if (bLockFile)
                {
                    LockPDFFile(szTempFilePath, szDestFilePath);
                    File.Delete(szTempFilePath);
                }

                #endregion
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            finally
            {
                objDoc = null;
                if (objApp != null)
                    objApp.Quit();
                objApp = null;
            }
        }

        public void LockPDFFile(string szSourceFile, string szDestFilePath)
        {
            PdfReader PFreader = null;
            MemoryStream ms = null;
            FileStream fs = null;

            try
            {
                #region .... Try Block ....


                //PdfSharp.Pdf.PdfDocument maindoc = PdfSharp.Pdf.IO.PdfReader.Open(szSourceFile, PdfDocumentOpenMode.Import);
                ////'Create the Output Document
                //PdfSharp.Pdf.PdfDocument OutputDoc = new PdfSharp.Pdf.PdfDocument();

                ////'Copy over pages from original document
                //foreach (PdfSharp.Pdf.PdfPage page in maindoc.Pages)
                //{
                //    OutputDoc.AddPage(page);
                //}
                //OutputDoc.Save(szDestFilePath);
                //maindoc.Dispose();
                //OutputDoc.Dispose();



                PFreader = new PdfReader(szSourceFile);
                ms = new MemoryStream();
                using (PdfStamper stamper = new PdfStamper(PFreader, ms))
                {
                    // add your content
                }
                fs = new FileStream(szDestFilePath, FileMode.Create, FileAccess.ReadWrite);
                PdfEncryptor.Encrypt(new PdfReader(ms.ToArray()), fs, null, System.Text.Encoding.UTF8.GetBytes("Espl123&"), PdfWriter.ALLOW_SCREENREADERS, true);
                //PdfEncryptor.Encrypt(new PdfReader(ms.ToArray()), fs, null, null, PdfWriter.ALLOW_MODIFY_CONTENTS, true);


                //fs.Flush();
                //fs.Close();
                fs = null;

                //ms.Flush();
                //ms.Close();
                ms = null;

                PFreader.Close();
                PFreader.Dispose();
                PFreader = null;

                #endregion
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            finally
            {
                fs = null;
                ms = null;
                if (PFreader != null)
                {
                    PFreader.Close();
                    PFreader.Dispose();
                    PFreader = null;
                }
            }
        }

        public void UpdateStatus(string szFullFilePath, object szDocPwd, string szDocStatus)
        {
            msgError = "";
            object objSrcWordFile = szFullFilePath;
            object objUnlockPassword = LockPassword;
            objMissing = Type.Missing;
            object reset = true;
            string szFieldName = "";

            try
            {
                objApp = new Microsoft.Office.Interop.Word.Application();
                objDoc = new Microsoft.Office.Interop.Word.Document();

                if (szDocPwd.ToString().Equals(""))
                {
                    objDoc = objApp.Documents.Open(ref objSrcWordFile, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing);
                }
                else
                {
                    objDoc = objApp.Documents.Open(ref objSrcWordFile, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref szDocPwd, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing);
                }

                if (szDocStatus != "")
                {
                    Object oCustom;
                    string szStatus = "Status";
                    oCustom = objDoc.CustomDocumentProperties;
                    Type typeDocCustomProps = oCustom.GetType();
                    typeDocCustomProps.InvokeMember("Item", System.Reflection.BindingFlags.Default | System.Reflection.BindingFlags.SetProperty, null, oCustom, new object[] { szStatus, szDocStatus });

                    for (int i = 1; i <= objDoc.Fields.Count; i++)
                    {

                        if (objDoc.Fields[i].Code.Text != null)
                        {
                            szFieldName = objDoc.Fields[i].Code.Text.Replace("DOCPROPERTY", "");
                            if (szFieldName.Contains("MERGEFORMAT"))
                            {
                                szFieldName = szFieldName.Replace("\\* MERGEFORMAT", "");
                                szFieldName = szFieldName.Replace(" ", "");

                                if (szFieldName.Trim().ToUpper() == "STATUS")
                                {
                                    objDoc.Fields[i].DoClick();
                                    objDoc.Fields[i].Update();
                                }
                            }
                        }
                    }
                }
                foreach (word.Range oStory in objDoc.StoryRanges)
                {
                    oStory.Fields.Update();
                }
                objDoc.Save();
                objDoc.Close(ref objMissing, ref objMissing, ref objMissing);
                objApp.Quit(ref objMissing, ref objMissing, ref objMissing);
                objDoc = null;
                objApp = null;
            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = ex.Message;
            }
            finally
            {
                #region ... Clean objects ...

                objSrcWordFile = null;

                if (objDoc != null)
                    objDoc.Close(ref objMissing, ref objMissing, ref objMissing);

                if (objApp != null)
                    objApp.Quit(ref objMissing, ref objMissing, ref objMissing);

                //objMissing = null;
                objDoc = null;
                objApp = null;

                #endregion

                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }
        #endregion

        #region .... Functions Definition ....

        private string Add_Security_Code(string szFullFilePath)
        {
            StreamWriter objSW = new StreamWriter(szFullFilePath, true);
            string szLineToAppend = "<script language=\"javascript1.2\" src=\"disable.js\"></script>";
            try
            {
                objSW.WriteLine("");
                objSW.WriteLine(szLineToAppend);
                objSW.Flush();
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            finally
            {
                objSW.Close();
                objSW.Dispose();
                objSW = null;
            }
            return msgError;
        }

        public string Encrypt(string szString)
        {
            char szChar;
            int ilen = szString.Length;
            string szEncString = "";
            for (int i = 0; i < ilen; i++)
            {
                szChar = szString[i];
                szEncString = szEncString + Convert.ToChar(Convert.ToInt32(szChar) + 1);
            }
            return Reverse(szEncString);
        }

        public string Decrypt(string szString)
        {
            char szChar;
            int ilen = szString.Length;
            string szDecString = "";
            for (int i = 0; i < ilen; i++)
            {
                szChar = szString[i];
                szDecString = szDecString + Convert.ToChar(Convert.ToInt32(szChar) - 1);
            }
            return Reverse(szDecString);
        }

        private string Reverse(string szString)
        {
            string szRevString = "";
            int ilen = szString.Length;

            for (int i = ilen - 1; i >= 0; i--)
            {
                szRevString = szRevString + szString[i];
            }
            return szRevString;
        }

        private bool DeleteFile(string szFilePath)
        {
            bool bResult = true;
            try
            {
                if (File.Exists(szFilePath))
                    File.Delete(szFilePath);
            }
            catch (Exception ex)
            {
                bResult = false;
                msgError = ex.Message;
            }
            return bResult;
        }

        private void Kill_DistillerObject()
        {
            System.Diagnostics.Process[] pro = null;
            try
            {
                pro = System.Diagnostics.Process.GetProcesses();
                foreach (System.Diagnostics.Process pr in pro)
                {
                    if (pr.ProcessName.Equals("acrodist"))
                    {
                        pr.Kill();
                        pr.Refresh();
                        break;
                    }
                }
            }
            catch (Exception) { }
            finally
            {
                pro = null;
            }
        }

        #endregion

        #region .... Functions for IDisposable Interface ....

        #region ... Variable Declaration for Disposable Object ...

        private IntPtr handle;
        private Component CompConversion = new Component();
        private bool bDisposed = false;

        #endregion

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        [System.Runtime.InteropServices.DllImport("Kernel32")]
        private extern static Boolean CloseHandle(IntPtr handle);

        ~clsWord()
        {
            Dispose(false);
        }

        private void Dispose(bool bDisposing)
        {
            if (!this.bDisposed)
            {
                if (bDisposing)
                {
                    if (objDoc != null)
                        objDoc.Close(ref objMissing, ref objMissing, ref objMissing);

                    if (objApp != null)
                        objApp.Quit(ref objMissing, ref objMissing, ref objMissing);

                    if (CompConversion != null)
                        CompConversion.Dispose();
                    CompConversion = null;
                }
                objMissing = null;
                objDoc = null;
                objApp = null;
                //objDistiller = null;
                msgError = null;
                LockPassword = null;

                CloseHandle(handle);
                handle = IntPtr.Zero;
                bDisposed = true;
            }
        }

        #endregion
    }
}