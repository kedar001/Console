#region //
#endregion

using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Acrobat = ACRODISTXLib;
using System.IO;
using System.Windows.Forms;

 namespace eDocsDN_Common_LockContent
{
    internal class clsExcel
    {

        #region ...Class Variables...

        Excel.Application objExcel;
        Excel.Workbooks objExcelWorkBooks;
        Excel._Workbook objCurrentWorkBook;
        Acrobat.PdfDistillerClass objDistiller;
        Object objMissing = Type.Missing;
        Object objReadOnly = false;
        Process[] mobjProcess;

        private int mintExcelPID;
        //private int mintAcrobatPID;

        string szErrorMsg="";
        object szPwd;
        private bool mblnDistiller = false;

        #endregion

        #region ...Class Constructor...

        public clsExcel()
        {
            szPwd = Decrypt("<'432mqtF").ToString();
        }

        #endregion
        
        #region...Functions To Convert Xls To Htm...

        internal string Convert_Excel_To_Html(string szSourceExcelFilePath, string szTargetHtmlFilePath)
        {

            string szFolder;
            int iSheetCount = 0;

            objExcel = new Excel.Application();

            try
            {
                Object objtargetFileFormat = Excel.XlFileFormat.xlHtml;
                Object objMissing = Type.Missing;
                Object objReadOnly = false;
                Object objCreateBackUp = false;

                objExcelWorkBooks = objExcel.Workbooks;

                objCurrentWorkBook = objExcelWorkBooks.Open(szSourceExcelFilePath, objMissing, objReadOnly, objMissing, objMissing, objMissing, objMissing, objMissing,
                                           objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing);

                iSheetCount = objCurrentWorkBook.Sheets.Count;

                objCurrentWorkBook.Application.ActiveWorkbook.SaveAs(szTargetHtmlFilePath , objtargetFileFormat, objMissing, objMissing, objMissing, objMissing, Excel.XlSaveAsAccessMode.xlNoChange, objMissing,
                                                                    objMissing, objMissing, objMissing, objMissing);
                Object objSrcFile = (Object)szTargetHtmlFilePath;  //szPublishHtm + szHtmFileName;

                objCurrentWorkBook.Close(objMissing, objSrcFile, objMissing);

                objExcel.Quit();

                //if sheet count > 1 then append security code to each sheet
                if (iSheetCount > 1)
                {
                    szFolder = szTargetHtmlFilePath.Replace(".htm", "_files\\");
                    if (System.IO.Directory.Exists(szFolder))
                    {
                        string[] szfiles = System.IO.Directory.GetFiles(szFolder, "sheet*.htm");
                        foreach (string objfile in szfiles)
                        {
                           szErrorMsg = Add_Security_Code(objfile, true);
                        }
                    }
                }
                else
                {
                    szErrorMsg = Add_Security_Code(szTargetHtmlFilePath , false);
                }
            }
            catch (Exception ex)
            {
                szErrorMsg = ex.Message;
            }
            finally
            {
                mobjProcess = Process.GetProcessesByName("EXCEL");
                if (!mobjProcess[0].HasExited)
                {
                    mobjProcess[0].Kill();
                }
            }
            return szErrorMsg;
        }

        internal string Convert_Excel_To_Pdf(string szExcelFullFilePath, string szPdfFFullFilePath)
        {
            objExcel = new Excel.Application();

            try
            {

                Object objMissing = Type.Missing;
                Object objReadOnly = false;
                Object objCreateBackUp = false;
                Object objCopies = 1;
                Object objCollate = false;
                Object objPreView = false;
                Object objPrintToFile = true;

                objExcelWorkBooks = objExcel.Workbooks;

                objCurrentWorkBook = objExcelWorkBooks.Open(szExcelFullFilePath, objMissing, objReadOnly, objMissing, objMissing, objMissing, objMissing,
                                                          objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing);

                mobjProcess = Process.GetProcessesByName("EXCEL");
                mintExcelPID = mobjProcess[0].Id;

                Object objTargetFilePath = (Object)szPdfFFullFilePath.Replace(".pdf", ".ps");
                objCurrentWorkBook.Application.ActiveWorkbook.PrintOut(objMissing, objMissing, objCopies, objPreView, objMissing, objPrintToFile, objCollate, objTargetFilePath);

                Object objSrcFile = (Object)szExcelFullFilePath;

                objCurrentWorkBook.Close(objMissing, objSrcFile, objMissing);

                objExcel.Quit();

                if (mblnDistiller == false)
                {
                    objDistiller = new Acrobat.PdfDistillerClass();
                    mblnDistiller = true;
                }

                string strJobOption = "Print";
                objDistiller.FileToPDF(szPdfFFullFilePath.Replace(".pdf", ".ps"), szPdfFFullFilePath, strJobOption);

                System.Diagnostics.Process[] pro = System.Diagnostics.Process.GetProcesses();
                foreach (System.Diagnostics.Process pr in pro)
                {
                    if (pr.ProcessName == "acrodist" || pr.ProcessName == "EXCEL")
                    {
                        pr.Kill();
                        pr.Refresh();
                    }
                }
            }
            catch (Exception ex)
            {
                szErrorMsg = ex.Message;
            }
            finally
            {
                mobjProcess = Process.GetProcessesByName("EXCEL");
                if (!mobjProcess[0].HasExited)
                {
                    mobjProcess[0].Kill();
                }
            }
            return szErrorMsg;
        }

        private string Add_Security_Code(string szFullFilePath, bool bFolderExist)
        {
            StreamWriter objSW = new StreamWriter(szFullFilePath, true);
            string szLineToAppend = "";
            if (bFolderExist == false)
            {
                szLineToAppend = "<script language=\"javascript1.2\" src=\"disable.js\"></script>";
            }
            else
            {
                szLineToAppend = "<script language=\"javascript1.2\" src=\"..\\disable.js\"></script>";
            }

            try
            {
                objSW.WriteLine("");
                objSW.WriteLine(szLineToAppend);
                objSW.Flush();
            }
            catch (Exception ex)
            {
                szErrorMsg = ex.Message;
            }
            finally
            {
                objSW.Close();
                objSW = null;
            }
            return szErrorMsg;
        }//private string Add_Security_Code(string szFullFilePath, bool bFolderExist)

        #endregion

        #region...Functions To Lock Xls File...

        internal string LockExcelContent(string szFullFilePath, string szLockType, string szExcelPwd, string szDocStatus, bool bPrint)
        {
            objExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Worksheet objWorkSheet = new Microsoft.Office.Interop.Excel.Worksheet();

            Object objMissing = Type.Missing;
            Object objDrawing = false;
            Object objContents = true;
            try
            {
                if (szExcelPwd.Equals(""))
                {
                    objExcel.Workbooks.Open(szFullFilePath, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing);
                }
                else
                {
                    objExcel.Workbooks.Open(szFullFilePath, objMissing, objMissing, objMissing, objMissing, szExcelPwd.ToString(), objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing);
                }

                if (szLockType.ToUpper() == "COMM")
                {
                    objDrawing = false;
                    objContents = true;
                    for (int i = 1; i <= objExcel.ActiveWorkbook.Sheets.Count; i++)
                    {
                        objWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)objExcel.ActiveWorkbook.Sheets[i];
                        objWorkSheet.Protect(szPwd, objDrawing, objContents, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing);
                        if (bPrint == false)
                        {
                            objWorkSheet.PageSetup.PrintArea = "$A$65536:$B$65536";
                        }
                        else
                        {
                            objWorkSheet.PageSetup.PrintArea = "";
                        }
                    }
                }
                else
                {
                    objDrawing = true;
                    objContents = true;
                    for (int i = 1; i <= objExcel.ActiveWorkbook.Sheets.Count; i++)
                    {
                        objWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)objExcel.ActiveWorkbook.Sheets[i];
                        objWorkSheet.Protect(szPwd, objDrawing, objContents, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing);
                        if (bPrint == false)
                        {
                            objWorkSheet.PageSetup.PrintArea = "$A$65536:$B$65536";
                        }
                        else
                        {
                            objWorkSheet.PageSetup.PrintArea = "";
                        }
                    }
                }
                objExcel.ActiveWorkbook.Save();
            }//try
            catch (Exception ex)
            {
                szErrorMsg = ex.Message;
            }
            finally
            {
                mobjProcess = Process.GetProcessesByName("EXCEL");
                if (!mobjProcess[0].HasExited)
                {
                    mobjProcess[0].Kill();
                }
            }//finally
            return szErrorMsg;
        }

        internal string UnLockContentXls(string szFullFilePath, object szExcelPwd)
        {
            objExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Worksheet objWorkSheet = new Microsoft.Office.Interop.Excel.Worksheet();

            Object objMissing = Type.Missing;
            Object objDrawing = false;
            Object objContents = true;
            try
            {
                if (szExcelPwd.Equals(""))
                {
                    objExcel.Workbooks.Open(szFullFilePath, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing);
                }
                else
                {
                    objExcel.Workbooks.Open(szFullFilePath, objMissing, objMissing, objMissing, objMissing, szExcelPwd.ToString(), objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing);
                }

                objDrawing = false;
                objContents = true;
                for (int i = 1; i <= objExcel.ActiveWorkbook.Sheets.Count; i++)
                {
                    objWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)objExcel.ActiveWorkbook.Sheets[i];
                    objWorkSheet.Unprotect(szPwd);
                    objWorkSheet.PageSetup.PrintArea = "";
                }
                objExcel.ActiveWorkbook.Save();
            }//try
            catch (Exception ex)
            {
                szErrorMsg = ex.Message;
            }//catch
            finally
            {
                mobjProcess = Process.GetProcessesByName("EXCEL");
                if (!mobjProcess[0].HasExited)
                {
                    mobjProcess[0].Kill();
                }
            }//finally
            return szErrorMsg;
        }//internal string UnLockContentXls(string szFullFilePath, object szExcelPwd)

        private string Encrypt(string szString)
        {
            int i, ilen;
            string szTemp;
            char szChar;
            ilen = szString.Length;
            szTemp = "";
            for (i = 0; i < ilen; i++)
            {
                szChar = Convert.ToChar(szString.Substring(i, 1));
                szTemp = szTemp + Convert.ToChar(Convert.ToInt32(szChar) + 1);
            }
            szTemp = Reverse(szTemp);
            return szTemp;
        }

        private string Decrypt(string szString)
        {
            int i, ilen;
            string szTemp;
            char szChar;
            ilen = szString.Length;
            szTemp = "";
            for (i = 0; i < ilen; i++)
            {
                szChar = Convert.ToChar(szString.Substring(i, 1));
                szTemp = szTemp + Convert.ToChar(Convert.ToInt32(szChar) - 1);
            }
            return Reverse(szTemp);
        }

        private string Reverse(string str)
        {
            int i, ilen;
            string szTemp;
            char ch;

            szTemp = "";
            ilen = str.Length;
            for (i = ilen - 1; i >= 0; i--)
            {
                ch = str[i];
                szTemp = szTemp + ch;
            }
            return szTemp;
        }

        #endregion

    }//public class clsExcel
}//namespace eDocsDN_Common_LockContent

