using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using word = Microsoft.Office.Interop.Word;
using eDocsDN_Common_LockContent;
using eDocsDN_ReadAppXml;
using System.ServiceModel;



namespace eDocsDN_Word_File_Operations
{
    public class ClsWord_File_Operations:IDisposable
    {
        #region ..... Variable Declaration .....
        clsReadAppXml _objINI = null;
        string _szFilePath = string.Empty;
        string _szAppXmlPath = string.Empty;
        //string _szLogFileName = string.Empty;
        #endregion

        #region ..... Properties .....
        public string msgError { get; set; }
        #endregion

        #region ..... Constroctor .....
        //public ClsWord_File_Operations(string szFilePath)
        //{
        //    msgError = string.Empty;
        //    _szFilePath = szFilePath;
        //}
        public ClsWord_File_Operations(string szAppXmlPath)
        {
            msgError = string.Empty;
            _szAppXmlPath = szAppXmlPath;
            //_szLogFileName = AppDomain.CurrentDomain.BaseDirectory + @"\Log\FileOperationLog.txt";
            //if (!Directory.Exists(AppDomain.CurrentDomain.BaseDirectory + @"\Log\"))
            //    Directory.CreateDirectory(AppDomain.CurrentDomain.BaseDirectory + @"\Log\");

        }
        #endregion

        #region .....Public Functions .....

        //   public bool Process_Word_Document()
        //   {
        //       bool bResult = true;
        //       Microsoft.Office.Interop.Word.Application wordApp = null;
        //       Microsoft.Office.Interop.Word.Document objDoc = null;
        //       Object objMiss = Type.Missing, objSave = true, objDocFile;
        //       Object oCustom;
        //       Microsoft.Office.Interop.Word.HeaderFooter headerg;
        //       Microsoft.Office.Interop.Word.Range rangeg;
        //       try
        //       {

        //           #region .... initialize Variables ...
        //           objDocFile = _szFilePath;
        //           wordApp = new Microsoft.Office.Interop.Word.Application();
        //           objDoc = new Microsoft.Office.Interop.Word.Document();

        //           #endregion

        //           #region ... Performance ...

        //           wordApp.ScreenUpdating = false;
        //           wordApp.Visible = false;
        //           wordApp.CheckLanguage = false;
        //           wordApp.Options.PrintBackground = false;
        //           wordApp.Options.AutoFormatAsYouTypeApplyHeadings = false;
        //           wordApp.Options.AutoFormatAsYouTypeApplyBorders = false;
        //           wordApp.Options.AutoFormatAsYouTypeApplyBulletedLists = false;
        //           wordApp.Options.AutoFormatAsYouTypeApplyNumberedLists = false;
        //           wordApp.Options.AutoFormatAsYouTypeApplyTables = false;
        //           wordApp.Options.AutoFormatAsYouTypeReplacePlainTextEmphasis = false;
        //           wordApp.Options.AutoFormatAsYouTypeDefineStyles = false;
        //           wordApp.Options.LabelSmartTags = false;
        //           wordApp.Options.AnimateScreenMovements = false;
        //           wordApp.Options.CheckGrammarAsYouType = false;
        //           wordApp.Options.CtrlClickHyperlinkToOpen = false;
        //           wordApp.AutoCorrect.CorrectKeyboardSetting = false;
        //           wordApp.OMathAutoCorrect.UseOutsideOMath = false;

        //           #endregion

        //           #region ..... Update Custom Properties ...

        //           objDoc = wordApp.Documents.Open(ref objDocFile, ref objMiss, ref objMiss, ref objMiss, ref objMiss, ref objMiss, ref objMiss, ref objMiss, ref objMiss, ref objMiss, ref objMiss, ref objMiss, ref objMiss, ref objMiss, ref objMiss, ref objMiss);

        //           #region .... Set Track Changes Off ....
        //           objDoc.TrackRevisions = false;
        //           #endregion

        //           oCustom = objDoc.CustomDocumentProperties;
        //           for (int i = 1; i <= objDoc.Fields.Count; i++)
        //           {
        //               if (objDoc.Fields[i].Code.Text != null)
        //               {
        //                   objDoc.Fields[i].DoClick();
        //                   objDoc.Fields[i].Update();
        //               }
        //           }

        //           headerg = objDoc.Sections[1].Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage];
        //           rangeg = headerg.Range;
        //           rangeg.Fields.Update();
        //           headerg = objDoc.Sections[1].Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
        //           rangeg = headerg.Range;
        //           rangeg.Fields.Update();


        //           #endregion

        //           #region ....... Accept  Revision .....
        //           if (objDoc.Comments.Count > 0)
        //           {
        //               objDoc.DeleteAllComments();
        //           }


        //           objDoc.AcceptAllRevisions();
        //           objDoc.PrintPreview();
        //           objDoc.ClosePrintPreview();
        //           objDoc.PrintFormsData = true;
        //           objDoc.Save();

        //           #endregion

        //           #region ..... Close Object .....
        //           if (objDoc != null)
        //               objDoc.Close(ref objSave, ref objMiss, ref objMiss);
        //           objDoc = null;

        //           if (wordApp != null)
        //               wordApp.Quit(ref objSave, ref objMiss, ref objMiss);
        //           wordApp = null;

        //           #endregion
        //       }
        //       catch (Exception ex)
        //       {
        //           bResult = false;
        //           string message =
        //   "Exception type " + ex.GetType() + Environment.NewLine +
        //   "Exception message: " + ex.Message + Environment.NewLine +
        //   "Stack trace: " + ex.StackTrace + Environment.NewLine;
        //           if (ex.InnerException != null)
        //           {
        //               message += "---BEGIN InnerException--- " + Environment.NewLine +
        //                          "Exception type " + ex.InnerException.GetType() + Environment.NewLine +
        //                          "Exception message: " + ex.InnerException.Message + Environment.NewLine +
        //                          "Stack trace: " + ex.InnerException.StackTrace + Environment.NewLine +
        //                          "---END Inner Exception";
        //           }
        //           throw new Exception("Something went wrong. please generate Preview again");


        //       }
        //       finally
        //       {
        //           #region ..... Close Documents ...
        //           objDocFile = null;
        //           oCustom = null;
        //           objMiss = null; objSave = null;
        //           headerg = null;
        //           rangeg = null;

        //           if (objDoc != null)
        //               objDoc.Close(ref objSave, ref objMiss, ref objMiss);
        //           objDoc = null;

        //           if (wordApp != null)
        //               wordApp.Quit(ref objSave, ref objMiss, ref objMiss);
        //           wordApp = null;
        //           GC.Collect();
        //           #endregion
        //       }
        //       return bResult;
        //   }

        //   public bool Update_Document_Properties()
        //   {
        //       bool bResult = true;
        //       Microsoft.Office.Interop.Word.Application wordApp = null;
        //       Microsoft.Office.Interop.Word.Document objDoc = null;
        //       Object objMiss = Type.Missing, objSave = true, objDocFile = null;
        //       Object oCustom = null;
        //       Microsoft.Office.Interop.Word.HeaderFooter headerg = null;
        //       Microsoft.Office.Interop.Word.Range rangeg = null;
        //       try
        //       {
        //           objDocFile = _szFilePath;
        //           wordApp = new Microsoft.Office.Interop.Word.Application();

        //           #region ... Performance ...

        //           wordApp.ScreenUpdating = false;
        //           wordApp.Visible = false;
        //           wordApp.CheckLanguage = false;
        //           wordApp.Options.PrintBackground = false;
        //           wordApp.Options.AutoFormatAsYouTypeApplyHeadings = false;
        //           wordApp.Options.AutoFormatAsYouTypeApplyBorders = false;
        //           wordApp.Options.AutoFormatAsYouTypeApplyBulletedLists = false;
        //           wordApp.Options.AutoFormatAsYouTypeApplyNumberedLists = false;
        //           wordApp.Options.AutoFormatAsYouTypeApplyTables = false;
        //           wordApp.Options.AutoFormatAsYouTypeReplacePlainTextEmphasis = false;
        //           wordApp.Options.AutoFormatAsYouTypeDefineStyles = false;
        //           wordApp.Options.LabelSmartTags = false;
        //           wordApp.Options.AnimateScreenMovements = false;
        //           wordApp.Options.CheckGrammarAsYouType = false;
        //           wordApp.Options.CtrlClickHyperlinkToOpen = false;
        //           wordApp.AutoCorrect.CorrectKeyboardSetting = false;
        //           wordApp.OMathAutoCorrect.UseOutsideOMath = false;

        //           #endregion

        //           #region ..... Update Custom Properties ...
        //           objDoc = wordApp.Documents.Open(ref objDocFile, ref objMiss, ref objMiss, ref objMiss, ref objMiss, ref objMiss, ref objMiss, ref objMiss, ref objMiss, ref objMiss, ref objMiss, ref objMiss, ref objMiss, ref objMiss, ref objMiss, ref objMiss);
        //           if (objDoc == null)
        //               System.Threading.Thread.Sleep(10000);
        //           objDoc.TrackRevisions = false;
        //           oCustom = objDoc.CustomDocumentProperties;
        //           if (oCustom != null)
        //               for (int i = 1; i <= objDoc.Fields.Count; i++)
        //               {
        //                   try
        //                   {
        //                       if (objDoc.Fields[i].Code.Text != null)
        //                       {
        //                           objDoc.Fields[i].DoClick();
        //                           objDoc.Fields[i].Update();
        //                       }
        //                   }
        //                   catch (Exception ex)
        //                   { }
        //               }
        //           headerg = objDoc.Sections[1].Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage];
        //           rangeg = headerg.Range;
        //           rangeg.Fields.Update();
        //           headerg = objDoc.Sections[1].Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
        //           rangeg = headerg.Range;
        //           rangeg.Fields.Update();

        //           objDoc.PrintPreview();
        //           objDoc.ClosePrintPreview();

        //           #endregion

        //           #region ..... Close Object .....

        //           if (objDoc != null)
        //               objDoc.Close(ref objSave, ref objMiss, ref objMiss);
        //           objDoc = null;

        //           if (wordApp != null)
        //               wordApp.Quit(ref objSave, ref objMiss, ref objMiss);
        //           wordApp = null;
        //           #endregion

        //       }
        //       catch (Exception ex)
        //       {
        //           bResult = false;
        //           ////WriteLog("Error : " + ex.Message + " StackTrace : " + ex.StackTrace);
        //           msgError =
        //"Exception type " + ex.GetType() + Environment.NewLine +
        //"Exception message: " + ex.Message + Environment.NewLine +
        //"Stack trace: " + ex.StackTrace + Environment.NewLine;
        //           if (ex.InnerException != null)
        //           {
        //               msgError += "---BEGIN InnerException--- " + Environment.NewLine +
        //                          "Exception type " + ex.InnerException.GetType() + Environment.NewLine +
        //                          "Exception message: " + ex.InnerException.Message + Environment.NewLine +
        //                          "Stack trace: " + ex.InnerException.StackTrace + Environment.NewLine +
        //                          "---END Inner Exception";
        //           }
        //       }
        //       finally
        //       {
        //           objDocFile = null;
        //           oCustom = null;
        //           objMiss = null; objSave = null;
        //           headerg = null;
        //           rangeg = null;

        //           if (objDoc != null)
        //               objDoc.Close(ref objSave, ref objMiss, ref objMiss);
        //           objDoc = null;

        //           if (wordApp != null)
        //               wordApp.Quit(ref objSave, ref objMiss, ref objMiss);
        //           wordApp = null;
        //           GC.Collect();
        //       }
        //       return bResult;
        //   }

        public bool Convert_DOcument_To_PDF(string szFilePath)
        {
           
            this._objINI = new clsReadAppXml(this._szAppXmlPath);
            bool flag = true;
            ChannelFactory<Process_Word_Document.ICalcService> channelFactory = null;
            try
            {
                try
                {
                   
                    string locationVariable = this._objINI.GetLocationVariable(this._objINI.GetCurrentLocation(), "", "Selfip");
                 
                    NetTcpBinding netTcpBinding = new NetTcpBinding()
                    {
                        SendTimeout = new TimeSpan(0, 30, 0),
                        ReceiveTimeout = new TimeSpan(0, 30, 0),
                        OpenTimeout = new TimeSpan(0, 30, 0),
                        CloseTimeout = new TimeSpan(0, 30, 0)
                    };
                    EndpointAddress endpointAddress = new EndpointAddress(string.Concat("net.tcp://", locationVariable, "/CalcService"));
                    ChannelFactory<Process_Word_Document.ICalcService> channelFactory1 = new ChannelFactory<Process_Word_Document.ICalcService>(netTcpBinding, endpointAddress);
                    channelFactory = channelFactory1;
                    ChannelFactory<Process_Word_Document.ICalcService> channelFactory2 = channelFactory1;
                    try
                    {
                        Process_Word_Document.ICalcService calcService = channelFactory.CreateChannel();
                       
                        calcService.Convert_To_PDF(szFilePath);
                      
                        ((ICommunicationObject)channelFactory).Close();
                        channelFactory.Close();
                        channelFactory.Abort();
                        channelFactory = null;
                    }
                    finally
                    {
                        if (channelFactory2 != null)
                        {
                            ((IDisposable)channelFactory2).Dispose();
                        }
                    }
                }
                catch (Exception exception1)
                {
                    Exception exception = exception1;
                  
                    flag = false;
                    this.msgError = exception.Message;
                }
            }
            finally
            {
                this._objINI.Dispose();
                this._objINI = null;
                if (channelFactory != null)
                {
                    channelFactory.Close();
                    channelFactory.Abort();
                    channelFactory = null;
                }
            }
            return flag;
        }

        public bool Process_Word_Document(string szFilePath)
        {
           
            this._objINI = new clsReadAppXml(this._szAppXmlPath);
            bool flag = true;
            ChannelFactory<Process_Word_Document.ICalcService> channelFactory = null;
            try
            {
                try
                {
                   
                    string locationVariable = this._objINI.GetLocationVariable(this._objINI.GetCurrentLocation(), "", "Selfip");
                  
                    NetTcpBinding netTcpBinding = new NetTcpBinding()
                    {
                        SendTimeout = new TimeSpan(0, 30, 0),
                        ReceiveTimeout = new TimeSpan(0, 30, 0),
                        OpenTimeout = new TimeSpan(0, 30, 0),
                        CloseTimeout = new TimeSpan(0, 30, 0),
                        MaxBufferPoolSize = (long)2147483647,
                        MaxReceivedMessageSize = (long)2147483647
                    };
                    netTcpBinding.ReaderQuotas.MaxDepth = 2147483647;
                    netTcpBinding.ReaderQuotas.MaxStringContentLength = 2147483647;
                    netTcpBinding.ReaderQuotas.MaxArrayLength = 2147483647;
                    netTcpBinding.ReaderQuotas.MaxBytesPerRead = 2147483647;
                    netTcpBinding.ReaderQuotas.MaxNameTableCharCount = 2147483647;
                    netTcpBinding.Security.Mode = SecurityMode.Transport;
                    netTcpBinding.Security.Transport.ClientCredentialType = TcpClientCredentialType.Windows;
                    EndpointAddress endpointAddress = new EndpointAddress(string.Concat("net.tcp://", locationVariable, "/CalcService"));
                    ChannelFactory<Process_Word_Document.ICalcService> channelFactory1 = new ChannelFactory<Process_Word_Document.ICalcService>(netTcpBinding, endpointAddress);
                    channelFactory = channelFactory1;
                    ChannelFactory<Process_Word_Document.ICalcService> channelFactory2 = channelFactory1;
                    try
                    {
                        channelFactory.CreateChannel().Process_Word_Document(szFilePath);
                        ((ICommunicationObject)channelFactory).Close();
                        channelFactory.Close();
                        channelFactory.Abort();
                        channelFactory = null;
                    }
                    finally
                    {
                        if (channelFactory2 != null)
                        {
                            ((IDisposable)channelFactory2).Dispose();
                        }
                    }
                }
                catch (Exception exception1)
                {
                    Exception exception = exception1;
                    flag = false;
                    this.msgError = exception.Message;
                }
            }
            finally
            {
                this._objINI.Dispose();
                this._objINI = null;
                if (channelFactory != null)
                {
                    channelFactory.Close();
                    channelFactory.Abort();
                    channelFactory = null;
                }
            }
            return flag;
        }

        public bool Update_Document_Properties(string szFilePath)
        {
           
            this._objINI = new clsReadAppXml(this._szAppXmlPath);
            bool flag = true;
            ChannelFactory<Process_Word_Document.ICalcService> channelFactory = null;
            try
            {
                try
                {
                   
                    string locationVariable = this._objINI.GetLocationVariable(this._objINI.GetCurrentLocation(), "", "Selfip");
                    NetTcpBinding netTcpBinding = new NetTcpBinding()
                    {
                        SendTimeout = new TimeSpan(0, 30, 0),
                        ReceiveTimeout = new TimeSpan(0, 30, 0),
                        OpenTimeout = new TimeSpan(0, 30, 0),
                        CloseTimeout = new TimeSpan(0, 30, 0),
                        MaxBufferPoolSize = (long)2147483647,
                        MaxReceivedMessageSize = (long)2147483647
                    };
                    netTcpBinding.ReaderQuotas.MaxDepth = 2147483647;
                    netTcpBinding.ReaderQuotas.MaxStringContentLength = 2147483647;
                    netTcpBinding.ReaderQuotas.MaxArrayLength = 2147483647;
                    netTcpBinding.ReaderQuotas.MaxBytesPerRead = 2147483647;
                    netTcpBinding.ReaderQuotas.MaxNameTableCharCount = 2147483647;
                    netTcpBinding.Security.Mode = SecurityMode.Transport;
                    netTcpBinding.Security.Transport.ClientCredentialType = TcpClientCredentialType.Windows;
                    EndpointAddress endpointAddress = new EndpointAddress(string.Concat("net.tcp://", locationVariable, "/CalcService"));
                    ChannelFactory<Process_Word_Document.ICalcService> channelFactory1 = new ChannelFactory<Process_Word_Document.ICalcService>(netTcpBinding, endpointAddress);
                    channelFactory = channelFactory1;
                    ChannelFactory<Process_Word_Document.ICalcService> channelFactory2 = channelFactory1;
                    try
                    {
                        Process_Word_Document.ICalcService calcService = channelFactory.CreateChannel();
                        calcService.Update_Document_Properties(szFilePath);
                        ((ICommunicationObject)channelFactory).Close();
                        channelFactory.Close();
                        channelFactory.Abort();
                        channelFactory = null;
                    }
                    finally
                    {
                        if (channelFactory2 != null)
                        {
                            ((IDisposable)channelFactory2).Dispose();
                        }
                    }
                }
                catch (Exception exception1)
                {
                    Exception exception = exception1;
                    flag = false;
                    this.msgError = exception.Message;
                }
            }
            finally
            {
                this._objINI.Dispose();
                this._objINI = null;
                if (channelFactory != null)
                {
                    channelFactory.Close();
                    channelFactory.Abort();
                    channelFactory = null;
                }
            }
            return flag;
        }

        public void Dispose()
        {
        }


        #endregion

        #region .....Private Functions .....
        #endregion
    }
}
