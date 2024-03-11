using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.IO;

namespace eDocsDN_OpenXml_Operations
{
    public class ClsUpdate_Comment_Author
    {
        #region .... Property ...
        public string msgError { get; set; }
        #endregion

        #region .... Constructor ...
        public ClsUpdate_Comment_Author()
        {
            msgError = string.Empty;
        }
        #endregion

        #region .... Public Method ....
        public bool Set_Author_To_Comments(string szFilePath, string szCurruntUser, DateTime dtDateTime)
        {
            WordprocessingDocument _objDoc;
            bool bResult = true;
            msgError = "";
            try
            {
                if (System.IO.File.Exists(szFilePath))
                {
                    using (_objDoc = WordprocessingDocument.Open(szFilePath, true))
                    {
                        WordprocessingCommentsPart commentsPart = _objDoc.MainDocumentPart.WordprocessingCommentsPart;
                        if (commentsPart != null && commentsPart.Comments != null)
                        {
                            foreach (Comment comment in commentsPart.Comments.Elements<Comment>())
                            {
                                if (dtDateTime < comment.Date.Value)
                                {
                                    Regex initials = new Regex(@"(\b[a-zA-Z])[a-zA-Z]* ?");
                                    string init = initials.Replace(szCurruntUser, "$1");
                                    comment.Initials = init;
                                    comment.Author = szCurruntUser;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                bResult = false;
                msgError = ex.ToString();
            }
            finally
            { }
            return bResult;

        }
        public Stream Set_Author_To_Comments(Stream strmDocument, string szCurruntUser, DateTime dtDateTime)
        {
            WordprocessingDocument _objDoc = null;
            msgError = "";
            try
            {
                _objDoc = WordprocessingDocument.Open(strmDocument, true);
                WordprocessingCommentsPart commentsPart = _objDoc.MainDocumentPart.WordprocessingCommentsPart;
                if (commentsPart != null && commentsPart.Comments != null)
                {
                    foreach (Comment comment in commentsPart.Comments.Elements<Comment>())
                    {
                        if (dtDateTime < comment.Date.Value)
                        {
                            Regex initials = new Regex(@"(\b[a-zA-Z])[a-zA-Z]* ?");
                            string init = initials.Replace(szCurruntUser, "$1");
                            comment.Initials = init;
                            comment.Author = szCurruntUser;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                msgError = ex.ToString();
            }
            finally
            {
                if (_objDoc != null)
                {
                    _objDoc.Close();
                    _objDoc.Dispose();
                }
                _objDoc = null;
            }
            return strmDocument;

        }

        #endregion
    }
}
