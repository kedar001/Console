using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using eDocsDN_DocX;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using docx = eDocsDN_DocX;
using System.Xml.Linq;
using System.Xml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Office2013.PowerPoint;
using System.Text.RegularExpressions;

using System.IO.Compression;
using DocumentFormat.OpenXml.ExtendedProperties;

namespace eDocDN_Document_Pre_Check
{
    public class ClsDocumentPre_Check
    {

        bool _bAllowEmbededImages = false;
        public string msgError { get; set; }
        public string FileName { get; set; }
        public Stream _strmDocument { get; set; }

        public ClsDocumentPre_Check(string szFilePath, bool bAllowEmbededImages)
        {
            msgError = string.Empty;
            _bAllowEmbededImages = bAllowEmbededImages;
            FileName = szFilePath;
            _strmDocument = null;
        }

        public ClsDocumentPre_Check(Stream strmDocument, bool bAllowEmbededImages)
        {
            msgError = string.Empty;
            _bAllowEmbededImages = bAllowEmbededImages;
            _strmDocument = strmDocument;
            FileName = string.Empty;
        }

        public bool PreCheck_Document()
        {
            bool bResult = true;

            try
            {
                if (IsPasswordProtectedDocument())
                    throw new Exception("Document is Password Protected");
                CheckInvalidHyperlink();
                if (!_bAllowEmbededImages)
                    CheckForImbededImages();
            }
            catch (Exception ex)
            {
                bResult = false;
                msgError = ex.Message;
            }
            finally
            {
            }
            return bResult;
        }

        private static Uri FixUri(string brokenUri)
        {
            return new Uri("http://broken-link/");
        }


        private void CheckInvalidHyperlink()
        {
            WordprocessingDocument wDoc;
            docx.DocX oDocx = null;
            Uri uriResult;
            string szURI;
            string szTempPath = System.Windows.Forms.Application.StartupPath + "\\temp.docx";
            try
            {
                if (_strmDocument != null)
                {
                    using (wDoc = WordprocessingDocument.Open(_strmDocument, false))
                    {
                        var elementCount = wDoc.MainDocumentPart.Document.Descendants().Count();
                        //foreach (FieldCode field in wDoc.MainDocumentPart.RootElement.Descendants<FieldCode>())
                        //{
                        //    if (field.InnerText.StartsWith("HYPERLINK"))
                        //    {
                        //        szURI = field.InnerText.Replace("HYPERLINK", "");
                        //        szURI = szURI.Trim().Replace("\"", "");
                        //        //bool result = Uri.TryCreate(szURI, UriKind.Absolute, out uriResult) && (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);
                        //        //if (!result)
                        //        //    throw new Exception("Invalid Hyperlink");

                        //        //if (!Uri.IsWellFormedUriString(szURI, UriKind.Absolute))
                        //        //throw new Exception("Invalid Hyperlink");

                        //        if (!IsValidURL(szURI))
                        //            throw new OpenXmlPackageException("Invalid Hyperlink");


                        //    }
                        //}

                    }
                }
                else
                {
                    using (wDoc = WordprocessingDocument.Open(FileName, false))
                    {
                        var elementCount = wDoc.MainDocumentPart.Document.Descendants().Count();
                        //foreach (FieldCode field in wDoc.MainDocumentPart.RootElement.Descendants<FieldCode>())
                        //{

                        //    if (field.InnerText.StartsWith("HYPERLINK"))
                        //    {
                        //        szURI = field.InnerText.Replace("HYPERLINK", "");
                        //        szURI = szURI.Trim().Replace("\"", "");
                        //        //bool result = Uri.TryCreate(szURI, UriKind.Absolute, out uriResult) && (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);
                        //        //if (!result)
                        //        //    throw new Exception("Invalid Hyperlink");

                        //        //if (!Uri.IsWellFormedUriString(szURI, UriKind.Absolute))
                        //        //    throw new Exception("Invalid Hyperlink");

                        //        if (!IsValidURL(szURI))
                        //            throw new OpenXmlPackageException("Invalid Hyperlink");
                        //    }

                        //}

                    }
                }
            }
            catch (OpenXmlPackageException e)
            {
                if (e.ToString().Contains("Invalid Hyperlink"))
                {
                    if (e.ToString().Contains("Invalid Hyperlink"))
                    {
                        if (_strmDocument != null)
                        {
                            File.WriteAllBytes(szTempPath, _strmDocument.ReadAllBytes());
                            using (FileStream fileStream = new FileStream(szTempPath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                            {
                                UriFixer.FixInvalidUri(fileStream, brokenUri => FixUri(brokenUri));
                            }
                            _strmDocument = Convert_Document_To_Stream(File.ReadAllBytes(szTempPath));
                            File.Delete(szTempPath);
                        }
                        else
                        {
                            using (FileStream fs = new FileStream(FileName, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                            {
                                UriFixer.FixInvalidUri(fs, brokenUri => FixUri(brokenUri));
                            }
                        }
                        CheckInvalidHyperlink();
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                oDocx = null;
                wDoc = null;
            }
        }

        bool IsValidURL(string URL)
        {
            bool isUri = Uri.IsWellFormedUriString(URL, UriKind.RelativeOrAbsolute);
            if (!isUri)
                return false;

            string Pattern = @"^(?:http(s)?:\/\/)?[\w.-]+(?:\.[\w\.-]+)+[\w\-\._~:/?#[\]@!\$&'\(\)\*\+,;=.]+$";
            Regex Rgx = new Regex(Pattern, RegexOptions.Compiled | RegexOptions.IgnoreCase);
            return Rgx.IsMatch(URL);
        }



        internal static Stream Convert_Document_To_Stream(byte[] arrDocument)
        {
            MemoryStream strmDocument = new MemoryStream();
            strmDocument.Write(arrDocument, 0, (int)arrDocument.Length);
            return strmDocument;
        }


        private void CheckForImbededImages()
        {
            WordprocessingDocument wDoc = null;
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";  // Change on Here Before send code to Kedar
            XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";	// Change on Here Before send code to Kedar

            try
            {
                //..OLEObject 
                if (_strmDocument != null)
                    wDoc = WordprocessingDocument.Open(_strmDocument, false);
                else
                    wDoc = WordprocessingDocument.Open(FileName, false);



                if (wDoc.MainDocumentPart.GetPartsCountOfType<EmbeddedPackagePart>() > 0)
                {
                    throw new Exception("There is EmbededPackage like Excel Spreadsheet, Word Document");
                }

                foreach (Ovml.OleObject item in wDoc.MainDocumentPart.Document.Descendants<Ovml.OleObject>())
                {
                    if (item.Ancestors<DeletedRun>() != null)
                        if (!(item.Ancestors<DeletedRun>().FirstOrDefault() is DeletedRun))
                        {
                            throw new Exception("Please check Document contains Embeded Images");
                        }
                }


                foreach (HeaderPart Header in wDoc.MainDocumentPart.HeaderParts)
                {
                    foreach (Ovml.OleObject item in Header.RootElement.Descendants<Ovml.OleObject>())
                    {
                        if (item.Ancestors<DeletedRun>() != null)
                            if (!(item.Ancestors<DeletedRun>().FirstOrDefault() is DeletedRun))
                            {
                                throw new Exception("Please check Document contains Embeded Images");
                            }

                    }

                }
                foreach (FooterPart Footer in wDoc.MainDocumentPart.FooterParts)
                {
                    foreach (Ovml.OleObject item in Footer.RootElement.Descendants<Ovml.OleObject>())
                    {
                        if (item.Ancestors<DeletedRun>() != null)
                            if (!(item.Ancestors<DeletedRun>().FirstOrDefault() is DeletedRun))
                            {
                                throw new Exception("Please check Document contains Embeded Images");
                            }
                    }

                }


            }
            finally
            {
                if (wDoc != null)
                    wDoc.Dispose();
                wDoc = null;
            }
        }

        public bool IsPasswordProtectedDocument()
        {
            bool _bResult = true;
            WordprocessingDocument wDoc = null;
            try
            {
                if (_strmDocument != null)
                    wDoc = WordprocessingDocument.Open(_strmDocument, false);
                else
                    wDoc = WordprocessingDocument.Open(FileName, false);


                WriteProtection Wp = wDoc.MainDocumentPart.DocumentSettingsPart.Settings.GetFirstChild<WriteProtection>();
                if (Wp == null)
                    _bResult = false;
                if (Wp != null)
                    if (Wp.CryptographicAlgorithmClass.HasValue)
                        _bResult = true;

                //if (!_bResult)
                //{
                //    DocumentProtection dp = wDoc.MainDocumentPart.DocumentSettingsPart.Settings.GetFirstChild<DocumentProtection>();
                //    if (dp == null)
                //        _bResult = false;
                //    if (dp != null && (dp.Edit != null))
                //        _bResult = true;
                //}


            }
            finally
            {
                if (wDoc != null)
                    wDoc.Dispose();
                wDoc = null;
            }
            return _bResult;
        }



        //private XDocument GetXDocument(this OpenXmlPart part)
        //{
        //    XDocument xdoc = part.Annotation<XDocument>();
        //    if (xdoc != null)
        //        return xdoc;
        //    using (StreamReader sr = new StreamReader(part.GetStream()))
        //    using (XmlReader xr = XmlReader.Create(sr))
        //        xdoc = XDocument.Load(xr);
        //    part.AddAnnotation(xdoc);
        //    return xdoc;
        //}



    }

    public static class UriFixer
    {
        public static void FixInvalidUri(Stream fs, Func<string, Uri> invalidUriHandler)
        {
            XNamespace relNs = "http://schemas.openxmlformats.org/package/2006/relationships";
            using (ZipArchive za = new ZipArchive(fs, ZipArchiveMode.Update))
            {
                foreach (var entry in za.Entries.ToList())
                {
                    if (!entry.Name.EndsWith(".rels"))
                        continue;
                    bool replaceEntry = false;
                    XDocument entryXDoc = null;
                    using (var entryStream = entry.Open())
                    {
                        try
                        {
                            entryXDoc = XDocument.Load(entryStream);
                            if (entryXDoc.Root != null && entryXDoc.Root.Name.Namespace == relNs)
                            {
                                var urisToCheck = entryXDoc
                                    .Descendants(relNs + "Relationship")
                                    .Where(r => r.Attribute("TargetMode") != null && (string)r.Attribute("TargetMode") == "External");
                                foreach (var rel in urisToCheck)
                                {
                                    var target = (string)rel.Attribute("Target");
                                    if (target != null)
                                    {
                                        try
                                        {
                                            Uri uri = new Uri(target);
                                        }
                                        catch (UriFormatException)
                                        {
                                            Uri newUri = invalidUriHandler(target);
                                            rel.Attribute("Target").Value = newUri.ToString();
                                            replaceEntry = true;
                                        }
                                    }
                                }
                            }
                        }
                        catch (XmlException)
                        {
                            continue;
                        }
                    }
                    if (replaceEntry)
                    {
                        var fullName = entry.FullName;
                        entry.Delete();
                        var newEntry = za.CreateEntry(fullName);
                        using (StreamWriter writer = new StreamWriter(newEntry.Open()))
                        using (XmlWriter xmlWriter = XmlWriter.Create(writer))
                        {
                            entryXDoc.WriteTo(xmlWriter);
                        }
                    }
                }
            }
        }
    }

    public static class StreamExtensions
    {
        public static byte[] ReadAllBytes(this Stream instream)
        {
            if (instream is MemoryStream)
                return ((MemoryStream)instream).ToArray();

            using (var memoryStream = new MemoryStream())
            {
                instream.CopyTo(memoryStream);
                return memoryStream.ToArray();
            }
        }
    }


}
