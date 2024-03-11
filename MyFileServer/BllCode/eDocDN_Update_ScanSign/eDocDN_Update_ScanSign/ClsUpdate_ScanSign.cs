using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.IO;

using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using V = DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Wordprocessing;
using Lock = DocumentFormat.OpenXml.Vml.Office.Lock;
using DocumentFormat.OpenXml;
using wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using a = DocumentFormat.OpenXml.Drawing;
using pic = DocumentFormat.OpenXml.Drawing.Pictures;



namespace eDocDN_Update_ScanSign
{
    public class ClsUpdate_ScanSign : IDisposable
    {
        #region .... Variable Declaration ....


        Stream _strmDocumentStream = null;


        #endregion

        #region .... Properties ....

        public string msgError { get; set; }
        private string FilePath { get; set; }

        #endregion

        #region .... Constuctor ....

        public ClsUpdate_ScanSign(Stream strmDocument)
        {
            msgError = "";
            FilePath = string.Empty;
            _strmDocumentStream = strmDocument;
        }
        public ClsUpdate_ScanSign(string szFileName)
        {
            msgError = "";
            FilePath = szFileName;
        }


        #endregion

        #region .... Public Functions ...

        public Stream UpdateScanSign(Dictionary<string, string> szScanSignOfUser, bool bRemoveScanSign)
        {
            msgError = "";
            try
            {
                _strmDocumentStream = ScanSign(szScanSignOfUser, bRemoveScanSign);
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            finally
            {

            }
            return _strmDocumentStream;
        }

        public Stream RemoveScanSign()
        {
            string szScansignFilePath = string.Empty;
            Run rScanSignToRemove = null;
            string szScanSignCustumProperty = string.Empty;
            WordprocessingDocument doc = null;
            try
            {
                //DR NO:919833
                if (_strmDocumentStream != null)
                    doc = WordprocessingDocument.Open(_strmDocumentStream, true);
                else
                    doc = WordprocessingDocument.Open(FilePath, true);

                #region.... Remove Scan Sign ...
                List<Table> lstTables = doc.MainDocumentPart.Document.Body.Elements<Table>().ToList();
                foreach (var item in lstTables)
                {
                    List<TableRow> tableRow = item.Elements<TableRow>().ToList();
                    foreach (var item1 in tableRow)
                    {
                        List<TableCell> tableCell = item1.Elements<TableCell>().ToList();
                        foreach (var TblCell in tableCell)
                        {
                            List<Paragraph> Para = TblCell.Elements<Paragraph>().ToList();
                            foreach (var p in Para)
                            {
                                List<Run> Run = p.Elements<Run>().ToList();
                                foreach (var r in Run)
                                {
                                    var Pic = r.Elements<Drawing>().FirstOrDefault();
                                    if (Pic != null)
                                    {
                                        //.. changed by Kedar for DR No. 920060 on 30-01-2016
                                        if (Pic.Inline != null && Pic.Inline.DocProperties != null && Pic.Inline.DocProperties.Description != null)
                                        {
                                            szScanSignCustumProperty = Pic.Inline.DocProperties.Description.Value;
                                            if (szScanSignCustumProperty.ToString().ToUpper().Contains("WORKING\\DOC"))
                                            {
                                                rScanSignToRemove = r;
                                            }
                                            else
                                            {
                                                if (Pic.Inline.DocProperties.Name.Value.Contains("ScanSignDrawing_"))
                                                {
                                                    string instructionText = String.Format(" MERGEFIELD  {0}  \\* MERGEFORMAT", szScanSignCustumProperty);
                                                    SimpleField simpleField1 = new SimpleField() { Instruction = instructionText };
                                                    Run run1 = new Run();
                                                    RunProperties runProperties1 = new RunProperties();
                                                    NoProof noProof1 = new NoProof();
                                                    runProperties1.Append(noProof1);
                                                    Text text1 = new Text();
                                                    text1.Text = String.Format("{0}", szScanSignCustumProperty);
                                                    run1.Append(runProperties1);
                                                    run1.Append(text1);
                                                    simpleField1.Append(run1);
                                                    p.Append(new OpenXmlElement[] { simpleField1 });
                                                    a.Blip blip1 = Pic.Descendants<a.Blip>().FirstOrDefault();
                                                    IdPartPair idpp = doc.MainDocumentPart.Parts.Where(pa => pa.RelationshipId == blip1.Embed).FirstOrDefault();
                                                    if (idpp != null)
                                                        doc.MainDocumentPart.DeletePart(idpp.RelationshipId);
                                                    rScanSignToRemove = r;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            if (rScanSignToRemove != null)
                            {
                                rScanSignToRemove.Remove();
                                rScanSignToRemove = null;
                            }
                        }
                    }
                }

                #endregion

                #region .... Check in Header ....
                if (doc.MainDocumentPart.HeaderParts != null)
                {
                    foreach (var header in doc.MainDocumentPart.HeaderParts)
                    {
                        #region.... Table ...
                        if (header.Header.Descendants<Table>() != null)
                        {

                            List<Table> lstFooterTables = header.Header.Descendants<Table>().ToList();
                            foreach (var item in lstFooterTables)
                            {
                                List<TableRow> tableRow = item.Elements<TableRow>().ToList();
                                foreach (var item1 in tableRow)
                                {
                                    List<TableCell> tableCell = item1.Elements<TableCell>().ToList();
                                    foreach (var TblCell in tableCell)
                                    {
                                        List<Paragraph> Para = TblCell.Elements<Paragraph>().ToList();
                                        foreach (var p in Para)
                                        {
                                            List<Run> Run = p.Elements<Run>().ToList();
                                            foreach (var r in Run)
                                            {
                                                var Pic = r.Elements<Drawing>().FirstOrDefault();
                                                if (Pic != null)
                                                {
                                                    //.. changed by Kedar for DR No. 920060 on 30-01-2016
                                                    if (Pic.Inline != null && Pic.Inline.DocProperties != null && Pic.Inline.DocProperties.Description != null)
                                                    {
                                                        szScanSignCustumProperty = Pic.Inline.DocProperties.Description.Value;
                                                        if (szScanSignCustumProperty.ToString().ToUpper().Contains("WORKING\\DOC"))
                                                        {
                                                            rScanSignToRemove = r;
                                                        }
                                                        else
                                                        {
                                                            if (Pic.Inline.DocProperties.Name.Value.Contains("ScanSignDrawing_"))
                                                            {
                                                                szScanSignCustumProperty = Pic.Inline.DocProperties.Description.Value;
                                                                string instructionText = String.Format(" MERGEFIELD  {0}  \\* MERGEFORMAT", szScanSignCustumProperty);
                                                                SimpleField simpleField1 = new SimpleField() { Instruction = instructionText };
                                                                Run run1 = new Run();
                                                                RunProperties runProperties1 = new RunProperties();
                                                                NoProof noProof1 = new NoProof();
                                                                runProperties1.Append(noProof1);
                                                                Text text1 = new Text();
                                                                text1.Text = String.Format("{0}", szScanSignCustumProperty);
                                                                run1.Append(runProperties1);
                                                                run1.Append(text1);
                                                                simpleField1.Append(run1);
                                                                p.Append(new OpenXmlElement[] { simpleField1 });
                                                                //p.Append(new SimpleField(new Run(new RunProperties(new NoProof()), new Text(szScanSignCustumProperty))));
                                                                a.Blip blip1 = Pic.Descendants<a.Blip>().FirstOrDefault();
                                                                IdPartPair idpp = header.Parts.Where(pa => pa.RelationshipId == blip1.Embed).FirstOrDefault();
                                                                if (idpp != null)
                                                                    header.DeletePart(idpp.RelationshipId);
                                                                rScanSignToRemove = r;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        if (rScanSignToRemove != null)
                                        {
                                            rScanSignToRemove.Remove();
                                            rScanSignToRemove = null;
                                        }
                                    }
                                }
                            }
                        }
                        #endregion
                    }
                }
                #endregion

                #region .... Check in Footer ....
                if (doc.MainDocumentPart.FooterParts != null)
                {
                    foreach (var footer in doc.MainDocumentPart.FooterParts)
                    {
                        #region.... Table ...
                        if (footer.Footer.Descendants<Table>() != null)
                        {

                            List<Table> lstFooterTables = footer.Footer.Descendants<Table>().ToList();
                            foreach (var item in lstFooterTables)
                            {
                                List<TableRow> tableRow = item.Elements<TableRow>().ToList();
                                foreach (var item1 in tableRow)
                                {
                                    List<TableCell> tableCell = item1.Elements<TableCell>().ToList();
                                    foreach (var TblCell in tableCell)
                                    {
                                        List<Paragraph> Para = TblCell.Elements<Paragraph>().ToList();
                                        foreach (var p in Para)
                                        {
                                            List<Run> Run = p.Elements<Run>().ToList();
                                            foreach (var r in Run)
                                            {
                                                var Pic = r.Elements<Drawing>().FirstOrDefault();
                                                if (Pic != null)
                                                {

                                                    //.. changed by Kedar for DR No. 920060 on 30-01-2016
                                                    if (Pic.Inline != null && Pic.Inline.DocProperties != null && Pic.Inline.DocProperties.Description != null)
                                                    {
                                                        szScanSignCustumProperty = Pic.Inline.DocProperties.Description.Value;
                                                        if (szScanSignCustumProperty.ToString().ToUpper().Contains("WORKING\\DOC"))
                                                        {
                                                            rScanSignToRemove = r;
                                                        }
                                                        else
                                                        {
                                                            if (Pic.Inline.DocProperties.Name.Value.Contains("ScanSignDrawing_"))
                                                            {
                                                                szScanSignCustumProperty = Pic.Inline.DocProperties.Description.Value;
                                                                string instructionText = String.Format(" MERGEFIELD  {0}  \\* MERGEFORMAT", szScanSignCustumProperty);
                                                                SimpleField simpleField1 = new SimpleField() { Instruction = instructionText };
                                                                Run run1 = new Run();
                                                                RunProperties runProperties1 = new RunProperties();
                                                                NoProof noProof1 = new NoProof();
                                                                runProperties1.Append(noProof1);
                                                                Text text1 = new Text();
                                                                text1.Text = String.Format("{0}", szScanSignCustumProperty);
                                                                run1.Append(runProperties1);
                                                                run1.Append(text1);
                                                                simpleField1.Append(run1);
                                                                p.Append(new OpenXmlElement[] { simpleField1 });
                                                                //p.Append(new SimpleField(new Run(new RunProperties(new NoProof()), new Text(szScanSignCustumProperty))));
                                                                a.Blip blip1 = Pic.Descendants<a.Blip>().FirstOrDefault();
                                                                IdPartPair idpp = footer.Parts.Where(pa => pa.RelationshipId == blip1.Embed).FirstOrDefault();
                                                                if (idpp != null)
                                                                    footer.DeletePart(idpp.RelationshipId);
                                                                rScanSignToRemove = r;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        if (rScanSignToRemove != null)
                                        {
                                            rScanSignToRemove.Remove();
                                            rScanSignToRemove = null;
                                        }
                                    }
                                }
                            }
                        }
                        #endregion
                    }
                }
                #endregion

            }
            catch (Exception ex)
            {
                msgError = ex.StackTrace;
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close();
                    doc.Dispose();
                }
                doc = null;

            }
            return _strmDocumentStream;
        }
        #endregion

        #region .... Private Functions ...

        private Stream ScanSign(Dictionary<string, string> ScanSignOfUser, bool bRemoveScanSign)
        {
            string szScansignFilePath = string.Empty;
            Run rScanSignToRemove = null;
            Paragraph paraAuthorScanSign = null;
            bool bFieldCode = false;
            string szSCanSignOfUser = string.Empty;
            WordprocessingDocument doc = null;
            try
            {
                if (string.IsNullOrEmpty(FilePath))
                    doc = WordprocessingDocument.Open(_strmDocumentStream, true);
                else
                    doc = WordprocessingDocument.Open(FilePath, true);

                #region.... Table ...
                List<Table> lstTables = doc.MainDocumentPart.Document.Body.Elements<Table>().ToList();
                foreach (var item in lstTables)
                {
                    List<TableRow> tableRow = item.Elements<TableRow>().ToList();
                    foreach (var item1 in tableRow)
                    {
                        List<TableCell> tableCell = item1.Elements<TableCell>().ToList();
                        foreach (var TblCell in tableCell)
                        {
                            List<Paragraph> Para = TblCell.Elements<Paragraph>().ToList();
                            foreach (var p in Para)
                            {
                                List<Run> Run = p.Elements<Run>().ToList();
                                foreach (var r in Run)
                                {
                                    var Pic = r.Elements<Drawing>().FirstOrDefault();
                                    if (Pic != null)
                                    {
                                        //.. changed by Kedar for DR No. 920060 on 30-01-2016
                                        if (Pic.Inline != null && Pic.Inline.DocProperties != null && Pic.Inline.DocProperties.Description != null)
                                        {
                                            if (Pic.Inline.DocProperties.Name.Value.Contains("ScanSignDrawing_"))
                                                rScanSignToRemove = r;
                                        }
                                    }
                                    var ScanSign = r.Elements<FieldCode>().FirstOrDefault();
                                    if (ScanSign != null)
                                    {
                                        //A1_Sign
                                        foreach (var dicItem in ScanSignOfUser)
                                        {
                                            if (ScanSign.InnerText.Contains(dicItem.Key))
                                            {
                                                bFieldCode = true;
                                                r.Remove();
                                                szSCanSignOfUser = dicItem.Key;
                                                szScansignFilePath = dicItem.Value;
                                                szScansignFilePath = szScansignFilePath.Replace("\\", "\\\\");
                                                bFieldCode = true;
                                                paraAuthorScanSign = p;
                                            }
                                        }
                                    }
                                }
                                if (!bFieldCode)
                                {
                                    var ScanSign = p.Elements<SimpleField>().FirstOrDefault();
                                    if (ScanSign != null)
                                    {
                                        //A1_Sign
                                        foreach (var dicItem in ScanSignOfUser)
                                        {
                                            if (ScanSign.InnerText.Contains(dicItem.Key))
                                            {
                                                bFieldCode = true;
                                                ScanSign.Remove();
                                                szSCanSignOfUser = dicItem.Key;
                                                szScansignFilePath = dicItem.Value;
                                                szScansignFilePath = szScansignFilePath.Replace("\\", "\\\\");
                                                bFieldCode = true;
                                                paraAuthorScanSign = p;
                                            }
                                        }
                                    }
                                }
                            }
                            if (rScanSignToRemove != null)
                            {
                                rScanSignToRemove.Remove();
                                rScanSignToRemove = null;
                            }
                            if (!bRemoveScanSign)
                            {
                                if (paraAuthorScanSign != null)
                                {
                                    AddParts(doc, szScansignFilePath, paraAuthorScanSign, szSCanSignOfUser);
                                    paraAuthorScanSign = null;
                                }
                            }
                            bFieldCode = false;
                            rScanSignToRemove = null;
                        }
                    }
                }

                #endregion

                #region .... Check in Header ....
                if (doc.MainDocumentPart.HeaderParts != null)
                {
                    foreach (var Header in doc.MainDocumentPart.HeaderParts)
                    {
                        #region.... Table ...
                        if (Header.Header.Descendants<Table>() != null)
                        {

                            List<Table> lstFooterTables = Header.Header.Descendants<Table>().ToList();
                            foreach (var item in lstFooterTables)
                            {
                                List<TableRow> tableRow = item.Elements<TableRow>().ToList();
                                foreach (var item1 in tableRow)
                                {
                                    List<TableCell> tableCell = item1.Elements<TableCell>().ToList();
                                    foreach (var TblCell in tableCell)
                                    {
                                        List<Paragraph> Para = TblCell.Elements<Paragraph>().ToList();
                                        foreach (var p in Para)
                                        {
                                            List<Run> Run = p.Elements<Run>().ToList();
                                            foreach (var r in Run)
                                            {
                                                var Pic = r.Elements<Drawing>().FirstOrDefault();
                                                if (Pic != null)
                                                {
                                                    //.. changed by Kedar for DR No. 920060 on 30-01-2016
                                                    if (Pic.Inline != null && Pic.Inline.DocProperties != null && Pic.Inline.DocProperties.Description != null)
                                                    {
                                                        if (Pic.Inline.DocProperties.Name.Value.Contains("ScanSignDrawing_"))
                                                            rScanSignToRemove = r;
                                                    }
                                                }
                                                var ScanSign = r.Elements<FieldCode>().FirstOrDefault();
                                                if (ScanSign != null)
                                                {
                                                    //A1_Sign
                                                    foreach (var dicItem in ScanSignOfUser)
                                                    {
                                                        if (ScanSign.InnerText.Contains(dicItem.Key))
                                                        {

                                                            bFieldCode = true;
                                                            r.Remove();
                                                            szSCanSignOfUser = dicItem.Key;
                                                            szScansignFilePath = dicItem.Value;
                                                            szScansignFilePath = szScansignFilePath.Replace("\\", "\\\\");
                                                            bFieldCode = true;
                                                            paraAuthorScanSign = p;
                                                        }
                                                    }
                                                }
                                            }
                                            if (!bFieldCode)
                                            {
                                                var ScanSign = p.Elements<SimpleField>().FirstOrDefault();
                                                if (ScanSign != null)
                                                {
                                                    //A1_Sign
                                                    foreach (var dicItem in ScanSignOfUser)
                                                    {
                                                        if (ScanSign.InnerText.Contains(dicItem.Key))
                                                        {
                                                            bFieldCode = true;
                                                            ScanSign.Remove();
                                                            szSCanSignOfUser = dicItem.Key;
                                                            szScansignFilePath = dicItem.Value;
                                                            szScansignFilePath = szScansignFilePath.Replace("\\", "\\\\");
                                                            bFieldCode = true;
                                                            paraAuthorScanSign = p;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        if (rScanSignToRemove != null)
                                        {
                                            rScanSignToRemove.Remove();
                                            rScanSignToRemove = null;
                                        }
                                        if (!bRemoveScanSign)
                                        {
                                            if (paraAuthorScanSign != null)
                                            {
                                                AddParts(Header, szScansignFilePath, paraAuthorScanSign, szSCanSignOfUser);
                                                paraAuthorScanSign = null;
                                            }
                                        }
                                        bFieldCode = false;
                                        rScanSignToRemove = null;
                                    }
                                }
                            }
                        }


                        #endregion
                    }
                }

                #endregion

                #region .... Check in Footer ....
                if (doc.MainDocumentPart.FooterParts != null)
                {
                    foreach (var footer in doc.MainDocumentPart.FooterParts)
                    {
                        #region.... Table ...
                        if (footer.Footer.Descendants<Table>() != null)
                        {

                            List<Table> lstFooterTables = footer.Footer.Descendants<Table>().ToList();
                            foreach (var item in lstFooterTables)
                            {
                                List<TableRow> tableRow = item.Elements<TableRow>().ToList();
                                foreach (var item1 in tableRow)
                                {
                                    List<TableCell> tableCell = item1.Elements<TableCell>().ToList();
                                    foreach (var TblCell in tableCell)
                                    {
                                        List<Paragraph> Para = TblCell.Elements<Paragraph>().ToList();
                                        foreach (var p in Para)
                                        {
                                            List<Run> Run = p.Elements<Run>().ToList();
                                            foreach (var r in Run)
                                            {
                                                var Pic = r.Elements<Drawing>().FirstOrDefault();
                                                if (Pic != null)
                                                {
                                                    //.. changed by Kedar for DR No. 920060 on 30-01-2016
                                                    if (Pic.Inline != null && Pic.Inline.DocProperties != null && Pic.Inline.DocProperties.Description != null)
                                                    {
                                                        if (Pic.Inline.DocProperties.Name.Value.Contains("ScanSignDrawing_"))
                                                            rScanSignToRemove = r;
                                                    }
                                                }
                                                var ScanSign = r.Elements<FieldCode>().FirstOrDefault();
                                                if (ScanSign != null)
                                                {
                                                    //A1_Sign
                                                    foreach (var dicItem in ScanSignOfUser)
                                                    {
                                                        if (ScanSign.InnerText.Contains(dicItem.Key))
                                                        {

                                                            bFieldCode = true;
                                                            r.Remove();
                                                            szSCanSignOfUser = dicItem.Key;
                                                            szScansignFilePath = dicItem.Value;
                                                            szScansignFilePath = szScansignFilePath.Replace("\\", "\\\\");
                                                            bFieldCode = true;
                                                            paraAuthorScanSign = p;
                                                        }
                                                    }
                                                }
                                            }
                                            if (!bFieldCode)
                                            {
                                                var ScanSign = p.Elements<SimpleField>().FirstOrDefault();
                                                if (ScanSign != null)
                                                {
                                                    //A1_Sign
                                                    foreach (var dicItem in ScanSignOfUser)
                                                    {
                                                        if (ScanSign.InnerText.Contains(dicItem.Key))
                                                        {
                                                            bFieldCode = true;
                                                            ScanSign.Remove();
                                                            szSCanSignOfUser = dicItem.Key;
                                                            szScansignFilePath = dicItem.Value;
                                                            szScansignFilePath = szScansignFilePath.Replace("\\", "\\\\");
                                                            bFieldCode = true;
                                                            paraAuthorScanSign = p;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        if (rScanSignToRemove != null)
                                        {
                                            rScanSignToRemove.Remove();
                                            rScanSignToRemove = null;
                                        }
                                        if (!bRemoveScanSign)
                                        {
                                            if (paraAuthorScanSign != null)
                                            {
                                                AddParts(footer, szScansignFilePath, paraAuthorScanSign, szSCanSignOfUser);
                                                paraAuthorScanSign = null;
                                            }
                                        }
                                        bFieldCode = false;
                                        rScanSignToRemove = null;
                                    }
                                }
                            }
                        }


                        #endregion
                    }
                }
                #endregion

            }
            finally
            {
                if (doc != null)
                {
                    doc.Close();
                    doc.Dispose();
                }
                doc = null;
            }
            return _strmDocumentStream;
        }

        private void AddParts(WordprocessingDocument parent, string imageFilePath, Paragraph P, string szRole)
        {
            long imageWidthEMU = 0;
            long imageHeightEMU = 0;

            var imagePart = parent.MainDocumentPart.AddNewPart<ImagePart>("image/jpeg", "rId" + System.IO.Path.GetFileNameWithoutExtension(imageFilePath));
            GenerateImagePart(imagePart, imageFilePath,
             ref imageWidthEMU, ref imageHeightEMU);
            var picture = GenerateMainDocumentPart(System.IO.Path.GetFileNameWithoutExtension(imageFilePath), imageWidthEMU, imageHeightEMU, szRole);
            P.AppendChild(new Run(picture));
        }
        private void AddParts(WordprocessingDocument parent, Stream strmScanSign, Paragraph P, string szRole)
        {
            long imageWidthEMU = 0;
            long imageHeightEMU = 0;

            var imagePart = parent.MainDocumentPart.AddNewPart<ImagePart>("image/jpeg", "rId" + strmScanSign.Length.ToString());
            GenerateImagePart(imagePart, strmScanSign,
             ref imageWidthEMU, ref imageHeightEMU);
            var picture = GenerateMainDocumentPart(strmScanSign.Length.ToString(), imageWidthEMU, imageHeightEMU, szRole);
            P.AppendChild(new Run(picture));
        }
        private void AddParts(FooterPart parent, Stream strmScanSign, Paragraph P, string szRole)
        {
            long imageWidthEMU = 0;
            long imageHeightEMU = 0;
            var imagePart = parent.AddNewPart<ImagePart>("image/jpeg", "rId" + strmScanSign.Length);
            GenerateImagePart(imagePart, strmScanSign, ref imageWidthEMU, ref imageHeightEMU);
            var picture = GenerateMainDocumentPart(strmScanSign.Length.ToString(), imageWidthEMU, imageHeightEMU, szRole);
            P.AppendChild(new Run(picture));
        }
        private void AddParts(HeaderPart parent, Stream strmScanSign, Paragraph P, string szRole)
        {
            long imageWidthEMU = 0;
            long imageHeightEMU = 0;
            var imagePart = parent.AddNewPart<ImagePart>("image/jpeg", "rId" + strmScanSign.Length.ToString());
            GenerateImagePart(imagePart, strmScanSign,
             ref imageWidthEMU, ref imageHeightEMU);
            var picture = GenerateMainDocumentPart(strmScanSign.Length.ToString(), imageWidthEMU, imageHeightEMU, szRole);
            P.AppendChild(new Run(picture));
        }

        private void AddParts(FooterPart parent, string imageFilePath, Paragraph P, string szRole)
        {
            long imageWidthEMU = 0;
            long imageHeightEMU = 0;
            var imagePart = parent.AddNewPart<ImagePart>("image/jpeg", "rId" + System.IO.Path.GetFileNameWithoutExtension(imageFilePath));
            GenerateImagePart(imagePart, imageFilePath,
             ref imageWidthEMU, ref imageHeightEMU);
            var picture = GenerateMainDocumentPart(System.IO.Path.GetFileNameWithoutExtension(imageFilePath), imageWidthEMU, imageHeightEMU, szRole);
            P.AppendChild(new Run(picture));
        }

        private void AddParts(HeaderPart parent, string imageFilePath, Paragraph P, string szRole)
        {
            long imageWidthEMU = 0;
            long imageHeightEMU = 0;
            var imagePart = parent.AddNewPart<ImagePart>("image/jpeg", "rId" + System.IO.Path.GetFileNameWithoutExtension(imageFilePath));
            GenerateImagePart(imagePart, imageFilePath,
             ref imageWidthEMU, ref imageHeightEMU);
            var picture = GenerateMainDocumentPart(System.IO.Path.GetFileNameWithoutExtension(imageFilePath), imageWidthEMU, imageHeightEMU, szRole);
            P.AppendChild(new Run(picture));
        }

        private void GenerateImagePart(OpenXmlPart part, string imageFilePath, ref long imageWidthEMU, ref long imageHeightEMU)
        {
            byte[] imageFileBytes;
            Bitmap imageFile;

            // Open a stream on the image file and read it's contents.
            using (FileStream fsImageFile = File.OpenRead(imageFilePath))
            {
                imageFileBytes = new byte[fsImageFile.Length];
                fsImageFile.Read(imageFileBytes, 0, imageFileBytes.Length);

                imageFile = new Bitmap(fsImageFile);
            }

            // Get the dimensions of the image in English Metric Units (EMU)
            // for use when adding the markup for the image to the document.
            imageWidthEMU =
              (long)(
              (imageFile.Width / imageFile.HorizontalResolution) * 914400L);
            imageHeightEMU =
              (long)(
              (imageFile.Height / imageFile.VerticalResolution) * 914400L);

            // Write the contents of the image to the ImagePart.
            using (BinaryWriter writer = new BinaryWriter(part.GetStream()))
            {
                writer.Write(imageFileBytes);
                writer.Flush();
            }
        }
        private void GenerateImagePart(OpenXmlPart part, Stream strmScanSign, ref long imageWidthEMU, ref long imageHeightEMU)
        {
            byte[] imageFileBytes;
            Bitmap imageFile;
            imageFileBytes = new byte[strmScanSign.Length];
            imageFile = new Bitmap(strmScanSign);

            // Get the dimensions of the image in English Metric Units (EMU)
            // for use when adding the markup for the image to the document.
            imageWidthEMU =
              (long)(
              (imageFile.Width / imageFile.HorizontalResolution) * 914400L);
            imageHeightEMU =
              (long)(
              (imageFile.Height / imageFile.VerticalResolution) * 914400L);

            // Write the contents of the image to the ImagePart.
            using (BinaryWriter writer = new BinaryWriter(part.GetStream()))
            {
                writer.Write(imageFileBytes);
                writer.Flush();
            }
        }

        private Drawing GenerateMainDocumentPart(string imageFileName, long imageWidthEMU, long imageHeightEMU, string szScanSignOfRole)
        {
            string GraphicDataUri =
              "http://schemas.openxmlformats.org/drawingml/2006/picture";

            double imageWidthInInches = imageWidthEMU / 914400.0;
            double imageHeightInInches = imageHeightEMU / 914400.0;

            long horizontalWrapPolygonUnitsPerInch =
              (long)(21600L / imageWidthInInches);

            long verticalWrapPolygonUnitsPerInch =
              (long)(21600L / imageHeightInInches);

            var element =
              new Drawing(
                new wp.Inline(

                  new wp.Extent()
                  {
                      Cx = imageWidthEMU,
                      Cy = imageHeightEMU
                  },

                  new wp.EffectExtent()
                  {
                      LeftEdge = 19050L,
                      TopEdge = 0L,
                      RightEdge = 9525L,
                      BottomEdge = 0L
                  },

                  new wp.DocProperties()
                  {
                      Id = (UInt32Value)1U,
                      Name = "ScanSignDrawing_" + imageFileName,
                      Description = szScanSignOfRole
                  },

                  new wp.NonVisualGraphicFrameDrawingProperties(
                    new a.GraphicFrameLocks() { NoChangeAspect = true }),

                  new a.Graphic(
                    new a.GraphicData(
                      new pic.Picture(

                        new pic.NonVisualPictureProperties(
                          new pic.NonVisualDrawingProperties()
                          {
                              Id = (UInt32Value)0U,
                              Name = imageFileName
                          },
                          new pic.NonVisualPictureDrawingProperties()),

                        new pic.BlipFill(
                          new a.Blip() { Embed = "rId" + imageFileName },
                          new a.Stretch(
                            new a.FillRectangle())),

                        new pic.ShapeProperties(
                          new a.Transform2D(
                            new a.Offset() { X = 0L, Y = 0L },
                            new a.Extents()
                            {
                                Cx = imageWidthEMU,
                                Cy = imageHeightEMU
                            }),

                          new a.PresetGeometry(
                            new a.AdjustValueList()
                          ) { Preset = a.ShapeTypeValues.Rectangle }))
                    ) { Uri = GraphicDataUri })
                )
                {
                    DistanceFromTop = (UInt32Value)0U,
                    DistanceFromBottom = (UInt32Value)0U,
                    DistanceFromLeft = (UInt32Value)0U,
                    DistanceFromRight = (UInt32Value)0U
                });
            return element;
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

            }
            else
            {
            }
        }

        ~ClsUpdate_ScanSign()
        {
            Dispose(false);
        }


        #endregion


    }
}
