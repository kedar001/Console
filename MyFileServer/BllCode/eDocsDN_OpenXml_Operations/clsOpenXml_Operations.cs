using System;
using System.Collections.Generic;
using System.Xml;
using System.IO.Packaging;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using V = DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Wordprocessing;
using Lock = DocumentFormat.OpenXml.Vml.Office.Lock;
using DocumentFormat.OpenXml;
using System.Drawing;
using wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using a = DocumentFormat.OpenXml.Drawing;
using pic = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using System.Security.Cryptography;
using System.Text;
using System.IO.Compression;






namespace eDocsDN_OpenXml_Operations
{
    public static class clsOpenXml_Operations
    {

        static internal XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        static internal XNamespace rel = "http://schemas.openxmlformats.org/package/2006/relationships";
        static internal XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        static internal XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        static internal XNamespace customPropertiesSchema = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
        static internal XNamespace customVTypesSchema = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
        static internal XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        //static internal XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        static internal XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        static internal XNamespace v = "urn:schemas-microsoft-com:vml";
        internal static XNamespace n = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";


        #region .... Variable Declaration ....
        static bool _bResult;
        static Dictionary<string, string> dicScanSign;

        #endregion

        #region ... Property ....

        public static string msgError { get; set; }

        #endregion

        #region .... Enum Propery Type
        public enum PropertyTypes : int
        {
            YesNo,
            Text,
            DateTime,
            NumberInteger,
            NumberDouble
        }

        public enum LockType
        {
            ReadOnly, None, Comments, TrackedChanges, Forms
        }

        #endregion

        #region .... Public Methods ...

        public static bool Check_Custom_Properties(string szFilePath, string szDCR_Number)
        {
            bool bResult = true;
            try
            {
                using (var document = WordprocessingDocument.Open(szFilePath, false))
                {
                    var customProps = document.CustomFilePropertiesPart;
                    var props = customProps.Properties;
                    if (props != null)
                    {

                        var prop =
                            props.Where(
                            p => ((CustomDocumentProperty)p).Name.Value.Equals("ARF")).FirstOrDefault();

                        if (prop.InnerText.Equals(szDCR_Number))
                            bResult = true;
                        else
                            bResult = false;
                    }
                }
            }
            catch (Exception ex)
            {
                bResult = false;
                msgError = ex.Message;
            }
            return bResult;
        }

        public static bool SetCustomProperty(string szFileName, string szPropertyName, object oPropertyValue, PropertyTypes propertyType)
        {
            msgError = "";
            const string documentRelationshipType =
                "http://schemas.openxmlformats.org/officeDocument/" +
                "2006/relationships/officeDocument";
            const string customPropertiesRelationshipType =
                "http://schemas.openxmlformats.org/officeDocument/" +
                "2006/relationships/custom-properties";
            const string customPropertiesSchema =
                "http://schemas.openxmlformats.org/officeDocument/" +
                "2006/custom-properties";
            const string customVTypesSchema =
                "http://schemas.openxmlformats.org/officeDocument/" +
                "2006/docPropsVTypes";

            bool retVal = false;
            PackagePart documentPart = null;
            string propertyTypeName = "vt:lpwstr";
            string propertyValueString = null;
            try
            {
                //  Calculate the correct type.
                switch (propertyType)
                {
                    case PropertyTypes.DateTime:
                        propertyTypeName = "vt:filetime";
                        if (oPropertyValue.GetType() == typeof(System.DateTime))
                        {
                            propertyValueString = string.Format("{0:s}Z",
                              Convert.ToDateTime(oPropertyValue));
                        }
                        break;

                    case PropertyTypes.NumberInteger:
                        propertyTypeName = "vt:i4";
                        if (oPropertyValue.GetType() == typeof(System.Int32))
                        {
                            propertyValueString =
                              Convert.ToInt32(oPropertyValue).ToString();
                        }
                        break;

                    case PropertyTypes.NumberDouble:
                        propertyTypeName = "vt:r8";
                        if (oPropertyValue.GetType() == typeof(System.Double))
                        {
                            propertyValueString =
                              Convert.ToDouble(oPropertyValue).ToString();
                        }
                        break;

                    case PropertyTypes.Text:
                        propertyTypeName = "vt:lpwstr";
                        propertyValueString = Convert.ToString(oPropertyValue);
                        break;

                    case PropertyTypes.YesNo:
                        propertyTypeName = "vt:bool";
                        if (oPropertyValue.GetType() == typeof(System.Boolean))
                        {
                            //  Must be lower case!
                            propertyValueString =
                              Convert.ToBoolean(oPropertyValue).ToString().ToLower();
                        }
                        break;
                }

                if (propertyValueString == null)
                    throw new InvalidDataException("Invalid parameter value.");

                using (Package wdPackage = Package.Open(
                       szFileName, FileMode.Open, FileAccess.ReadWrite))
                {
                    //  Get the main document part (document.xml).
                    foreach (PackageRelationship relationship in
                    wdPackage.GetRelationshipsByType(documentRelationshipType))
                    {
                        Uri documentUri = PackUriHelper.ResolvePartUri(
                            new Uri("/", UriKind.Relative), relationship.TargetUri);
                        documentPart = wdPackage.GetPart(documentUri);
                        //  There is only one document.
                        break;
                    }

                    //  Work with the custom properties part.
                    PackagePart customPropsPart = null;

                    //  Get the custom part (custom.xml). It may not exist.
                    foreach (PackageRelationship relationship in
                      wdPackage.GetRelationshipsByType(
                      customPropertiesRelationshipType))
                    {
                        Uri documentUri = PackUriHelper.ResolvePartUri(
                            new Uri("/", UriKind.Relative), relationship.TargetUri);
                        customPropsPart = wdPackage.GetPart(documentUri);
                        //  There is only one custom properties part, 
                        // if it exists at all.
                        break;
                    }

                    //  Manage namespaces to perform Xml XPath queries.
                    NameTable nt = new NameTable();
                    XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
                    nsManager.AddNamespace("d", customPropertiesSchema);
                    nsManager.AddNamespace("vt", customVTypesSchema);

                    Uri customPropsUri =
                      new Uri("/docProps/custom.xml", UriKind.Relative);
                    XmlDocument customPropsDoc = null;
                    XmlNode rootNode = null;

                    if (customPropsPart == null)
                    {
                        customPropsDoc = new XmlDocument(nt);

                        //  The part does not exist. Create it now.
                        customPropsPart = wdPackage.CreatePart(
                          customPropsUri,
                          "application/vnd.openxmlformats-officedocument.custom-properties+xml");

                        //  Set up the rudimentary custom part.
                        rootNode = customPropsDoc.
                          CreateElement("Properties", customPropertiesSchema);
                        rootNode.Attributes.Append(
                          customPropsDoc.CreateAttribute("xmlns:vt"));
                        rootNode.Attributes["xmlns:vt"].Value = customVTypesSchema;

                        customPropsDoc.AppendChild(rootNode);
                        wdPackage.CreateRelationship(customPropsUri,
                          TargetMode.Internal, customPropertiesRelationshipType);
                    }
                    else
                    {
                        customPropsDoc = new XmlDocument(nt);
                        customPropsDoc.Load(customPropsPart.GetStream());
                        rootNode = customPropsDoc.DocumentElement;
                    }

                    string searchString =
                      string.Format("d:Properties/d:property[@name='{0}']",
                      szPropertyName);
                    XmlNode node = customPropsDoc.SelectSingleNode(
                      searchString, nsManager);

                    XmlNode valueNode = null;

                    if (node != null)
                    {
                        //  You found the node. Now check its type.
                        if (node.HasChildNodes)
                        {
                            valueNode = node.ChildNodes[0];
                            if (valueNode != null)
                            {
                                string typeName = valueNode.Name;
                                if (propertyTypeName == typeName)
                                {
                                    valueNode.InnerText = propertyValueString;
                                    retVal = true;
                                }
                                else
                                {
                                    node.ParentNode.RemoveChild(node);
                                    node = null;
                                }
                            }
                        }
                    }

                    if (node == null)
                    {
                        string pidValue = "2";

                        XmlNode propertiesNode = customPropsDoc.DocumentElement;
                        if (propertiesNode.HasChildNodes)
                        {
                            XmlNode lastNode = propertiesNode.LastChild;
                            if (lastNode != null)
                            {
                                XmlAttribute pidAttr = lastNode.Attributes["pid"];
                                if (!(pidAttr == null))
                                {
                                    pidValue = pidAttr.Value;
                                    int value = 0;
                                    if (int.TryParse(pidValue, out value))
                                    {
                                        pidValue = Convert.ToString(value + 1);
                                    }
                                }
                            }
                        }

                        node = customPropsDoc.CreateElement("property", customPropertiesSchema);
                        node.Attributes.Append(customPropsDoc.CreateAttribute("name"));
                        node.Attributes["name"].Value = szPropertyName;

                        node.Attributes.Append(customPropsDoc.CreateAttribute("fmtid"));
                        node.Attributes["fmtid"].Value = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

                        node.Attributes.Append(customPropsDoc.CreateAttribute("pid"));
                        node.Attributes["pid"].Value = pidValue;

                        valueNode = customPropsDoc.
                        CreateElement(propertyTypeName, customVTypesSchema);
                        valueNode.InnerText = propertyValueString;
                        node.AppendChild(valueNode);
                        rootNode.AppendChild(node);
                        retVal = true;
                    }

                    //  Save the properties XML back to its part.
                    customPropsDoc.Save(customPropsPart.GetStream(
                                        FileMode.Create, FileAccess.Write));

                }
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            return retVal;
        }

        public static string UpdateCustomProperty(string fileName, string propertyName, object propertyValue, PropertyTypes propertyType)
        {
            string returnValue = null;
            var newProp = new CustomDocumentProperty();
            bool propSet = false;
            msgError = "";
            try
            {
                // Calculate the correct type.
                switch (propertyType)
                {
                    case PropertyTypes.DateTime:

                        if ((propertyValue) is DateTime)
                        {
                            newProp.VTFileTime =
                                new VTFileTime(string.Format("{0:s}Z",
                                    Convert.ToDateTime(propertyValue)));
                            propSet = true;
                        }

                        break;

                    case PropertyTypes.NumberInteger:
                        if ((propertyValue) is int)
                        {
                            newProp.VTInt32 = new VTInt32(propertyValue.ToString());
                            propSet = true;
                        }

                        break;

                    case PropertyTypes.NumberDouble:
                        if (propertyValue is double)
                        {
                            newProp.VTFloat = new VTFloat(propertyValue.ToString());
                            propSet = true;
                        }

                        break;

                    case PropertyTypes.Text:
                        newProp.VTLPWSTR = new VTLPWSTR(propertyValue.ToString());
                        propSet = true;

                        break;

                    case PropertyTypes.YesNo:
                        if (propertyValue is bool)
                        {
                            // Must be lowercase.
                            newProp.VTBool = new VTBool(
                              Convert.ToBoolean(propertyValue).ToString().ToLower());
                            propSet = true;
                        }
                        break;
                }

                if (!propSet)
                {
                    throw new InvalidDataException("propertyValue");
                }

                newProp.FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
                newProp.Name = propertyName;

                using (var document = WordprocessingDocument.Open(fileName, true))
                {
                    var customProps = document.CustomFilePropertiesPart;
                    if (customProps == null)
                    {
                        customProps = document.AddCustomFilePropertiesPart();
                        customProps.Properties =
                            new DocumentFormat.OpenXml.CustomProperties.Properties();
                    }

                    var props = customProps.Properties;
                    if (props != null)
                    {

                        var prop =
                            props.Where(
                            p => ((CustomDocumentProperty)p).Name.Value
                                == propertyName).FirstOrDefault();

                        // Does the property exist? If so, get the return value, 
                        // and then delete the property.
                        if (prop != null)
                        {
                            returnValue = prop.InnerText;
                            prop.Remove();
                        }

                        props.AppendChild(newProp);
                        int pid = 2;
                        foreach (CustomDocumentProperty item in props)
                        {
                            item.PropertyId = pid++;
                        }
                        props.Save();
                    }

                }
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            return returnValue;
        }

        public static bool updateFields(string szFileName)
        {
            WordprocessingDocument _objDoc;
            _bResult = true;
            msgError = "";
            try
            {
                using (_objDoc = WordprocessingDocument.Open(szFileName, true))
                {
                    DocumentSettingsPart settingsPart = _objDoc.MainDocumentPart.GetPartsOfType<DocumentSettingsPart>().First();
                    UpdateFieldsOnOpen updateFields = new UpdateFieldsOnOpen();
                    updateFields.Val = new DocumentFormat.OpenXml.OnOffValue(true);
                    settingsPart.Settings.PrependChild<UpdateFieldsOnOpen>(updateFields);
                    settingsPart.Settings.Save();
                }
            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = ex.Message;
            }
            finally
            {
                _objDoc = null;
            }
            return _bResult;

        }

        public static bool InsertEditRemoveWatermark(string szFilePath, string watermarkText)
        {

            _bResult = true;
            WordprocessingDocument _objDoc;
            msgError = "";
            try
            {
                if (watermarkText != "")
                    InsertEditRemoveWatermark(szFilePath, "");
                using (_objDoc = WordprocessingDocument.Open(szFilePath, true))
                {
                    if (_objDoc.MainDocumentPart.HeaderParts != null)
                    {
                        foreach (var header in _objDoc.MainDocumentPart.HeaderParts)
                        {

                            if (string.IsNullOrEmpty(watermarkText))
                            {
                                //Remove
                                if (header.Header.Descendants<Paragraph>() != null)
                                {
                                    foreach (var para in header.Header.Descendants<Paragraph>())
                                    {
                                        foreach (Run r in para.Descendants<Run>())
                                        {
                                            FindRemoveWatermark(r);
                                        }
                                    }
                                }
                            }
                            else
                            {
                                //Insert or change
                                if (header.Header.Descendants<Paragraph>() == null)
                                {
                                    //No paragraph exists in the header
                                    //Insert run
                                    Run r = CreateWatermarkRun(watermarkText);
                                    Paragraph para = new Paragraph();
                                    para.Append(r);
                                }
                                else
                                {
                                    //Loop over all paragraphs
                                    bool test = false;
                                    foreach (var para in header.Header.Descendants<Paragraph>())
                                    {

                                        foreach (Run r in para.Descendants<Run>())
                                        {
                                            test = FindReplaceWatermarkText(r, watermarkText);
                                            break;
                                        }
                                    }
                                    if (!test)
                                    {
                                        Run r = CreateWatermarkRun(watermarkText);
                                        Paragraph fp = header.Header.Descendants<Paragraph>().LastOrDefault();
                                        fp.Append(r);
                                    }
                                }
                            }
                            header.Header.Save(header);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = ex.Message;
            }
            finally
            {
                _objDoc = null;
            }
            return _bResult;
        }

        public static bool TrackRevisions(string szFileName, bool bTrackRevisions)
        {

            _bResult = true;
            string szPassword = "Espl123&;";
            WordprocessingDocument objDoc = null;
            try
            {

                objDoc = WordprocessingDocument.Open(szFileName, true);
                TrackRevisions newrevision = new TrackRevisions();
                newrevision.Val = new DocumentFormat.OpenXml.OnOffValue(bTrackRevisions);
                objDoc.MainDocumentPart.DocumentSettingsPart.Settings.RemoveAllChildren<TrackRevisions>();
                objDoc.MainDocumentPart.DocumentSettingsPart.Settings.AppendChild(newrevision);
                objDoc.MainDocumentPart.DocumentSettingsPart.Settings.Save();
                if (isDocumentisEditable(objDoc))
                {
                    var documentSettings = objDoc.MainDocumentPart.DocumentSettingsPart;
                    var documentProtection = documentSettings
                                                .Settings
                                                .FirstOrDefault(it => it is DocumentProtection) as DocumentProtection;

                    #region ...... Password Encryption .....

                    byte[] arrSalt = new byte[16];
                    RandomNumberGenerator rand = new RNGCryptoServiceProvider();
                    rand.GetNonZeroBytes(arrSalt);
                    byte[] generatedKey = new byte[4];

                    int intMaxPasswordLength = 15;
                    if (!String.IsNullOrEmpty(szPassword))
                    {
                        szPassword = szPassword.Substring(0, Math.Min(szPassword.Length, intMaxPasswordLength));
                        byte[] arrByteChars = new byte[szPassword.Length];

                        for (int intLoop = 0; intLoop < szPassword.Length; intLoop++)
                        {
                            int intTemp = Convert.ToInt32(szPassword[intLoop]);
                            arrByteChars[intLoop] = Convert.ToByte(intTemp & 0x00FF);
                            if (arrByteChars[intLoop] == 0)
                                arrByteChars[intLoop] = Convert.ToByte((intTemp & 0xFF00) >> 8);
                        }
                        int intHighOrderWord = InitialCodeArray[arrByteChars.Length - 1];
                        for (int intLoop = 0; intLoop < arrByteChars.Length; intLoop++)
                        {
                            int tmp = intMaxPasswordLength - arrByteChars.Length + intLoop;
                            for (int intBit = 0; intBit < 7; intBit++)
                            {
                                if ((arrByteChars[intLoop] & (0x0001 << intBit)) != 0)
                                {
                                    intHighOrderWord ^= EncryptionMatrix[tmp, intBit];
                                }
                            }
                        }
                        int intLowOrderWord = 0;
                        for (int intLoopChar = arrByteChars.Length - 1; intLoopChar >= 0; intLoopChar--)
                        {
                            intLowOrderWord = (((intLowOrderWord >> 14) & 0x0001) | ((intLowOrderWord << 1) & 0x7FFF)) ^ arrByteChars[intLoopChar];
                        }
                        intLowOrderWord = (((intLowOrderWord >> 14) & 0x0001) | ((intLowOrderWord << 1) & 0x7FFF)) ^ arrByteChars.Length ^ 0xCE4B;
                        int intCombinedkey = (intHighOrderWord << 16) + intLowOrderWord;
                        for (int intTemp = 0; intTemp < 4; intTemp++)
                        {
                            generatedKey[intTemp] = Convert.ToByte(((uint)(intCombinedkey & (0x000000FF << (intTemp * 8)))) >> (intTemp * 8));
                        }
                    }
                    StringBuilder sb = new StringBuilder();
                    for (int intTemp = 0; intTemp < 4; intTemp++)
                    {
                        sb.Append(Convert.ToString(generatedKey[intTemp], 16));
                    }
                    generatedKey = Encoding.Unicode.GetBytes(sb.ToString().ToUpper());
                    byte[] tmpArray1 = generatedKey;
                    byte[] tmpArray2 = arrSalt;
                    byte[] tempKey = new byte[tmpArray1.Length + tmpArray2.Length];
                    Buffer.BlockCopy(tmpArray2, 0, tempKey, 0, tmpArray2.Length);
                    Buffer.BlockCopy(tmpArray1, 0, tempKey, tmpArray2.Length, tmpArray1.Length);
                    generatedKey = tempKey;
                    int iterations = 50000;
                    HashAlgorithm sha1 = new SHA1Managed();
                    generatedKey = sha1.ComputeHash(generatedKey);
                    byte[] iterator = new byte[4];
                    for (int intTmp = 0; intTmp < iterations; intTmp++)
                    {

                        //When iterating on the hash, you are supposed to append the current iteration number.
                        iterator[0] = Convert.ToByte((intTmp & 0x000000FF) >> 0);
                        iterator[1] = Convert.ToByte((intTmp & 0x0000FF00) >> 8);
                        iterator[2] = Convert.ToByte((intTmp & 0x00FF0000) >> 16);
                        iterator[3] = Convert.ToByte((intTmp & 0xFF000000) >> 24);

                        generatedKey = concatByteArrays(iterator, generatedKey);
                        generatedKey = sha1.ComputeHash(generatedKey);
                    }
                    #endregion

                    DocumentFormat.OpenXml.OnOffValue docProtection = new DocumentFormat.OpenXml.OnOffValue(true);
                    if (documentProtection == null)
                    {
                        var documentProtectionElement = new DocumentProtection();
                        documentProtectionElement.Edit = DocumentProtectionValues.TrackedChanges;
                        documentProtectionElement.Enforcement = OnOffValue.FromBoolean(bTrackRevisions);
                        documentProtectionElement.CryptographicAlgorithmClass = CryptAlgorithmClassValues.Hash;
                        documentProtectionElement.CryptographicProviderType = CryptProviderValues.RsaFull;
                        documentProtectionElement.CryptographicAlgorithmType = CryptAlgorithmValues.TypeAny;
                        documentProtectionElement.CryptographicAlgorithmSid = 4; // SHA1
                        //    The iteration count is unsigned
                        UInt32 uintVal = new UInt32();
                        uintVal = (uint)iterations;
                        documentProtectionElement.CryptographicSpinCount = uintVal;
                        documentProtectionElement.Hash = Convert.ToBase64String(generatedKey);
                        documentProtectionElement.Salt = Convert.ToBase64String(arrSalt);
                        documentSettings.Settings.AppendChild(documentProtectionElement);
                        objDoc.MainDocumentPart.DocumentSettingsPart.Settings.AppendChild(documentProtection);
                    }
                    else
                    {
                        documentSettings.Settings.RemoveAllChildren<DocumentProtection>();
                        documentProtection.Enforcement = OnOffValue.FromBoolean(bTrackRevisions);
                        documentProtection.Edit = DocumentProtectionValues.TrackedChanges;
                        documentProtection.Enforcement = OnOffValue.FromBoolean(bTrackRevisions);
                        documentProtection.CryptographicAlgorithmClass = CryptAlgorithmClassValues.Hash;
                        documentProtection.CryptographicProviderType = CryptProviderValues.RsaFull;
                        documentProtection.CryptographicAlgorithmType = CryptAlgorithmValues.TypeAny;
                        documentProtection.CryptographicAlgorithmSid = 4; // SHA1
                        //    The iteration count is unsigned
                        UInt32 uintVal = new UInt32();
                        uintVal = (uint)iterations;
                        documentProtection.CryptographicSpinCount = uintVal;
                        documentProtection.Hash = Convert.ToBase64String(generatedKey);
                        documentProtection.Salt = Convert.ToBase64String(arrSalt);
                        documentSettings.Settings.AppendChild(documentProtection);
                    }
                    objDoc.MainDocumentPart.DocumentSettingsPart.Settings.Save();
                }

                objDoc.Close();
                objDoc.Dispose();
                objDoc = null;
            }
            catch (OpenXmlPackageException e)
            {
                if (e.ToString().Contains("Invalid Hyperlink"))
                {
                    if (e.ToString().Contains("Invalid Hyperlink"))
                    {
                        using (FileStream fs = new FileStream(szFileName, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                        {
                            UriFixer.FixInvalidUri(fs, brokenUri => FixUri(brokenUri));
                        }
                        TrackRevisions(szFileName, bTrackRevisions);
                    }
                }
            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = ex.Message;
            }
            finally
            {
                objDoc = null;
            }
            return _bResult;
        }
        public static Stream TrackRevisions(Stream strmDocument, bool bTrackRevisions)
        {
            _bResult = true;
            string szPassword = "Espl123&;";
            WordprocessingDocument objDoc = null;
            msgError = "";
            string szTempPath = System.Windows.Forms.Application.StartupPath + "\\temp.docx";

            try
            {

                objDoc = WordprocessingDocument.Open(strmDocument, true);
                TrackRevisions newrevision = new TrackRevisions();
                newrevision.Val = new DocumentFormat.OpenXml.OnOffValue(bTrackRevisions);
                objDoc.MainDocumentPart.DocumentSettingsPart.Settings.RemoveAllChildren<TrackRevisions>();
                objDoc.MainDocumentPart.DocumentSettingsPart.Settings.AppendChild(newrevision);
                objDoc.MainDocumentPart.DocumentSettingsPart.Settings.Save();
                if (isDocumentisEditable(objDoc))
                {
                    var documentSettings = objDoc.MainDocumentPart.DocumentSettingsPart;
                    var documentProtection = documentSettings
                                                .Settings
                                                .FirstOrDefault(it => it is DocumentProtection) as DocumentProtection;

                    #region ...... Password Encryption .....

                    byte[] arrSalt = new byte[16];
                    RandomNumberGenerator rand = new RNGCryptoServiceProvider();
                    rand.GetNonZeroBytes(arrSalt);
                    byte[] generatedKey = new byte[4];

                    int intMaxPasswordLength = 15;
                    if (!String.IsNullOrEmpty(szPassword))
                    {
                        szPassword = szPassword.Substring(0, Math.Min(szPassword.Length, intMaxPasswordLength));
                        byte[] arrByteChars = new byte[szPassword.Length];

                        for (int intLoop = 0; intLoop < szPassword.Length; intLoop++)
                        {
                            int intTemp = Convert.ToInt32(szPassword[intLoop]);
                            arrByteChars[intLoop] = Convert.ToByte(intTemp & 0x00FF);
                            if (arrByteChars[intLoop] == 0)
                                arrByteChars[intLoop] = Convert.ToByte((intTemp & 0xFF00) >> 8);
                        }
                        int intHighOrderWord = InitialCodeArray[arrByteChars.Length - 1];
                        for (int intLoop = 0; intLoop < arrByteChars.Length; intLoop++)
                        {
                            int tmp = intMaxPasswordLength - arrByteChars.Length + intLoop;
                            for (int intBit = 0; intBit < 7; intBit++)
                            {
                                if ((arrByteChars[intLoop] & (0x0001 << intBit)) != 0)
                                {
                                    intHighOrderWord ^= EncryptionMatrix[tmp, intBit];
                                }
                            }
                        }
                        int intLowOrderWord = 0;
                        for (int intLoopChar = arrByteChars.Length - 1; intLoopChar >= 0; intLoopChar--)
                        {
                            intLowOrderWord = (((intLowOrderWord >> 14) & 0x0001) | ((intLowOrderWord << 1) & 0x7FFF)) ^ arrByteChars[intLoopChar];
                        }
                        intLowOrderWord = (((intLowOrderWord >> 14) & 0x0001) | ((intLowOrderWord << 1) & 0x7FFF)) ^ arrByteChars.Length ^ 0xCE4B;
                        int intCombinedkey = (intHighOrderWord << 16) + intLowOrderWord;
                        for (int intTemp = 0; intTemp < 4; intTemp++)
                        {
                            generatedKey[intTemp] = Convert.ToByte(((uint)(intCombinedkey & (0x000000FF << (intTemp * 8)))) >> (intTemp * 8));
                        }
                    }
                    StringBuilder sb = new StringBuilder();
                    for (int intTemp = 0; intTemp < 4; intTemp++)
                    {
                        sb.Append(Convert.ToString(generatedKey[intTemp], 16));
                    }
                    generatedKey = Encoding.Unicode.GetBytes(sb.ToString().ToUpper());
                    byte[] tmpArray1 = generatedKey;
                    byte[] tmpArray2 = arrSalt;
                    byte[] tempKey = new byte[tmpArray1.Length + tmpArray2.Length];
                    Buffer.BlockCopy(tmpArray2, 0, tempKey, 0, tmpArray2.Length);
                    Buffer.BlockCopy(tmpArray1, 0, tempKey, tmpArray2.Length, tmpArray1.Length);
                    generatedKey = tempKey;
                    int iterations = 50000;
                    HashAlgorithm sha1 = new SHA1Managed();
                    generatedKey = sha1.ComputeHash(generatedKey);
                    byte[] iterator = new byte[4];
                    for (int intTmp = 0; intTmp < iterations; intTmp++)
                    {

                        //When iterating on the hash, you are supposed to append the current iteration number.
                        iterator[0] = Convert.ToByte((intTmp & 0x000000FF) >> 0);
                        iterator[1] = Convert.ToByte((intTmp & 0x0000FF00) >> 8);
                        iterator[2] = Convert.ToByte((intTmp & 0x00FF0000) >> 16);
                        iterator[3] = Convert.ToByte((intTmp & 0xFF000000) >> 24);

                        generatedKey = concatByteArrays(iterator, generatedKey);
                        generatedKey = sha1.ComputeHash(generatedKey);
                    }
                    #endregion

                    DocumentFormat.OpenXml.OnOffValue docProtection = new DocumentFormat.OpenXml.OnOffValue(true);
                    if (documentProtection == null)
                    {
                        var documentProtectionElement = new DocumentProtection();
                        documentProtectionElement.Edit = DocumentProtectionValues.TrackedChanges;
                        documentProtectionElement.Enforcement = OnOffValue.FromBoolean(bTrackRevisions);
                        documentProtectionElement.CryptographicAlgorithmClass = CryptAlgorithmClassValues.Hash;
                        documentProtectionElement.CryptographicProviderType = CryptProviderValues.RsaFull;
                        documentProtectionElement.CryptographicAlgorithmType = CryptAlgorithmValues.TypeAny;
                        documentProtectionElement.CryptographicAlgorithmSid = 4; // SHA1
                        //    The iteration count is unsigned
                        UInt32 uintVal = new UInt32();
                        uintVal = (uint)iterations;
                        documentProtectionElement.CryptographicSpinCount = uintVal;
                        documentProtectionElement.Hash = Convert.ToBase64String(generatedKey);
                        documentProtectionElement.Salt = Convert.ToBase64String(arrSalt);
                        documentSettings.Settings.AppendChild(documentProtectionElement);
                        objDoc.MainDocumentPart.DocumentSettingsPart.Settings.AppendChild(documentProtection);
                    }
                    else
                    {
                        documentSettings.Settings.RemoveAllChildren<DocumentProtection>();
                        documentProtection.Enforcement = OnOffValue.FromBoolean(bTrackRevisions);
                        documentProtection.Edit = DocumentProtectionValues.TrackedChanges;
                        documentProtection.Enforcement = OnOffValue.FromBoolean(bTrackRevisions);
                        documentProtection.CryptographicAlgorithmClass = CryptAlgorithmClassValues.Hash;
                        documentProtection.CryptographicProviderType = CryptProviderValues.RsaFull;
                        documentProtection.CryptographicAlgorithmType = CryptAlgorithmValues.TypeAny;
                        documentProtection.CryptographicAlgorithmSid = 4; // SHA1
                        //    The iteration count is unsigned
                        UInt32 uintVal = new UInt32();
                        uintVal = (uint)iterations;
                        documentProtection.CryptographicSpinCount = uintVal;
                        documentProtection.Hash = Convert.ToBase64String(generatedKey);
                        documentProtection.Salt = Convert.ToBase64String(arrSalt);
                        documentSettings.Settings.AppendChild(documentProtection);
                    }
                    objDoc.MainDocumentPart.DocumentSettingsPart.Settings.Save();
                }

                objDoc.Close();
                objDoc.Dispose();
                objDoc = null;


            }
            catch (OpenXmlPackageException e)
            {
                if (e.ToString().Contains("Invalid Hyperlink"))
                {
                    if (e.ToString().Contains("Invalid Hyperlink"))
                    {
                        File.WriteAllBytes(szTempPath, strmDocument.ReadAllBytes());
                        using (FileStream fileStream = new FileStream(szTempPath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                        {
                            UriFixer.FixInvalidUri(fileStream, brokenUri => FixUri(brokenUri));
                        }
                        strmDocument = Convert_Document_To_Stream(File.ReadAllBytes(szTempPath));
                        File.Delete(szTempPath);
                    }
                }
            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = ex.Message;
            }
            finally
            {
                if (objDoc != null)
                {
                    objDoc.Close();
                    objDoc.Dispose();
                }
                objDoc = null;
            }
            return strmDocument;
        }
        public static bool IsPasswordProtectedDocument(string _FileName)
        {
            bool _bResult = true;
            try
            {
                using (WordprocessingDocument wdDoc =
                    WordprocessingDocument.Open(_FileName, false))
                {

                    WriteProtection Wp = wdDoc.MainDocumentPart.DocumentSettingsPart.Settings.GetFirstChild<WriteProtection>();
                    if (Wp == null)
                        _bResult = false;
                    if (Wp != null)
                        if (Wp.CryptographicAlgorithmClass.HasValue)
                            _bResult = true;

                    if (!_bResult)
                    {
                        DocumentProtection dp = wdDoc.MainDocumentPart.DocumentSettingsPart.Settings.GetFirstChild<DocumentProtection>();
                        if (dp == null)
                            _bResult = false;
                        if (dp != null && (dp.Edit != null))
                            _bResult = true;
                    }

                }
            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = ex.Message;
            }
            finally
            { }
            return _bResult;
        }

        internal static Stream Convert_Document_To_Stream(byte[] arrDocument)
        {
            MemoryStream strmDocument = new MemoryStream();
            strmDocument.Write(arrDocument, 0, (int)arrDocument.Length);
            return strmDocument;
        }

        private static bool isDocumentisEditable(WordprocessingDocument wd)
        {

            bool bResult = false;
            try
            {

                DocumentProtection dp = wd.MainDocumentPart.DocumentSettingsPart.Settings.GetFirstChild<DocumentProtection>();
                if (dp == null)
                    bResult = true;
                if (dp != null && (dp.Edit != null) && (dp.Edit == DocumentProtectionValues.None || dp.Edit == DocumentProtectionValues.TrackedChanges))
                    bResult = true;


            }
            finally { }
            return bResult;
        }

        private static bool isDocumentPasswordProtected(WordprocessingDocument wd)
        {

            bool bResult = false;
            try
            {

                WriteProtection Wp = wd.MainDocumentPart.DocumentSettingsPart.Settings.GetFirstChild<WriteProtection>();
                if (wp == null)
                    bResult = false;



            }
            finally { }
            return bResult;
        }
     


        public static bool IsTrackChangesON(string szFileName)
        {
            WordprocessingDocument document = null;
            bool bIsTrackChanges = false;
            DocumentFormat.OpenXml.OpenXmlElement oList;
            msgError = "";
            try
            {
                #region .... Track Changes ON/OFF ....
                using (document = WordprocessingDocument.Open(szFileName, true))
                {
                    oList = document.MainDocumentPart.DocumentSettingsPart.Settings.ChildElements.Where(t => t is DocumentFormat.OpenXml.Wordprocessing.TrackRevisions).Last();
                    DocumentFormat.OpenXml.Wordprocessing.OnOffType onOfftype = (DocumentFormat.OpenXml.Wordprocessing.OnOffType)oList;
                    bIsTrackChanges = true;
                    if (onOfftype.Val != null)
                        if (onOfftype.Val == true)
                            bIsTrackChanges = true;
                        else
                            bIsTrackChanges = false;
                }
                #endregion
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
                bIsTrackChanges = false;
            }
            finally
            {
                document = null;
            }
            return bIsTrackChanges;
        }

        public static bool IsAccept_Revision_Pending(string szfileName)
        {
            _bResult = true;
            try
            {
                using (WordprocessingDocument wdDoc =
               WordprocessingDocument.Open(szfileName, true))
                {
                    Body body = wdDoc.MainDocumentPart.Document.Body;

                    // Handle the formatting changes.
                    List<OpenXmlElement> changes =
                       body.Descendants<ParagraphPropertiesChange>()
                   .Cast<OpenXmlElement>().ToList();
                    changes.AddRange(body.Descendants<Deleted>()
                       .Cast<OpenXmlElement>().ToList());
                    changes.AddRange(body.Descendants<DeletedRun>()
                        .Cast<OpenXmlElement>().ToList());
                    changes.AddRange(body.Descendants<DeletedMathControl>()
                        .Cast<OpenXmlElement>().ToList());
                    changes.AddRange(body.Descendants<Inserted>()
                       .Cast<OpenXmlElement>().ToList());
                    changes.AddRange(body.Descendants<InsertedRun>()
                        .Cast<OpenXmlElement>().ToList());
                    changes.AddRange(body.Descendants<InsertedMathControl>()
                        .Cast<OpenXmlElement>().ToList());
                    changes.AddRange(body.Descendants<Comments>()
                        .Cast<OpenXmlElement>().ToList());
                    if (changes.Count > 0) _bResult = true; else _bResult = false;
                }
            }
            catch (Exception ex)
            {
                _bResult = false;
            }
            finally
            {

            }
            return _bResult;
        }

        public static bool isComments_Exist(string szFileName)
        {
            _bResult = false;
            msgError = "";
            try
            {
                using (WordprocessingDocument wordDoc =
                WordprocessingDocument.Open(szFileName, false))
                {
                    WordprocessingCommentsPart commentsPart =
                        wordDoc.MainDocumentPart.WordprocessingCommentsPart;

                    if (commentsPart != null && commentsPart.Comments != null)
                    {
                        foreach (Comment comment in commentsPart.Comments.Elements<Comment>())
                        {
                            _bResult = true;
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = ex.Message;
            }
            finally
            {

            }
            return _bResult;

        }

        public static Stream PrintFormsData(Stream strmDocument, bool bPrintFormData)
        {
            _bResult = true;
            WordprocessingDocument docPart = null;
            msgError = "";
            try
            {
                docPart = WordprocessingDocument.Open(strmDocument, true);
                PrintFormsData objPrintFormData = new PrintFormsData();
                objPrintFormData.Val = new DocumentFormat.OpenXml.OnOffValue(bPrintFormData);
                docPart.MainDocumentPart.DocumentSettingsPart.Settings.AppendChild(objPrintFormData);
                docPart.MainDocumentPart.DocumentSettingsPart.Settings.Save();

            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = ex.Message;
            }
            finally
            {
                if (docPart != null)
                {
                    docPart.Close();
                    docPart.Dispose();
                }
                docPart = null;
            }
            return strmDocument;
        }
        public static void PrintFormsData(string szFilePath, bool bPrintFormData)
        {
            _bResult = true;
            WordprocessingDocument docPart = null;
            msgError = "";
            try
            {
                docPart = WordprocessingDocument.Open(szFilePath, true);
                PrintFormsData objPrintFormData = new PrintFormsData();
                objPrintFormData.Val = new DocumentFormat.OpenXml.OnOffValue(bPrintFormData);
                docPart.MainDocumentPart.DocumentSettingsPart.Settings.AppendChild(objPrintFormData);
                docPart.MainDocumentPart.DocumentSettingsPart.Settings.Save();

            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = ex.Message;
            }
            finally
            {
                if (docPart != null)
                {
                    docPart.Close();
                    docPart.Dispose();
                }
                docPart = null;
            }
        }

        public static bool DeleteAllCommentsFromDocument(string szFileName)
        {
            _bResult = true;
            msgError = "";
            try
            {
                // Get an existing Wordprocessing document.
                using (WordprocessingDocument document =
                    WordprocessingDocument.Open(szFileName, true))
                {
                    // Set commentPart to the document WordprocessingCommentsPart, 
                    // if it exists.
                    WordprocessingCommentsPart commentPart =
                        document.MainDocumentPart.WordprocessingCommentsPart;

                    // If no WordprocessingCommentsPart exists, there can be no 
                    // comments. Stop execution and return from the method.
                    if (commentPart == null)
                    {
                        return true;
                    }

                    // Create a list of comments by the specified author, or
                    // if the author name is empty, all authors.
                    List<Comment> commentsToDelete =
                        commentPart.Comments.Elements<Comment>().ToList();
                    IEnumerable<string> commentIds =
                        commentsToDelete.Select(r => r.Id.Value);

                    // Delete each comment in commentToDelete from the 
                    // Comments collection.
                    foreach (Comment c in commentsToDelete)
                    {
                        c.Remove();
                    }

                    // Save the comment part change.
                    commentPart.Comments.Save();

                    Document doc = document.MainDocumentPart.Document;
                    List<CommentRangeStart> commentRangeStartToDelete =
                        doc.Descendants<CommentRangeStart>().
                        Where(c => commentIds.Contains(c.Id.Value)).ToList();
                    foreach (CommentRangeStart c in commentRangeStartToDelete)
                    {
                        c.Remove();
                    }

                    // Delete CommentRangeEnd for each deleted comment in the main document.
                    List<CommentRangeEnd> commentRangeEndToDelete =
                        doc.Descendants<CommentRangeEnd>().
                        Where(c => commentIds.Contains(c.Id.Value)).ToList();
                    foreach (CommentRangeEnd c in commentRangeEndToDelete)
                    {
                        c.Remove();
                    }

                    // Delete CommentReference for each deleted comment in the main document.
                    List<CommentReference> commentRangeReferenceToDelete =
                        doc.Descendants<CommentReference>().
                        Where(c => commentIds.Contains(c.Id.Value)).ToList();
                    foreach (CommentReference c in commentRangeReferenceToDelete)
                    {
                        c.Remove();
                    }

                    // Save changes back to the MainDocumentPart part.
                    doc.Save();
                }

            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = ex.Message;
            }
            finally
            {
            }
            return _bResult;

        }

        public static bool DeleteAllCommentsFromDocument(Stream strmDocument)
        {
            _bResult = true;
            msgError = "";
            try
            {
                // Get an existing Wordprocessing document.
                using (WordprocessingDocument document =
                    WordprocessingDocument.Open(strmDocument, true))
                {
                    // Set commentPart to the document WordprocessingCommentsPart, 
                    // if it exists.
                    WordprocessingCommentsPart commentPart =
                        document.MainDocumentPart.WordprocessingCommentsPart;

                    // If no WordprocessingCommentsPart exists, there can be no 
                    // comments. Stop execution and return from the method.
                    if (commentPart == null)
                    {
                        return true;
                    }

                    // Create a list of comments by the specified author, or
                    // if the author name is empty, all authors.
                    List<Comment> commentsToDelete =
                        commentPart.Comments.Elements<Comment>().ToList();
                    IEnumerable<string> commentIds =
                        commentsToDelete.Select(r => r.Id.Value);

                    // Delete each comment in commentToDelete from the 
                    // Comments collection.
                    foreach (Comment c in commentsToDelete)
                    {
                        c.Remove();
                    }

                    // Save the comment part change.
                    commentPart.Comments.Save();

                    Document doc = document.MainDocumentPart.Document;
                    List<CommentRangeStart> commentRangeStartToDelete =
                        doc.Descendants<CommentRangeStart>().
                        Where(c => commentIds.Contains(c.Id.Value)).ToList();
                    foreach (CommentRangeStart c in commentRangeStartToDelete)
                    {
                        c.Remove();
                    }

                    // Delete CommentRangeEnd for each deleted comment in the main document.
                    List<CommentRangeEnd> commentRangeEndToDelete =
                        doc.Descendants<CommentRangeEnd>().
                        Where(c => commentIds.Contains(c.Id.Value)).ToList();
                    foreach (CommentRangeEnd c in commentRangeEndToDelete)
                    {
                        c.Remove();
                    }

                    // Delete CommentReference for each deleted comment in the main document.
                    List<CommentReference> commentRangeReferenceToDelete =
                        doc.Descendants<CommentReference>().
                        Where(c => commentIds.Contains(c.Id.Value)).ToList();
                    foreach (CommentReference c in commentRangeReferenceToDelete)
                    {
                        c.Remove();
                    }

                    // Save changes back to the MainDocumentPart part.
                    doc.Save();
                }

            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = ex.Message;
            }
            finally
            {
            }
            return _bResult;

        }

        public static bool AcceptRevisions(string fileName)
        {
            _bResult = true;
            msgError = "";
            WordprocessingDocument _objDoc = null;
            try
            {
                using (_objDoc = WordprocessingDocument.Open(fileName, true))
                {
                    AcceptAll(_objDoc);
                }

                #region .... Accpt Changes  ....
                DeleteComments(fileName, "");
                using (_objDoc = WordprocessingDocument.Open(fileName, true))
                {
                    Body body = _objDoc.MainDocumentPart.Document.Body;
                    List<OpenXmlElement> changes =
                        body.Descendants<ParagraphPropertiesChange>()
                    .Cast<OpenXmlElement>().ToList();
                    foreach (OpenXmlElement change in changes)
                    {
                        change.Remove();
                    }
                    List<OpenXmlElement> deletions =
                        body.Descendants<Deleted>()
                        .Cast<OpenXmlElement>().ToList();
                    deletions.AddRange(body.Descendants<DeletedRun>()
                        .Cast<OpenXmlElement>().ToList());
                    deletions.AddRange(body.Descendants<DeletedMathControl>()
                        .Cast<OpenXmlElement>().ToList());
                    foreach (OpenXmlElement deletion in deletions)
                    {
                        deletion.Remove();
                    }
                    List<OpenXmlElement> insertions =
                        body.Descendants<Inserted>()
                        .Cast<OpenXmlElement>().ToList();
                    insertions.AddRange(body.Descendants<InsertedRun>()
                        .Cast<OpenXmlElement>().ToList());
                    insertions.AddRange(body.Descendants<InsertedMathControl>()
                        .Cast<OpenXmlElement>().ToList());
                    foreach (OpenXmlElement insertion in insertions)
                    {
                        foreach (var run in insertion.Elements<Run>())
                        {
                            if (run == insertion.FirstChild)
                            {
                                insertion.InsertAfterSelf(new Run(run.OuterXml));
                            }
                            else
                            {
                                insertion.NextSibling().InsertAfterSelf(new Run(run.OuterXml));
                            }
                        }
                        insertion.RemoveAttribute("rsidR",
                            "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                        insertion.RemoveAttribute("rsidRPr",
                            "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                        insertion.Remove();
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = ex.Message;
            }
            finally
            {
                _objDoc = null;
            }
            return _bResult;
        }

        public static bool UpdateScanSign(string szFileName, Dictionary<string, string> szScanSignOfUser, bool bRemoveScanSign)
        {
            _bResult = true;
            msgError = "";
            eDocsDN_DocX.Image objImage;
            try
            {
                //foreach (var item in szScanSignOfUser)
                //{
                //    using (eDocsDN_DocX.DocX objDoc = eDocsDN_DocX.DocX.Load(szFileName))
                //    {
                //        objImage = objDoc.AddImage(item.Value);
                //    }
                //    ScanSign(szFileName, item.Key, item.Value, bRemoveScanSign);
                //}
                ScanSign(szFileName, szScanSignOfUser, bRemoveScanSign);
            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = ex.Message;
            }
            finally
            {

            }
            return _bResult;
        }

        public static void RemoveScanSign(string szFileName)
        {
            string szScansignFilePath = string.Empty;
            Run rScanSignToRemove = null;
            string szScanSignCustumProperty = string.Empty;
            string szName = string.Empty;
            string szDescription = string.Empty;

            try
            {

                using (eDocsDN_DocX.DocX objDoc = eDocsDN_DocX.DocX.Load(szFileName))
                {
                    List<XElement> xEleDocument = objDoc.Get_document_xElement(objDoc);
                    foreach (XElement doc in xEleDocument)
                    {
                        foreach (XElement e in doc.Descendants(XName.Get("drawing", w.NamespaceName)))
                        {
                            //..Get Element 
                            var szScanSignId =
                             (
                                 from d in e.Descendants()
                                 where d.Name.LocalName.Contains("ScanSignDrawing_")
                                 select d
                             ).SingleOrDefault();

                            szName = szScanSignId.Attribute(XName.Get("name")).Value;
                            szScanSignCustumProperty = szScanSignId.Attribute(XName.Get("descr")).Value;

                            //..Remove Image
                            e.Parent.RemoveAll();
                            //..
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
                            //..
                            e.Parent.Add(new Run(simpleField1).InnerXml);

                        }
                    }
                }


                //DR NO:919833
                //using (WordprocessingDocument doc = WordprocessingDocument.Open(szFileName, true))
                //{
                //    #region.... Remove Scan Sign ...
                //    List<Table> lstTables = doc.MainDocumentPart.Document.Body.Elements<Table>().ToList();
                //    foreach (var item in lstTables)
                //    {
                //        List<TableRow> tableRow = item.Elements<TableRow>().ToList();
                //        foreach (var item1 in tableRow)
                //        {
                //            List<TableCell> tableCell = item1.Elements<TableCell>().ToList();
                //            foreach (var TblCell in tableCell)
                //            {
                //                List<Paragraph> Para = TblCell.Elements<Paragraph>().ToList();
                //                foreach (var p in Para)
                //                {
                //                    List<Run> Run = p.Elements<Run>().ToList();
                //                    foreach (var r in Run)
                //                    {
                //                        var Pic = r.Elements<Drawing>().FirstOrDefault();
                //                        if (Pic != null)
                //                        {
                //                            szScanSignCustumProperty = Pic.Inline.DocProperties.Description.Value;
                //                            string instructionText = String.Format(" MERGEFIELD  {0}  \\* MERGEFORMAT", szScanSignCustumProperty);
                //                            SimpleField simpleField1 = new SimpleField() { Instruction = instructionText };
                //                            Run run1 = new Run();
                //                            RunProperties runProperties1 = new RunProperties();
                //                            NoProof noProof1 = new NoProof();
                //                            runProperties1.Append(noProof1);
                //                            Text text1 = new Text();
                //                            text1.Text = String.Format("{0}", szScanSignCustumProperty);
                //                            run1.Append(runProperties1);
                //                            run1.Append(text1);
                //                            simpleField1.Append(run1);
                //                            p.Append(new OpenXmlElement[] { simpleField1 });

                //                            //p.Append(new SimpleField(new Run(new RunProperties(new NoProof()), new Text(szScanSignCustumProperty))));
                //                            a.Blip blip1 = Pic.Descendants<a.Blip>().FirstOrDefault();
                //                            IdPartPair idpp = doc.MainDocumentPart.Parts.Where(pa => pa.RelationshipId == blip1.Embed).FirstOrDefault();
                //                            doc.MainDocumentPart.DeletePart(idpp.RelationshipId);
                //                            if (Pic.Inline.DocProperties.Name.Value.Contains("ScanSignDrawing"))
                //                            {
                //                                rScanSignToRemove = r;
                //                            }
                //                        }
                //                    }
                //                }
                //                if (rScanSignToRemove != null)
                //                {
                //                    rScanSignToRemove.Remove();
                //                    rScanSignToRemove = null;
                //                }
                //            }
                //        }
                //    }

                //    #endregion

                //    #region .... Check in Header ....
                //    if (doc.MainDocumentPart.HeaderParts != null)
                //    {
                //        foreach (var header in doc.MainDocumentPart.HeaderParts)
                //        {
                //            #region.... Table ...
                //            if (header.Header.Descendants<Table>() != null)
                //            {

                //                List<Table> lstFooterTables = header.Header.Descendants<Table>().ToList();
                //                foreach (var item in lstFooterTables)
                //                {
                //                    List<TableRow> tableRow = item.Elements<TableRow>().ToList();
                //                    foreach (var item1 in tableRow)
                //                    {
                //                        List<TableCell> tableCell = item1.Elements<TableCell>().ToList();
                //                        foreach (var TblCell in tableCell)
                //                        {
                //                            List<Paragraph> Para = TblCell.Elements<Paragraph>().ToList();
                //                            foreach (var p in Para)
                //                            {
                //                                List<Run> Run = p.Elements<Run>().ToList();
                //                                foreach (var r in Run)
                //                                {
                //                                    var Pic = r.Elements<Drawing>().FirstOrDefault();
                //                                    if (Pic != null)
                //                                    {
                //                                        szScanSignCustumProperty = Pic.Inline.DocProperties.Description.Value;
                //                                        string instructionText = String.Format(" MERGEFIELD  {0}  \\* MERGEFORMAT", szScanSignCustumProperty);
                //                                        SimpleField simpleField1 = new SimpleField() { Instruction = instructionText };
                //                                        Run run1 = new Run();
                //                                        RunProperties runProperties1 = new RunProperties();
                //                                        NoProof noProof1 = new NoProof();
                //                                        runProperties1.Append(noProof1);
                //                                        Text text1 = new Text();
                //                                        text1.Text = String.Format("{0}", szScanSignCustumProperty);
                //                                        run1.Append(runProperties1);
                //                                        run1.Append(text1);
                //                                        simpleField1.Append(run1);
                //                                        p.Append(new OpenXmlElement[] { simpleField1 });
                //                                        //p.Append(new SimpleField(new Run(new RunProperties(new NoProof()), new Text(szScanSignCustumProperty))));
                //                                        a.Blip blip1 = Pic.Descendants<a.Blip>().FirstOrDefault();
                //                                        IdPartPair idpp = header.Parts.Where(pa => pa.RelationshipId == blip1.Embed).FirstOrDefault();
                //                                        header.DeletePart(idpp.RelationshipId);
                //                                        if (Pic.Inline.DocProperties.Name.Value.Contains("ScanSignDrawing"))
                //                                        {
                //                                            rScanSignToRemove = r;
                //                                        }
                //                                    }
                //                                }
                //                            }
                //                            if (rScanSignToRemove != null)
                //                            {
                //                                rScanSignToRemove.Remove();
                //                                rScanSignToRemove = null;
                //                            }
                //                        }
                //                    }
                //                }
                //            }
                //            #endregion
                //        }
                //    }
                //    #endregion

                //    #region .... Check in Footer ....
                //    if (doc.MainDocumentPart.FooterParts != null)
                //    {
                //        foreach (var footer in doc.MainDocumentPart.FooterParts)
                //        {
                //            #region.... Table ...
                //            if (footer.Footer.Descendants<Table>() != null)
                //            {

                //                List<Table> lstFooterTables = footer.Footer.Descendants<Table>().ToList();
                //                foreach (var item in lstFooterTables)
                //                {
                //                    List<TableRow> tableRow = item.Elements<TableRow>().ToList();
                //                    foreach (var item1 in tableRow)
                //                    {
                //                        List<TableCell> tableCell = item1.Elements<TableCell>().ToList();
                //                        foreach (var TblCell in tableCell)
                //                        {
                //                            List<Paragraph> Para = TblCell.Elements<Paragraph>().ToList();
                //                            foreach (var p in Para)
                //                            {
                //                                List<Run> Run = p.Elements<Run>().ToList();
                //                                foreach (var r in Run)
                //                                {
                //                                    var Pic = r.Elements<Drawing>().FirstOrDefault();
                //                                    if (Pic != null)
                //                                    {


                //                                        szScanSignCustumProperty = Pic.Inline.DocProperties.Description.Value;
                //                                        string instructionText = String.Format(" MERGEFIELD  {0}  \\* MERGEFORMAT", szScanSignCustumProperty);
                //                                        SimpleField simpleField1 = new SimpleField() { Instruction = instructionText };
                //                                        Run run1 = new Run();
                //                                        RunProperties runProperties1 = new RunProperties();
                //                                        NoProof noProof1 = new NoProof();
                //                                        runProperties1.Append(noProof1);
                //                                        Text text1 = new Text();
                //                                        text1.Text = String.Format("{0}", szScanSignCustumProperty);
                //                                        run1.Append(runProperties1);
                //                                        run1.Append(text1);
                //                                        simpleField1.Append(run1);
                //                                        p.Append(new OpenXmlElement[] { simpleField1 });
                //                                        //p.Append(new SimpleField(new Run(new RunProperties(new NoProof()), new Text(szScanSignCustumProperty))));
                //                                        a.Blip blip1 = Pic.Descendants<a.Blip>().FirstOrDefault();
                //                                        IdPartPair idpp = footer.Parts.Where(pa => pa.RelationshipId == blip1.Embed).FirstOrDefault();
                //                                        footer.DeletePart(idpp.RelationshipId);
                //                                        if (Pic.Inline.DocProperties.Name.Value.Contains("ScanSignDrawing"))
                //                                        {
                //                                            rScanSignToRemove = r;
                //                                        }
                //                                    }
                //                                }
                //                            }
                //                            if (rScanSignToRemove != null)
                //                            {
                //                                rScanSignToRemove.Remove();
                //                                rScanSignToRemove = null;
                //                            }
                //                        }
                //                    }
                //                }
                //            }
                //            #endregion
                //        }
                //    }
                //    #endregion
                //}
            }
            finally
            {

            }
        }

        public static bool ContentSearch(string szFileName, string szContentToSearch)
        {
            msgError = "";
            _bResult = false;
            WordprocessingDocument _objDoc = null;
            try
            {
                string szFileToSearch = szFileName;
                using (_objDoc = WordprocessingDocument.Open(szFileToSearch, true))
                {
                    string szDocText = "";
                    if (_objDoc.MainDocumentPart.Document.Body.InnerText != "" || _objDoc.MainDocumentPart.Document.Body.InnerText != null)
                    {
                        szDocText = ((DocumentFormat.OpenXml.OpenXmlCompositeElement)(_objDoc.MainDocumentPart.Document.Body)).InnerText;
                        //   bResult = szDocText.Contains(txtSearch.Text);
                        _bResult = szDocText.ToLower().Contains(szContentToSearch.ToLower());
                    }
                }

                #region  .. Header ..
                using (_objDoc = WordprocessingDocument.Open(szFileToSearch, true))
                {
                    //_activeWordDocumement = open word open xml document
                    //Search through headers
                    if (_objDoc.MainDocumentPart.HeaderParts != null)
                    {
                        foreach (var header in _objDoc.MainDocumentPart.HeaderParts)
                        {
                            if (_bResult)
                                break;
                            if (header.Header.Descendants<Paragraph>() != null)
                            {
                                foreach (var para in header.Header.Descendants<Paragraph>())
                                {
                                    if (_bResult)
                                        break;
                                    foreach (Run r in para.Descendants<Run>())
                                    {
                                        if (_bResult)
                                            break;
                                        foreach (Text t in r.Descendants<Text>())
                                        {
                                            if (_bResult)
                                                break;
                                            //   bResult = t.InnerText.Contains(txtSearch.Text);
                                            _bResult = t.InnerText.ToLower().Contains(szContentToSearch.ToLower());

                                        }

                                    }
                                }
                            }

                        }
                    }
                }
                #endregion


                #region ... Footer ....
                using (_objDoc = WordprocessingDocument.Open(szFileToSearch, true))
                {
                    //_activeWordDocumement = open word open xml document
                    //Search through Footer
                    if (_objDoc.MainDocumentPart.FooterParts != null)
                    {
                        foreach (var footer in _objDoc.MainDocumentPart.FooterParts)
                        {
                            if (_bResult)
                                break;
                            if (footer.Footer.Descendants<Paragraph>() != null)
                            {
                                foreach (var para in footer.Footer.Descendants<Paragraph>())
                                {
                                    if (_bResult)
                                        break;
                                    foreach (Run r in para.Descendants<Run>())
                                    {
                                        if (_bResult)
                                            break;
                                        foreach (Text t in r.Descendants<Text>())
                                        {
                                            if (_bResult)
                                                break;
                                            _bResult = t.InnerText.ToLower().Contains(szContentToSearch.ToLower());
                                            // bResult = t.InnerText.Contains(txtSearch.Text);
                                        }

                                    }
                                }
                            }
                        }
                    }
                }

                #endregion
            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = ex.StackTrace;
            }
            finally
            {
                if (_objDoc != null)
                    _objDoc = null;
            }
            return _bResult;
        }

        /// <summary>
        /// .Update Custum variable For CP Convert to PDF for Bridge
        /// </summary>
        /// <param name="document"></param>
        /// 
        public static bool Update_CustumVariables(object szFileName)
        {
            object objMissing = Type.Missing;
            Word.Application objApp = null;
            Word.Document objDoc = null;
            objApp = new Word.Application();
            objDoc = new Word.Document();
            Object oCustom;
            _bResult = true;
            try
            {
                objDoc = objApp.Documents.Open(ref szFileName, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing);
                oCustom = objDoc.CustomDocumentProperties;
                for (int i = 1; i <= objDoc.Fields.Count; i++)
                {
                    if (objDoc.Fields[i].Code.Text != null)
                    {
                        objDoc.Fields[i].DoClick();
                        objDoc.Fields[i].Update();
                    }
                }
                Word.HeaderFooter headerg = objDoc.Sections[1].Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage];
                Word.Range rangeg = headerg.Range;
                rangeg.Fields.Update();

                headerg = objDoc.Sections[1].Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                rangeg = headerg.Range;
                rangeg.Fields.Update();

                objDoc.PrintPreview();
                objDoc.ClosePrintPreview();
            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = ex.Message;
            }
            finally
            {
                if (objDoc != null)
                {
                    if (objDoc != null)
                        objDoc.Close(ref objMissing, ref objMissing, ref objMissing);

                    if (objApp != null)
                        objApp.Quit(ref objMissing, ref objMissing, ref objMissing);

                    objDoc = null;
                    objApp = null;
                }
                GC.Collect();
            }
            return _bResult;
        }




        #endregion

        #region .... Private Methods ...
        private static Uri FixUri(string brokenUri)
        {
            return new Uri("http://broken-link/");
        }

        private static void AcceptAll(WordprocessingDocument document)
        {
            MainDocumentPart mainPart = document.MainDocumentPart;
            AcceptRevisionsForPart(mainPart);
            foreach (var p in mainPart.HeaderParts)
                AcceptRevisionsForPart(p);
            foreach (var p in mainPart.FooterParts)
                AcceptRevisionsForPart(p);
        }


        private static XDocument GetXDocument(this OpenXmlPart part)
        {
            XDocument xdoc = part.Annotation<XDocument>();
            if (xdoc != null)
                return xdoc;
            using (StreamReader sr = new StreamReader(part.GetStream()))
            using (XmlReader xr = XmlReader.Create(sr))
                xdoc = XDocument.Load(xr);
            part.AddAnnotation(xdoc);
            return xdoc;
        }



        private static void AcceptRevisionsForPart(OpenXmlPart part)
        {
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";  // Change on Here Before send code to Kedar
            XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";	// Change on Here Before send code to Kedar

            XDocument xDoc = part.GetXDocument();

            // Accept inserted text, run properties for paragraph marks, etc.
            // =============================================================================

            // Find all w:ins elements, remove the w:ins element and move its children nodes
            // up one level.
            //
            // Before:
            //
            //<w:p w:rsidR="005C743F"
            //     w:rsidRDefault="005C743F">
            //    <w:ins w:id="1"
            //           w:author="Eric White"
            //           w:date="2008-04-27T08:33:00Z">
            //        <w:r>
            //            <w:t xml:space="preserve">Text </w:t>
            //        </w:r>
            //    </w:ins>
            //    <w:r>
            //        <w:t>inserted at the beginning of the paragraph.</w:t>
            //    </w:r>
            //</w:p>
            //
            // After:
            //
            //<w:p w:rsidR="005C743F"
            //     w:rsidRDefault="005C743F">
            //    <w:r>
            //        <w:t xml:space="preserve">Text </w:t>
            //    </w:r>
            //    <w:r>
            //        <w:t>inserted at the beginning of the paragraph.</w:t>
            //    </w:r>
            //</w:p>
            //
            // Some of the w:ins elements have no children, for instance a run property that
            // indicates that a paragraph has been inserted looks like this:
            //
            //<w:p w:rsidR="005C743F"
            //     w:rsidRDefault="005C743F">
            //    <w:pPr>
            //        <w:rPr>
            //            <w:ins w:id="2"
            //                   w:author="Eric White"
            //                   w:date="2008-04-27T08:34:00Z"/>
            //        </w:rPr>
            //    </w:pPr>
            //    <w:r>
            //        <w:t xml:space="preserve">Text inserted at the end of the </w:t>
            //    </w:r>
            //    <w:r>
            //        <w:t>paragraph.</w:t>
            //    </w:r>
            //</w:p>
            //
            // and we want it to look like this:
            //
            //<w:p w:rsidR="005C743F"
            //     w:rsidRDefault="005C743F">
            //    <w:pPr>
            //        <w:rPr>
            //        </w:rPr>
            //    </w:pPr>
            //    <w:r>
            //        <w:t xml:space="preserve">Text inserted at the end of the </w:t>
            //    </w:r>
            //    <w:r>
            //        <w:t>paragraph.</w:t>
            //    </w:r>
            //</w:p>
            //
            // There are no child nodes of that w:ins element, so the following code works properly
            // to accept this revision too.

            foreach (var x in xDoc.Descendants(w + "ins").ToList())
                x.ReplaceWith(x.Nodes());

            // Accept deleted paragraphs.
            // =============================================================================
            //
            // Find all w:p/w:pPr/w:rPr/w:del nodes, append all child nodes of the following paragraph
            // to the paragraph containing the w:p/w:pPr/w:rPr/w:del node, delete the following paragraph,
            // and delete the w:p/w:pPr/w:rPr/w:del node.
            //
            // Before:
            //
            //<w:p w:rsidR="004A3F21"
            //     w:rsidDel="004A3F21"
            //     w:rsidRDefault="004A3F21"
            //     w:rsidP="008A4F80">
            //    <w:pPr>
            //        <w:rPr>
            //            <w:del w:id="17"
            //                   w:author="Eric White"
            //                   w:date="2008-04-27T13:02:00Z" />
            //        </w:rPr>
            //    </w:pPr>
            //    <w:r>
            //        <w:t xml:space="preserve">This paragraph is joined </w:t>
            //    </w:r>
            //</w:p>
            //<w:p w:rsidR="004A3F21"
            //     w:rsidRDefault="004A3F21"
            //     w:rsidP="008A4F80">
            //    <w:r>
            //        <w:t>with this paragraph.</w:t>
            //    </w:r>
            //</w:p>
            //
            // After:
            //
            //<w:p w:rsidR="004A3F21"
            //     w:rsidDel="004A3F21"
            //     w:rsidRDefault="004A3F21"
            //     w:rsidP="008A4F80">
            //    <w:pPr>
            //        <w:rPr>
            //        </w:rPr>
            //    </w:pPr>
            //    <w:r>
            //        <w:t xml:space="preserve">This paragraph is joined </w:t>
            //    </w:r>
            //    <w:r>
            //        <w:t>with this paragraph.</w:t>
            //    </w:r>
            //</w:p>

            foreach (var x in xDoc.Descendants(w + "p")
                                  .Elements(w + "pPr")
                                  .Elements(w + "rPr")
                                  .Elements(w + "del")
                                  .Reverse()
                                  .ToList())
            {
                // find the w:p element
                XElement p = x.Ancestors(w + "p").First();

                // add the elements of the paragraph following.  This code will work even if there
                // is no following paragraph.
                p.Add(p.ElementsAfterSelf(w + "p").Take(1).Elements());

                // Remove the next paragraph if there is one.
                p.ElementsAfterSelf(w + "p").Take(1).Remove();

                // remove the w:p/w:pPr/w:rPr/w:del node
                x.Remove();
            }

            // Accept changes for changes in formatting on paragraphs.
            // Accept changes for changes in formatting on runs.
            // Accept changes for applied styles to a table.
            // Accept changes for grid changes to a table.
            // Accept changes for column properties.
            // Accept changes for row properties.
            // Accept revisions for table level property exceptions.
            // Accept revisions for section properties.
            var pPrChange = w + "pPrChange";
            var rPrChange = w + "rPrChange";
            var tblPrChange = w + "tblPrChange";
            var tblGridChange = w + "tblGridChange";
            var tcPrChange = w + "tcPrChange";
            var trPrChange = w + "trPrChange";
            var tblPrExChange = w + "tblPrExChange";
            var sectPrChange = w + "sectPrChange";
            xDoc.Descendants()
                .Where(x =>
                        x.Name == pPrChange ||
                        x.Name == rPrChange ||
                        x.Name == tblPrChange ||
                        x.Name == tblGridChange ||
                        x.Name == tcPrChange ||
                        x.Name == trPrChange ||
                        x.Name == tblPrExChange ||
                        x.Name == sectPrChange)
                .Remove();

            // Accept changes for deleted rows in tables.
            // Find all w:tr/w:trPr/w:del elements, and remove the w:tr elements.
            foreach (var x in xDoc.Descendants(w + "tr")
                                  .Elements(w + "trPr")
                                  .Elements(w + "del")
                                  .ToList())
                x.Parent.Parent.Remove();

            // Accept deleted text in paragraphs.
            // =============================================================================
            //
            // Remove all w:p/w:del nodes.
            //
            // Before:
            //
            //<w:p w:rsidR="005C743F"
            //     w:rsidRDefault="005C743F"
            //     w:rsidP="005C743F">
            //    <w:r>
            //        <w:t xml:space="preserve">This line contains </w:t>
            //    </w:r>
            //    <w:del w:id="8"
            //           w:author="Eric White"
            //           w:date="2008-04-27T08:37:00Z">
            //        <w:r w:rsidDel="005C743F">
            //            <w:delText xml:space="preserve">deleted </w:delText>
            //        </w:r>
            //    </w:del>
            //    <w:r>
            //        <w:t>text.</w:t>
            //    </w:r>
            //</w:p>
            //
            // After:
            //
            //<w:p w:rsidR="005C743F"
            //     w:rsidRDefault="005C743F"
            //     w:rsidP="005C743F">
            //    <w:r>
            //        <w:t xml:space="preserve">This line contains </w:t>
            //    </w:r>
            //    <w:r>
            //        <w:t>text.</w:t>
            //    </w:r>
            //</w:p>

            // The Remove extension method uses snapshot semantics.
            //xDoc.Descendants(w + "p")
            //    .Elements(w + "del")
            //    .Remove();
            xDoc.Descendants(w + "del")
                .Remove();

            // Currently this code doesn't handle:
            // w:tblStylePr/w:trPr/w:del
            // w:style/w:trPr/w:del
            // MathML
            // Smart tags
            // Custom XML

            // I don't believe that the following is strictly necessary, but not sure.  I notice that if
            // you remove all rows from a table, and open and save using Word 2007, tables with no rows are deleted.
            // In any case, remove the tables that no longer have rows.
            xDoc.Descendants(w + "tbl")
                .Where(x => !x.Elements(w + "tr").Any())
                .Remove();

            // Accept moved paragraphs.
            // Find all w:p/w:moveFrom elements, and remove the w:p element.
            // Remove all w:moveFromRangeEnd elements
            // Find all w:p/w:moveTo elements, remove the w:moveTo elements, and promote their children to
            // be children of the w:p element
            // Remove all w:moveToRangeStart and w:moveToRangeEnd elements
            foreach (var x in xDoc.Descendants(w + "p").Elements(w + "moveFrom").ToList())
            {
                var p = x.Ancestors(w + "p").First();
                p.Remove();
            }
            xDoc.Descendants(w + "moveFromRangeEnd").Remove();
            foreach (var x in xDoc.Descendants(w + "p")
                                  .Elements(w + "moveTo")
                                  .ToList())
                x.ReplaceWith(x.Nodes());
            xDoc.Descendants(w + "moveToRangeStart").Remove();
            xDoc.Descendants(w + "moveToRangeEnd").Remove();

            using (XmlWriter xw = XmlWriter.Create(part.GetStream(FileMode.Create, FileAccess.Write)))
                xDoc.Save(xw);
        }

        /// <summary>
        /// Accepts all text change revisions in the document
        /// </summary>
        //public static void AcceptAll(WordprocessingDocument document)
        //{
        //    MainDocumentPart mainPart = document.MainDocumentPart;
        //    AcceptRevisionsForPart(mainPart);
        //    foreach (var p in mainPart.HeaderParts)
        //        AcceptRevisionsForPart(p);
        //    foreach (var p in mainPart.FooterParts)
        //        AcceptRevisionsForPart(p);
        //}

        //static void Main(string[] args)
        //{
        //    using (WordprocessingDocument doc = WordprocessingDocument.Open("Test.docx", true))
        //    {
        //        AcceptAll(doc);
        //    }
        //}
        private static void ScanSign(string szFileName, string szKey, string szValue, bool bRemoveScanSign)
        {
            string szScansignFilePath = string.Empty;
            Run rScanSignToRemove = null;
            Run rScanSignDocPropertyRemove = null;
            Paragraph pScanSignDocPropertyRemove = null;
            Paragraph paraAuthorScanSign = null;
            bool bFieldCode = false;
            string szSCanSignOfUser = string.Empty;
            try
            {

                //using (eDocsDN_DocX.DocX objDoc = eDocsDN_DocX.DocX.Load(szFileName))
                //{
                //    List<XElement> xEleDocument = objDoc.Get_document_xElement(objDoc);
                //    foreach (XElement doc in xEleDocument)
                //    {
                //        foreach (XElement e in doc.Descendants(XName.Get("drawing", w.NamespaceName)))
                //        {
                //            //..Get Element 
                //            var szScanSignId =
                //             (
                //                 from d in e.Descendants()
                //                 where d.Name.LocalName.Contains("ScanSignDrawing_")
                //                 select d
                //             ).SingleOrDefault();

                //            //child.Ancestors().TakeUntil(node => node.Id == desiredId);
                //            XElement Para = e.Ancestors().FirstOrDefault(node => node.NodeType == XmlNodeType.Element);


                //        }
                //    }
                //}

                using (WordprocessingDocument doc = WordprocessingDocument.Open(szFileName, true))
                {
                    //AddParts(doc, szImageFolderPath, szImageName, paraR1ScanSign);
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
                                            if (Pic.Inline.DocProperties.Name.Value.Contains("ScanSignDrawing_"))
                                                rScanSignToRemove = r;
                                        }
                                        var ScanSign = r.Elements<FieldCode>().FirstOrDefault();
                                        if (ScanSign != null)
                                        {
                                            //A1_Sign
                                            //foreach (var dicItem in ScanSignOfUser)
                                            //{
                                            if (ScanSign.InnerText.Contains(szKey))
                                            {
                                                bFieldCode = true;
                                                r.Remove();
                                                szSCanSignOfUser = szKey;
                                                szScansignFilePath = szValue;
                                                szScansignFilePath = szScansignFilePath.Replace("\\", "\\\\");
                                                bFieldCode = true;
                                                paraAuthorScanSign = p;
                                            }
                                            //}
                                        }
                                    }
                                    if (!bFieldCode)
                                    {
                                        var ScanSign = p.Elements<SimpleField>().FirstOrDefault();
                                        if (ScanSign != null)
                                        {
                                            //A1_Sign
                                            //foreach (var dicItem in ScanSignOfUser)
                                            //{
                                            if (ScanSign.InnerText.Contains(szKey))
                                            {
                                                bFieldCode = true;
                                                ScanSign.Remove();
                                                szSCanSignOfUser = szKey;
                                                szScansignFilePath = szValue;
                                                szScansignFilePath = szScansignFilePath.Replace("\\", "\\\\");
                                                bFieldCode = true;
                                                paraAuthorScanSign = p;
                                            }
                                            //}
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
                                                        if (Pic.Inline.DocProperties.Name.Value.Contains("ScanSignDrawing_"))
                                                            rScanSignToRemove = r;
                                                    }
                                                    var ScanSign = r.Elements<FieldCode>().FirstOrDefault();
                                                    if (ScanSign != null)
                                                    {
                                                        //A1_Sign
                                                        //foreach (var dicItem in ScanSignOfUser)
                                                        //{
                                                        if (ScanSign.InnerText.Contains(szKey))
                                                        {

                                                            bFieldCode = true;
                                                            r.Remove();
                                                            szSCanSignOfUser = szKey;
                                                            szScansignFilePath = szValue;
                                                            szScansignFilePath = szScansignFilePath.Replace("\\", "\\\\");
                                                            bFieldCode = true;
                                                            paraAuthorScanSign = p;
                                                        }
                                                        //}
                                                    }
                                                }
                                                if (!bFieldCode)
                                                {
                                                    var ScanSign = p.Elements<SimpleField>().FirstOrDefault();
                                                    if (ScanSign != null)
                                                    {
                                                        //A1_Sign
                                                        //foreach (var dicItem in ScanSignOfUser)
                                                        //{
                                                        if (ScanSign.InnerText.Contains(szKey))
                                                        {
                                                            bFieldCode = true;
                                                            ScanSign.Remove();
                                                            szSCanSignOfUser = szKey;
                                                            szScansignFilePath = szValue;
                                                            szScansignFilePath = szScansignFilePath.Replace("\\", "\\\\");
                                                            bFieldCode = true;
                                                            paraAuthorScanSign = p;
                                                        }
                                                        //}
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
                                                        if (Pic.Inline.DocProperties.Name.Value.Contains("ScanSignDrawing_"))
                                                            rScanSignToRemove = r;
                                                    }
                                                    var ScanSign = r.Elements<FieldCode>().FirstOrDefault();
                                                    if (ScanSign != null)
                                                    {
                                                        //A1_Sign
                                                        //foreach (var dicItem in ScanSignOfUser)
                                                        //{
                                                        if (ScanSign.InnerText.Contains(szKey))
                                                        {

                                                            bFieldCode = true;
                                                            r.Remove();
                                                            szSCanSignOfUser = szKey;
                                                            szScansignFilePath = szValue;
                                                            szScansignFilePath = szScansignFilePath.Replace("\\", "\\\\");
                                                            bFieldCode = true;
                                                            paraAuthorScanSign = p;
                                                        }
                                                        //}
                                                    }
                                                }
                                                if (!bFieldCode)
                                                {
                                                    var ScanSign = p.Elements<SimpleField>().FirstOrDefault();
                                                    if (ScanSign != null)
                                                    {
                                                        //A1_Sign
                                                        //foreach (var dicItem in ScanSignOfUser)
                                                        //{
                                                        if (ScanSign.InnerText.Contains(szKey))
                                                        {
                                                            bFieldCode = true;
                                                            ScanSign.Remove();
                                                            szSCanSignOfUser = szKey;
                                                            szScansignFilePath = szValue;
                                                            szScansignFilePath = szScansignFilePath.Replace("\\", "\\\\");
                                                            bFieldCode = true;
                                                            paraAuthorScanSign = p;
                                                        }
                                                        //}
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
            }
            finally
            {

            }
        }


        private static void ScanSign(string szFileName, Dictionary<string, string> ScanSignOfUser, bool bRemoveScanSign)
        {
            string szScansignFilePath = string.Empty;
            Run rScanSignToRemove = null;
            Run rScanSignDocPropertyRemove = null;
            Paragraph pScanSignDocPropertyRemove = null;
            Paragraph paraAuthorScanSign = null;
            bool bFieldCode = false;
            string szSCanSignOfUser = string.Empty;
            try
            {

                //...Remove ScanImage Tag
                //.. Check for Custom Property


                //using (eDocsDN_DocX.DocX objDoc = eDocsDN_DocX.DocX.Load(szFileName))
                //{
                //    List<XElement> xEleDocument = objDoc.Get_document_xElement(objDoc);
                //    foreach (XElement doc in xEleDocument)
                //    {
                //        foreach (XElement e in doc.Descendants(XName.Get("drawing", w.NamespaceName)))
                //        {
                //            //..Get Element 
                //            var szScanSignId =
                //             (
                //                 from d in e.Descendants()
                //                 where d.Name.LocalName.Contains("ScanSignDrawing_")
                //                 select d
                //             ).SingleOrDefault();

                //            //child.Ancestors().TakeUntil(node => node.Id == desiredId);
                //            XElement Para = e.Ancestors().FirstOrDefault(node => node.NodeType == XmlNodeType.Element);


                //        }
                //    }
                //}

                using (WordprocessingDocument doc = WordprocessingDocument.Open(szFileName, true))
                {
                    //AddParts(doc, szImageFolderPath, szImageName, paraR1ScanSign);
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
                                            if (Pic.Inline.DocProperties.Name.Value.Contains("ScanSignDrawing_"))
                                                rScanSignToRemove = r;
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
                                                        if (Pic.Inline.DocProperties.Name.Value.Contains("ScanSignDrawing_"))
                                                            rScanSignToRemove = r;
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
                                                        if (Pic.Inline.DocProperties.Name.Value.Contains("ScanSignDrawing_"))
                                                            rScanSignToRemove = r;
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
            }
            finally
            {

            }
        }

        private static void AddParts(WordprocessingDocument parent, string imageFilePath, Paragraph P, string szRole)
        {
            long imageWidthEMU = 0;
            long imageHeightEMU = 0;

            //var Part = parent.GetPartById("rId" + System.IO.Path.GetFileNameWithoutExtension(imageFilePath));
            //string PartID = parent.GetIdOfPart(Part);

            var imagePart = parent.MainDocumentPart.AddNewPart<ImagePart>("image/jpeg", "rId" + System.IO.Path.GetFileNameWithoutExtension(imageFilePath));

            GenerateImagePart(imagePart, imageFilePath,
             ref imageWidthEMU, ref imageHeightEMU);
            var picture = GenerateMainDocumentPart(System.IO.Path.GetFileNameWithoutExtension(imageFilePath), imageWidthEMU, imageHeightEMU, szRole);
            P.AppendChild(new Run(picture));
        }

        private static void AddParts(FooterPart parent, string imageFilePath, Paragraph P, string szRole)
        {
            long imageWidthEMU = 0;
            long imageHeightEMU = 0;
            var imagePart = parent.AddNewPart<ImagePart>("image/jpeg", "rId" + System.IO.Path.GetFileNameWithoutExtension(imageFilePath));
            GenerateImagePart(imagePart, imageFilePath,
             ref imageWidthEMU, ref imageHeightEMU);
            var picture = GenerateMainDocumentPart(System.IO.Path.GetFileNameWithoutExtension(imageFilePath), imageWidthEMU, imageHeightEMU, szRole);
            P.AppendChild(new Run(picture));
        }

        private static void AddParts(HeaderPart parent, string imageFilePath, Paragraph P, string szRole)
        {
            long imageWidthEMU = 0;
            long imageHeightEMU = 0;
            var imagePart = parent.AddNewPart<ImagePart>("image/jpeg", "rId" + System.IO.Path.GetFileNameWithoutExtension(imageFilePath));
            GenerateImagePart(imagePart, imageFilePath,
             ref imageWidthEMU, ref imageHeightEMU);
            var picture = GenerateMainDocumentPart(System.IO.Path.GetFileNameWithoutExtension(imageFilePath), imageWidthEMU, imageHeightEMU, szRole);
            P.AppendChild(new Run(picture));
        }


        public static void GenerateImagePart(OpenXmlPart part, string imageFilePath, ref long imageWidthEMU, ref long imageHeightEMU)
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

        private static Drawing GenerateMainDocumentPart(string imageFileName, long imageWidthEMU, long imageHeightEMU, string szScanSignOfRole)
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
                          )
                          { Preset = a.ShapeTypeValues.Rectangle }))
                    )
                    { Uri = GraphicDataUri })
                )
                {
                    DistanceFromTop = (UInt32Value)0U,
                    DistanceFromBottom = (UInt32Value)0U,
                    DistanceFromLeft = (UInt32Value)0U,
                    DistanceFromRight = (UInt32Value)0U
                });
            return element;
        }

        private static void DeleteComments(string fileName, string author = "")
        {
            // Get an existing Wordprocessing document.
            using (WordprocessingDocument document =
                WordprocessingDocument.Open(fileName, true))
            {
                // Set commentPart to the document WordprocessingCommentsPart, 
                // if it exists.
                WordprocessingCommentsPart commentPart =
                    document.MainDocumentPart.WordprocessingCommentsPart;

                // If no WordprocessingCommentsPart exists, there can be no 
                // comments. Stop execution and return from the method.
                if (commentPart == null)
                {
                    return;
                }

                // Create a list of comments by the specified author, or
                // if the author name is empty, all authors.
                List<Comment> commentsToDelete =
                    commentPart.Comments.Elements<Comment>().ToList();
                if (!String.IsNullOrEmpty(author))
                {
                    commentsToDelete = commentsToDelete.
                    Where(c => c.Author == author).ToList();
                }
                IEnumerable<string> commentIds =
                    commentsToDelete.Select(r => r.Id.Value);

                // Delete each comment in commentToDelete from the 
                // Comments collection.
                foreach (Comment c in commentsToDelete)
                {
                    c.Remove();
                }

                // Save the comment part change.
                commentPart.Comments.Save();

                Document doc = document.MainDocumentPart.Document;

                // Delete CommentRangeStart for each
                // deleted comment in the main document.
                List<CommentRangeStart> commentRangeStartToDelete =
                    doc.Descendants<CommentRangeStart>().
                    Where(c => commentIds.Contains(c.Id.Value)).ToList();
                foreach (CommentRangeStart c in commentRangeStartToDelete)
                {
                    c.Remove();
                }

                // Delete CommentRangeEnd for each deleted comment in the main document.
                List<CommentRangeEnd> commentRangeEndToDelete =
                    doc.Descendants<CommentRangeEnd>().
                    Where(c => commentIds.Contains(c.Id.Value)).ToList();
                foreach (CommentRangeEnd c in commentRangeEndToDelete)
                {
                    c.Remove();
                }

                // Delete CommentReference for each deleted comment in the main document.
                List<CommentReference> commentRangeReferenceToDelete =
                    doc.Descendants<CommentReference>().
                    Where(c => commentIds.Contains(c.Id.Value)).ToList();
                foreach (CommentReference c in commentRangeReferenceToDelete)
                {
                    c.Remove();
                }

                // Save changes back to the MainDocumentPart part.
                doc.Save();
            }
        }

        private static bool FindReplaceWatermarkText(Run runWatermark, string nWatermarktext)
        {
            bool success = false;
            //Check, if run contains watermark
            if (runWatermark.Descendants<Picture>() != null)
            {
                foreach (var pic in runWatermark.Descendants<Picture>())
                {
                    if (pic.Descendants<V.Shape>() != null)
                    {
                        foreach (var shape in pic.Descendants<V.Shape>())
                        {
                            //Loop over all shapes and replace textpath
                            if (shape.Descendants<V.TextPath>() != null)
                            {
                                V.TextPath txtPath = shape.Descendants<V.TextPath>().FirstOrDefault();
                                if (txtPath != null)
                                {
                                    txtPath.String = nWatermarktext;
                                    success = true;
                                }
                            }
                        }
                    }
                }
            }
            return success;
        }

        private static bool FindRemoveWatermark(Run runWatermark)
        {
            bool success = false;

            //Check, if run contains watermark
            if (runWatermark.Descendants<Picture>() != null)
            {
                var listPic = runWatermark.Descendants<Picture>().ToList();

                for (int n = listPic.Count; n > 0; n--)
                {
                    if (listPic[n - 1].Descendants<V.Shape>() != null)
                    {
                        if (listPic[n - 1].Descendants<V.Shape>().Where(s => s.Type == "#_x0000_t136").Count() > 0)
                        {
                            //Found -> remove
                            listPic[n - 1].Remove();
                            success = true;
                            break;
                        }
                    }
                }

            }

            return success;
        }

        private static Run CreateWatermarkRun(string watermarkText)
        {
            V.Shape shapeWM = null;
            Run runWatermark = new Run();

            RunProperties runWMProperties = new RunProperties();
            NoProof noProofWM = new NoProof();

            runWMProperties.Append(noProofWM);

            Picture pictureWM = new Picture();

            V.Shapetype shapetypeWM = new V.Shapetype()
            {
                Id = "_x0000_t136",
                CoordinateSize = "21600,21600",
                OptionalNumber = 136,
                Adjustment = "10800",
                EdgePath = "m@7,l@8,m@5,21600l@6,21600e"
            };

            V.Formulas formulasWM = new V.Formulas();
            V.Formula formula1 = new V.Formula()
            {
                Equation = "sum #0 0 10800"
            };
            V.Formula formula2 = new V.Formula()
            {
                Equation = "prod #0 2 1"
            };
            V.Formula formula3 = new V.Formula()
            {
                Equation = "sum 21600 0 @1"
            };
            V.Formula formula4 = new V.Formula()
            {
                Equation = "sum 0 0 @2"
            };
            V.Formula formula5 = new V.Formula()
            {
                Equation = "sum 21600 0 @3"
            };
            V.Formula formula6 = new V.Formula()
            {
                Equation = "if @0 @3 0"
            };
            V.Formula formula7 = new V.Formula()
            {
                Equation = "if @0 21600 @1"
            };
            V.Formula formula8 = new V.Formula()
            {
                Equation = "if @0 0 @2"
            };
            V.Formula formula9 = new V.Formula()
            {
                Equation = "if @0 @4 21600"
            };
            V.Formula formula10 = new V.Formula()
            {
                Equation = "mid @5 @6"
            };
            V.Formula formula11 = new V.Formula()
            {
                Equation = "mid @8 @5"
            };
            V.Formula formula12 = new V.Formula()
            {
                Equation = "mid @7 @8"
            };
            V.Formula formula13 = new V.Formula()
            {
                Equation = "mid @6 @7"
            };
            V.Formula formula14 = new V.Formula()
            {
                Equation = "sum @6 0 @5"
            };

            formulasWM.Append(formula1);
            formulasWM.Append(formula2);
            formulasWM.Append(formula3);
            formulasWM.Append(formula4);
            formulasWM.Append(formula5);
            formulasWM.Append(formula6);
            formulasWM.Append(formula7);
            formulasWM.Append(formula8);
            formulasWM.Append(formula9);
            formulasWM.Append(formula10);
            formulasWM.Append(formula11);
            formulasWM.Append(formula12);
            formulasWM.Append(formula13);
            formulasWM.Append(formula14);
            V.Path pathWM = new V.Path()
            {
                AllowTextPath = true,
                ConnectionPointType = ConnectValues.Custom,
                ConnectionPoints = "@9,0;@10,10800;@11,21600;@12,10800",
                ConnectAngles = "270,180,90,0"
            };
            V.TextPath textPathWM = new V.TextPath()
            {
                On = true,
                FitShape = true
            };

            V.ShapeHandles shapeHandlesWM = new V.ShapeHandles();
            V.ShapeHandle shapeHandleWM = new V.ShapeHandle()
            {
                Position = "#0,bottomRight",
                XRange = "6629,14971"
            };

            shapeHandlesWM.Append(shapeHandleWM);
            Lock lockWM = new Lock()
            {
                Extension = V.ExtensionHandlingBehaviorValues.Edit,
                TextLock = true,
                ShapeType = true
            };

            shapetypeWM.Append(formulasWM);
            shapetypeWM.Append(pathWM);
            shapetypeWM.Append(textPathWM);
            shapetypeWM.Append(shapeHandlesWM);
            shapetypeWM.Append(lockWM);

            if (watermarkText.Length < 6)
            {
                shapeWM = new V.Shape()
                {
                    Id = "PowerPlusWaterMarkObject346762751",
                    Style = "position:absolute;margin-left:0;margin-top:0;width:406.1pt;height:162.45pt;rotation:315;z-index:-1;visibility:visible;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin",
                    //Style= "position:absolute;margin-left:0;margin-top:0;width:406.1pt;height:auto;rotation:315;z-index:-1;visibility:visible;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin",
                    OptionalString = "_x0000_s2050",
                    AllowInCell = false,
                    FillColor = "silver",
                    Stroked = false,
                    Type = "#_x0000_t136"
                };
            }
            else
            {
                shapeWM = new V.Shape()
                {
                    Id = "PowerPlusWaterMarkObject346762751",
                    Style = "position:absolute;margin-left:0;margin-top:0;width:406.1pt;height:68pt;rotation:315;z-index:-1;visibility:visible;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin",
                    //Style = "position:absolute;margin-left:0;margin-top:0;width:406.1pt;height:auto;rotation:315;z-index:-1;visibility:visible;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin",
                    OptionalString = "_x0000_s2050",
                    AllowInCell = false,
                    FillColor = "silver",
                    Stroked = false,
                    Type = "#_x0000_t136"

                };
            }
            V.Fill fillWM = new V.Fill()
            {
                Opacity = "0.9"
            };
            V.TextPath textPath2WM = new V.TextPath()
            {
                Style = "font-family:\"Times New Roman\";font-size:1pt",
                String = watermarkText
            };

            shapeWM.Append(fillWM);
            shapeWM.Append(textPath2WM);

            pictureWM.Append(shapetypeWM);
            pictureWM.Append(shapeWM);

            runWatermark.Append(runWMProperties);
            runWatermark.Append(pictureWM);

            //Rückgabe des Runs
            return runWatermark;
        }
        static int[] InitialCodeArray = { 0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C, 0x0E10, 0xF1CE, 0x313E, 0x1872, 0xE139, 0xD40F, 0x84F9, 0x280C, 0xA96A, 0x4EC3 };
        static int[,] EncryptionMatrix = new int[15, 7]
    {
            
            /* char 1  */ {0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09},
            /* char 2  */ {0x7B61, 0xF6C2, 0xFDA5, 0xEB6B, 0xC6F7, 0x9DCF, 0x2BBF},
            /* char 3  */ {0x4563, 0x8AC6, 0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0},
            /* char 4  */ {0x0375, 0x06EA, 0x0DD4, 0x1BA8, 0x3750, 0x6EA0, 0xDD40},
            /* char 5  */ {0xD849, 0xA0B3, 0x5147, 0xA28E, 0x553D, 0xAA7A, 0x44D5},
            /* char 6  */ {0x6F45, 0xDE8A, 0xAD35, 0x4A4B, 0x9496, 0x390D, 0x721A},
            /* char 7  */ {0xEB23, 0xC667, 0x9CEF, 0x29FF, 0x53FE, 0xA7FC, 0x5FD9},
            /* char 8  */ {0x47D3, 0x8FA6, 0x0F6D, 0x1EDA, 0x3DB4, 0x7B68, 0xF6D0},
            /* char 9  */ {0xB861, 0x60E3, 0xC1C6, 0x93AD, 0x377B, 0x6EF6, 0xDDEC},
            /* char 10 */ {0x45A0, 0x8B40, 0x06A1, 0x0D42, 0x1A84, 0x3508, 0x6A10},
            /* char 11 */ {0xAA51, 0x4483, 0x8906, 0x022D, 0x045A, 0x08B4, 0x1168},
            /* char 12 */ {0x76B4, 0xED68, 0xCAF1, 0x85C3, 0x1BA7, 0x374E, 0x6E9C},
            /* char 13 */ {0x3730, 0x6E60, 0xDCC0, 0xA9A1, 0x4363, 0x86C6, 0x1DAD},
            /* char 14 */ {0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373, 0x06E6, 0x0DCC},
            /* char 15 */ {0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48C4}
      };


        private static byte[] concatByteArrays(byte[] array1, byte[] array2)
        {
            byte[] result = new byte[array1.Length + array2.Length];
            Buffer.BlockCopy(array2, 0, result, 0, array2.Length);
            Buffer.BlockCopy(array1, 0, result, array2.Length, array1.Length);
            return result;
        }
        #endregion
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



    public class objOpenXMLClass
    {
        #region .... Variable Declaration ....
        bool _bResult;
        Dictionary<string, string> dicScanSign;

        #endregion

        #region ... Property ....

        public string msgError { get; set; }

        #endregion

        #region .... Public Method ....
        public bool ContentSearch(string szFileName, string szContentToSearch)
        {
            msgError = "";
            _bResult = false;
            WordprocessingDocument _objDoc = null;
            try
            {
                string szFileToSearch = szFileName;
                using (_objDoc = WordprocessingDocument.Open(szFileToSearch, true))
                {
                    string szDocText = "";
                    if (_objDoc.MainDocumentPart.Document.Body.InnerText != "" || _objDoc.MainDocumentPart.Document.Body.InnerText != null)
                    {
                        szDocText = ((DocumentFormat.OpenXml.OpenXmlCompositeElement)(_objDoc.MainDocumentPart.Document.Body)).InnerText;
                        //   bResult = szDocText.Contains(txtSearch.Text);
                        _bResult = szDocText.ToLower().Contains(szContentToSearch.ToLower());
                    }
                }

                #region  .. Header ..
                using (_objDoc = WordprocessingDocument.Open(szFileToSearch, true))
                {
                    //_activeWordDocumement = open word open xml document
                    //Search through headers
                    if (_objDoc.MainDocumentPart.HeaderParts != null)
                    {
                        foreach (var header in _objDoc.MainDocumentPart.HeaderParts)
                        {
                            if (_bResult)
                                break;
                            if (header.Header.Descendants<Paragraph>() != null)
                            {
                                foreach (var para in header.Header.Descendants<Paragraph>())
                                {
                                    if (_bResult)
                                        break;
                                    foreach (Run r in para.Descendants<Run>())
                                    {
                                        if (_bResult)
                                            break;
                                        foreach (Text t in r.Descendants<Text>())
                                        {
                                            if (_bResult)
                                                break;
                                            //   bResult = t.InnerText.Contains(txtSearch.Text);
                                            _bResult = t.InnerText.ToLower().Contains(szContentToSearch.ToLower());

                                        }

                                    }
                                }
                            }

                        }
                    }
                }
                #endregion


                #region ... Footer ....
                using (_objDoc = WordprocessingDocument.Open(szFileToSearch, true))
                {
                    //_activeWordDocumement = open word open xml document
                    //Search through Footer
                    if (_objDoc.MainDocumentPart.FooterParts != null)
                    {
                        foreach (var footer in _objDoc.MainDocumentPart.FooterParts)
                        {
                            if (_bResult)
                                break;
                            if (footer.Footer.Descendants<Paragraph>() != null)
                            {
                                foreach (var para in footer.Footer.Descendants<Paragraph>())
                                {
                                    if (_bResult)
                                        break;
                                    foreach (Run r in para.Descendants<Run>())
                                    {
                                        if (_bResult)
                                            break;
                                        foreach (Text t in r.Descendants<Text>())
                                        {
                                            if (_bResult)
                                                break;
                                            _bResult = t.InnerText.ToLower().Contains(szContentToSearch.ToLower());
                                            // bResult = t.InnerText.Contains(txtSearch.Text);
                                        }

                                    }
                                }
                            }
                        }
                    }
                }

                #endregion
            }
            catch (Exception ex)
            {
                _bResult = false;
                msgError = ex.StackTrace;
            }
            finally
            {
                if (_objDoc != null)
                    _objDoc = null;
            }
            return _bResult;
        }
        #endregion
    }


}
