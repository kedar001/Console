using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using eDocDN_Update_Custom_Properties;

namespace eDocsDN_OpenXml_Operations
{
    public class ClsDocument_Operations
    {
        #region .... Variable Declaration ...
        bool _bResult;
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

        #region ..... Constructor ...
        public ClsDocument_Operations()
        {
            msgError = string.Empty;
        }
        #endregion

        #region .... Property ...
        public string msgError { get; set; }
        public string Document_Password { set { } get { return "Espl123&;"; } }
        #endregion

        #region .... Public method ....
        public Stream LockDocument(Stream strmDocument, LockType lockType, bool bFormsPrintData)
        {
            _bResult = true;
            WordprocessingDocument objDoc = null;
            DocumentProtection documentProtection;
            RandomNumberGenerator rand = null;
            StringBuilder sb = null;
            HashAlgorithm sha1 = null;
            UInt32 uintVal;

            byte[] tmpArray1;
            byte[] tmpArray2;
            byte[] tempKey;
            byte[] iterator = null;
            byte[] arrSalt = null;
            byte[] generatedKey = null;
            byte[] arrByteChars = null;

            int iterations;
            int intMaxPasswordLength;
            int intLowOrderWord = 0;

            try
            {
                #region .... Generate Password ...
                arrSalt = new byte[16];
                rand = new RNGCryptoServiceProvider();
                rand.GetNonZeroBytes(arrSalt);
                generatedKey = new byte[4];
                intMaxPasswordLength = 15;
                if (!String.IsNullOrEmpty(Document_Password))
                {
                    Document_Password = Document_Password.Substring(0, Math.Min(Document_Password.Length, intMaxPasswordLength));
                    arrByteChars = new byte[Document_Password.Length];

                    for (int intLoop = 0; intLoop < Document_Password.Length; intLoop++)
                    {
                        int intTemp = Convert.ToInt32(Document_Password[intLoop]);
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
                    intLowOrderWord = 0;
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
                sb = new StringBuilder();
                for (int intTemp = 0; intTemp < 4; intTemp++)
                {
                    sb.Append(Convert.ToString(generatedKey[intTemp], 16));
                }
                generatedKey = Encoding.Unicode.GetBytes(sb.ToString().ToUpper());
                tmpArray1 = generatedKey;
                tmpArray2 = arrSalt;
                tempKey = new byte[tmpArray1.Length + tmpArray2.Length];
                Buffer.BlockCopy(tmpArray2, 0, tempKey, 0, tmpArray2.Length);
                Buffer.BlockCopy(tmpArray1, 0, tempKey, tmpArray2.Length, tmpArray1.Length);
                generatedKey = tempKey;
                iterations = 50000;
                sha1 = new SHA1Managed();
                generatedKey = sha1.ComputeHash(generatedKey);
                iterator = new byte[4];
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

                #region .... Process Document .....
                objDoc = WordprocessingDocument.Open(strmDocument, true);
                var documentSettings = objDoc.MainDocumentPart.DocumentSettingsPart;
                documentProtection = documentSettings
                                            .Settings
                                            .FirstOrDefault(it =>
                                                    it is DocumentProtection)
                                            as DocumentProtection;

                if (documentProtection != null)
                {

                    documentSettings.Settings.RemoveAllChildren<DocumentProtection>();
                    objDoc.MainDocumentPart.DocumentSettingsPart.Settings.Save();
                    objDoc.CustomFilePropertiesPart.Properties.Save();

                    switch (lockType)
                    {
                        case LockType.ReadOnly:
                            documentProtection.Edit = DocumentProtectionValues.ReadOnly;
                            break;
                        case LockType.Comments:
                            documentProtection.Edit = DocumentProtectionValues.Comments;
                            break;
                        case LockType.Forms:
                            documentProtection.Edit = DocumentProtectionValues.Forms;
                            break;
                        case LockType.TrackedChanges:
                            documentProtection.Edit = DocumentProtectionValues.TrackedChanges;
                            break;
                        case LockType.None:
                        default:
                            documentProtection.Edit = DocumentProtectionValues.None;
                            break;
                    }

                    DocumentFormat.OpenXml.OnOffValue docProtection = new DocumentFormat.OpenXml.OnOffValue(true);
                    documentProtection.Enforcement = docProtection;
                    documentProtection.CryptographicAlgorithmClass = CryptAlgorithmClassValues.Hash;
                    documentProtection.CryptographicProviderType = CryptProviderValues.RsaFull;
                    documentProtection.CryptographicAlgorithmType = CryptAlgorithmValues.TypeAny;
                    documentProtection.CryptographicAlgorithmSid = 4; // SHA1
                    //    The iteration count is unsigned
                    uintVal = new UInt32();
                    uintVal = (uint)iterations;
                    documentProtection.CryptographicSpinCount = uintVal;
                    documentProtection.Hash = Convert.ToBase64String(generatedKey);
                    documentProtection.Salt = Convert.ToBase64String(arrSalt);
                    objDoc.MainDocumentPart.DocumentSettingsPart.Settings.AppendChild(documentProtection);
                    objDoc.MainDocumentPart.DocumentSettingsPart.Settings.Save();
                    objDoc.CustomFilePropertiesPart.Properties.Save();
                }
                else
                {
                    documentProtection = new DocumentProtection();
                    objDoc.MainDocumentPart.DocumentSettingsPart.Settings.PrintFormsData = new PrintFormsData();
                    if (bFormsPrintData)
                        objDoc.MainDocumentPart.DocumentSettingsPart.Settings.PrintFormsData.Val = OnOffValue.FromBoolean(true);
                    else
                        objDoc.MainDocumentPart.DocumentSettingsPart.Settings.PrintFormsData.Val = OnOffValue.FromBoolean(false);

                    switch (lockType)
                    {
                        case LockType.ReadOnly:
                            documentProtection.Edit = DocumentProtectionValues.ReadOnly;
                            break;
                        case LockType.Comments:
                            documentProtection.Edit = DocumentProtectionValues.Comments;
                            break;
                        case LockType.Forms:
                            documentProtection.Edit = DocumentProtectionValues.Forms;
                            break;
                        case LockType.TrackedChanges:
                            documentProtection.Edit = DocumentProtectionValues.TrackedChanges;
                            break;
                        case LockType.None:
                        default:
                            documentProtection.Edit = DocumentProtectionValues.None;
                            break;
                    }

                    DocumentFormat.OpenXml.OnOffValue docProtection = new DocumentFormat.OpenXml.OnOffValue(true);
                    documentProtection.Enforcement = docProtection;

                    documentProtection.CryptographicAlgorithmClass = CryptAlgorithmClassValues.Hash;
                    documentProtection.CryptographicProviderType = CryptProviderValues.RsaFull;
                    documentProtection.CryptographicAlgorithmType = CryptAlgorithmValues.TypeAny;
                    documentProtection.CryptographicAlgorithmSid = 4; // SHA1
                    //    The iteration count is unsigned
                    uintVal = new UInt32();
                    uintVal = (uint)iterations;
                    documentProtection.CryptographicSpinCount = uintVal;
                    documentProtection.Hash = Convert.ToBase64String(generatedKey);
                    documentProtection.Salt = Convert.ToBase64String(arrSalt);
                    objDoc.MainDocumentPart.DocumentSettingsPart.Settings.AppendChild(documentProtection);
                    objDoc.MainDocumentPart.DocumentSettingsPart.Settings.Save();
                    objDoc.CustomFilePropertiesPart.Properties.Save();
                }

                #endregion
            }
            catch (Exception ex)
            {
                _bResult = false;
                strmDocument = null;
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
                documentProtection = null;
                rand = null;
                sb = null;
                sha1 = null;
                tmpArray1 = null;
                tmpArray2 = null;
                tempKey = null;
                iterator = null;
                arrSalt = null;
                generatedKey = null;
                arrByteChars = null;

            }
            return strmDocument;
        }
        public Stream UnlockDocument(Stream strmDocument)
        {
            _bResult = true;
            WordprocessingDocument objDoc = null;
            DocumentProtection documentProtection;
            RandomNumberGenerator rand = null;
            StringBuilder sb = null;
            HashAlgorithm sha1 = null;
            UInt32 uintVal;

            byte[] tmpArray1;
            byte[] tmpArray2;
            byte[] tempKey;
            byte[] iterator = null;
            byte[] arrSalt = null;
            byte[] generatedKey = null;
            byte[] arrByteChars = null;

            int iterations;
            int intMaxPasswordLength;

            try
            {
                #region .... Generate Password ...
                arrSalt = new byte[16];
                rand = new RNGCryptoServiceProvider();
                rand.GetNonZeroBytes(arrSalt);
                generatedKey = new byte[4];
                intMaxPasswordLength = 15;
                if (!String.IsNullOrEmpty(Document_Password))
                {
                    Document_Password = Document_Password.Substring(0, Math.Min(Document_Password.Length, intMaxPasswordLength));
                    arrByteChars = new byte[Document_Password.Length];

                    for (int intLoop = 0; intLoop < Document_Password.Length; intLoop++)
                    {
                        int intTemp = Convert.ToInt32(Document_Password[intLoop]);
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
                sb = new StringBuilder();
                for (int intTemp = 0; intTemp < 4; intTemp++)
                {
                    sb.Append(Convert.ToString(generatedKey[intTemp], 16));
                }
                generatedKey = Encoding.Unicode.GetBytes(sb.ToString().ToUpper());
                tmpArray1 = generatedKey;
                tmpArray2 = arrSalt;
                tempKey = new byte[tmpArray1.Length + tmpArray2.Length];
                Buffer.BlockCopy(tmpArray2, 0, tempKey, 0, tmpArray2.Length);
                Buffer.BlockCopy(tmpArray1, 0, tempKey, tmpArray2.Length, tmpArray1.Length);
                generatedKey = tempKey;
                iterations = 50000;
                sha1 = new SHA1Managed();
                generatedKey = sha1.ComputeHash(generatedKey);
                iterator = new byte[4];
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

                #region .... Process Document .....

                using (objDoc = WordprocessingDocument.Open(strmDocument, true))
                {
                    var documentSettings = objDoc.MainDocumentPart.DocumentSettingsPart;
                    documentProtection = documentSettings.Settings.FirstOrDefault(it => it is DocumentProtection) as DocumentProtection;
                    if (documentProtection != null)
                    {
                        documentSettings.Settings.RemoveAllChildren<DocumentProtection>();
                        documentProtection.Edit = DocumentProtectionValues.None;
                        documentSettings.Settings.PrintFormsData = new PrintFormsData();
                        documentSettings.Settings.PrintFormsData.Val = OnOffValue.FromBoolean(false);
                        DocumentFormat.OpenXml.OnOffValue docProtection = new DocumentFormat.OpenXml.OnOffValue(true);
                        documentProtection.Enforcement = docProtection;
                        documentProtection.CryptographicAlgorithmClass = CryptAlgorithmClassValues.Hash;
                        documentProtection.CryptographicProviderType = CryptProviderValues.RsaFull;
                        documentProtection.CryptographicAlgorithmType = CryptAlgorithmValues.TypeAny;
                        documentProtection.CryptographicAlgorithmSid = 4; // SHA1
                        uintVal = new UInt32();
                        uintVal = (uint)iterations;
                        documentProtection.CryptographicSpinCount = uintVal;
                        documentProtection.Hash = Convert.ToBase64String(generatedKey);
                        documentProtection.Salt = Convert.ToBase64String(arrSalt);
                    }
                    else
                    {
                        documentProtection = new DocumentProtection();
                        documentProtection.Edit = DocumentProtectionValues.None;
                        objDoc.MainDocumentPart.DocumentSettingsPart.Settings.PrintFormsData = new PrintFormsData();
                        objDoc.MainDocumentPart.DocumentSettingsPart.Settings.PrintFormsData.Val = OnOffValue.FromBoolean(false);
                        DocumentFormat.OpenXml.OnOffValue docProtection = new DocumentFormat.OpenXml.OnOffValue(true);
                        documentProtection.Enforcement = docProtection;
                        documentProtection.CryptographicAlgorithmClass = CryptAlgorithmClassValues.Hash;
                        documentProtection.CryptographicProviderType = CryptProviderValues.RsaFull;
                        documentProtection.CryptographicAlgorithmType = CryptAlgorithmValues.TypeAny;
                        documentProtection.CryptographicAlgorithmSid = 4; // SHA1
                        uintVal = new UInt32();
                        uintVal = (uint)iterations;
                        documentProtection.CryptographicSpinCount = uintVal;
                        documentProtection.Hash = Convert.ToBase64String(generatedKey);
                        documentProtection.Salt = Convert.ToBase64String(arrSalt);
                    }
                    objDoc.MainDocumentPart.DocumentSettingsPart.Settings.AppendChild(documentProtection);
                    objDoc.MainDocumentPart.DocumentSettingsPart.Settings.Save();
                }
                #endregion
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
                _bResult = false;
                strmDocument = null;
            }
            finally
            {
                if (objDoc != null)
                {
                    objDoc.Close();
                    objDoc.Dispose();
                }
                objDoc = null;
                documentProtection = null;
                rand = null;
                sb = null;
                sha1 = null;
                tmpArray1 = null;
                tmpArray2 = null;
                tempKey = null;
                iterator = null;
                arrSalt = null;
                generatedKey = null;
                arrByteChars = null;
            }
            return strmDocument;
        }

        #endregion

        #region ..... Private Functions ...
        int[] InitialCodeArray = { 0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C, 0x0E10, 0xF1CE, 0x313E, 0x1872, 0xE139, 0xD40F, 0x84F9, 0x280C, 0xA96A, 0x4EC3 };
        int[,] EncryptionMatrix = new int[15, 7]
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
        private byte[] concatByteArrays(byte[] array1, byte[] array2)
        {
            byte[] result = new byte[array1.Length + array2.Length];
            Buffer.BlockCopy(array2, 0, result, 0, array2.Length);
            Buffer.BlockCopy(array1, 0, result, array2.Length, array1.Length);
            return result;
        }

        #endregion
    }
}
