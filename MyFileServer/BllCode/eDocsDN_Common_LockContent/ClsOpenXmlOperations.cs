/// <summary>
/// Developer Name:Kedar
/// Date:09/13/2014
/// DRT-4314
/// </summary>
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using System.Security;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using Op = DocumentFormat.OpenXml.CustomProperties;
using V = DocumentFormat.OpenXml.Vml;
using System.Security.Cryptography;
using System.IO;
using DocumentFormat.OpenXml.CustomProperties;
//using word = Microsoft.Office.Interop.Word;
//using excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;



namespace eDocsDN_Common_LockContent
{


    internal class ClsOpenXmlOperations
    {

        public string msgError { get; set; }

        bool _bResult;
        //word.Application objApp = null;
        //word.Document objDoc = null;
        object objMissing = Type.Missing;

        public enum LockType
        {
            ReadOnly, None, Comments, TrackedChanges, Forms
        }

        public ClsOpenXmlOperations()
        {
            msgError = "";
        }
        /// <summary>
        /// Developer Name:Kedar
        /// Date:09/13/2014
        /// DRT:
        /// </summary>
        /// <param name="szFilePath"></param>
        /// <param name="lockType"></param>
        /// <param name="szPassword"></param>
        /// <param name="szStatus"></param>
        /// <param name="bFormsPrintData"></param>
        /// <returns></returns>
        public bool LockDocument(string szFilePath, LockType lockType, string szPassword, string szStatus, bool bFormsPrintData)
        {
            _bResult = true;
            WordprocessingDocument objDoc = null;
            DocumentProtection documentProtection;
            try
            {

                if (szStatus != "")
                {
                    using (eDocDN_Update_Custom_Properties.ClsUpdate_Custom_Properties objUpdate = new eDocDN_Update_Custom_Properties.ClsUpdate_Custom_Properties())
                    {
                        objUpdate.FileName = szFilePath;
                        objUpdate.lstCustom_Properties = new List<eDocDN_Update_Custom_Properties.Custom_Property>();
                        objUpdate.lstCustom_Properties.Add(new eDocDN_Update_Custom_Properties.Custom_Property() { PropertyName = "Status", PropertyValue = szStatus });
                        if (!objUpdate.Update_Document_Custom_Property(objUpdate.lstCustom_Properties))
                            throw new Exception(objUpdate.msgError);
                    }
                }


                // Generate the Salt
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

                using (objDoc = WordprocessingDocument.Open(szFilePath, true))
                {
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
                        UInt32 uintVal = new UInt32();
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
                        UInt32 uintVal = new UInt32();
                        uintVal = (uint)iterations;
                        documentProtection.CryptographicSpinCount = uintVal;
                        documentProtection.Hash = Convert.ToBase64String(generatedKey);
                        documentProtection.Salt = Convert.ToBase64String(arrSalt);
                        objDoc.MainDocumentPart.DocumentSettingsPart.Settings.AppendChild(documentProtection);
                        objDoc.MainDocumentPart.DocumentSettingsPart.Settings.Save();
                        objDoc.CustomFilePropertiesPart.Properties.Save();
                    }
                }
                objDoc = null;

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
                    objDoc = null;
                }
            }
            return _bResult;
        }
        public bool UnlockDocument(string szFilePath, string szPassword)
        {
            _bResult = true;
            WordprocessingDocument objDoc;
            DocumentProtection documentProtection;

            try
            {

                // Generate the Salt
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

                using (objDoc = WordprocessingDocument.Open(szFilePath, true))
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
                        UInt32 uintVal = new UInt32();
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
                        UInt32 uintVal = new UInt32();
                        uintVal = (uint)iterations;
                        documentProtection.CryptographicSpinCount = uintVal;
                        documentProtection.Hash = Convert.ToBase64String(generatedKey);
                        documentProtection.Salt = Convert.ToBase64String(arrSalt);
                    }
                    objDoc.MainDocumentPart.DocumentSettingsPart.Settings.AppendChild(documentProtection);
                    objDoc.MainDocumentPart.DocumentSettingsPart.Settings.Save();
                }
                objDoc = null;


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

    }

}
