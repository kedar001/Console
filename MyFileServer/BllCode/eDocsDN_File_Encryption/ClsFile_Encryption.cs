using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using eDocsDN_Get_Directory_Info;

namespace eDocsDN_File_Encryption
{
    public class ClsFile_Encryption
    {
        #region .... Variable Declaration ...
        string _szSqlQuery = string.Empty;
        string _szAppXmlPath = string.Empty;
        public string EncryptionKey = "ESPL";

        #endregion

        #region .... Property .....
        public string msgError { get; set; }
        public string EncryptedFileExtension { get { return ".dgg"; } }
        #endregion

        public enum Action
        {
            Encrypt = 0,
            Decrypt
        }

        #region ..... Constructor ....
        public ClsFile_Encryption()
        {
            msgError = string.Empty;
        }

        #endregion

        #region .... Public Method ....
        public bool Check_File_Exist_in_Encrypted_Or_Decrypted_Format(Action eFile_To_Checked_In_format, Directory_Attributes oSource_Dir, File_Data oSourceFile)
        {
            bool bResult = true;
            try
            {
                switch (eFile_To_Checked_In_format)
                {
                    case Action.Encrypt:
                        bResult = File.Exists(oSource_Dir.Directory_Path + Path.GetFileNameWithoutExtension(oSource_Dir.Directory_Path) + "." + EncryptedFileExtension);
                        break;
                    case Action.Decrypt:
                        bResult = File.Exists(oSource_Dir.Directory_Path + oSourceFile.File_Name);
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            finally { }
            return bResult;
        }
        public File_Data Encrypt_Descypt_Files(Action eEncrypt_Decrypt, Directory_Attributes oSource_Dir, File_Data oSourceFile, Directory_Attributes oDestination_Dir, File_Data oDestination_File)
        {
            try
            {
                switch (eEncrypt_Decrypt)
                {
                    case Action.Encrypt:
                        Encrypt(oSource_Dir.Directory_Path + oSourceFile.File_Name, oDestination_Dir.Directory_Path + Path.GetFileNameWithoutExtension(oDestination_File.Destination_File_Name) + "." + EncryptedFileExtension);
                        oDestination_File.CheckSum = GetMd5_CheckSum(oDestination_Dir.Directory_Path + Path.GetFileNameWithoutExtension(oDestination_File.Destination_File_Name) + "." + EncryptedFileExtension);
                        break;
                    case Action.Decrypt:
                        Decrypt(oSource_Dir.Directory_Path + Path.GetFileNameWithoutExtension(oSourceFile.File_Name) + "." + EncryptedFileExtension, oDestination_Dir.Directory_Path + oDestination_File.Destination_File_Name);
                        oDestination_File.CheckSum = GetMd5_CheckSum(oDestination_Dir.Directory_Path + oDestination_File.Destination_File_Name);
                        oDestination_File.Source_File_CheckSum = GetMd5_CheckSum(oDestination_Dir.Directory_Path + oDestination_File.Destination_File_Name);
                        oDestination_File.Data = File.ReadAllBytes(oDestination_Dir.Directory_Path + oDestination_File.Destination_File_Name);
                        break;
                    default:
                        break;
                }


            }
            catch (Exception ex)
            {
                msgError = ex.ToString();
                oDestination_File = null;
            }
            return oDestination_File;
        }

        #endregion

        #region .... Private functions ....

        private void Encrypt(string inputFilePath, string outputfilePath)
        {
            using (Aes encryptor = Aes.Create())
            {
                Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(EncryptionKey, new byte[] { 0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76 });
                encryptor.Key = pdb.GetBytes(32);
                encryptor.IV = pdb.GetBytes(16);
                using (FileStream fsOutput = new FileStream(outputfilePath, FileMode.Create))
                {
                    using (CryptoStream cs = new CryptoStream(fsOutput, encryptor.CreateEncryptor(), CryptoStreamMode.Write))
                    {
                        using (FileStream fsInput = new FileStream(inputFilePath, FileMode.Open))
                        {
                            int data;
                            while ((data = fsInput.ReadByte()) != -1)
                            {
                                cs.WriteByte((byte)data);
                            }
                        }
                    }
                }
            }
        }

        private void Decrypt(string inputFilePath, string outputfilePath)
        {
            using (Aes encryptor = Aes.Create())
            {
                Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(EncryptionKey, new byte[] { 0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76 });
                encryptor.Key = pdb.GetBytes(32);
                encryptor.IV = pdb.GetBytes(16);
                using (FileStream fsInput = new FileStream(inputFilePath, FileMode.Open))
                {
                    using (CryptoStream cs = new CryptoStream(fsInput, encryptor.CreateDecryptor(), CryptoStreamMode.Read))
                    {
                        if (File.Exists(outputfilePath))
                            File.Delete(outputfilePath);
                        using (FileStream fsOutput = new FileStream(outputfilePath, FileMode.Create))
                        {
                            int data;
                            while ((data = cs.ReadByte()) != -1)
                            {
                                fsOutput.WriteByte((byte)data);
                            }
                        }
                    }
                }
            }
        }

        public string GetMd5_CheckSum(string szFilePath)
        {
            msgError = "";
            MD5 md5Hasher = MD5.Create();
            StreamReader sr = new StreamReader(szFilePath);
            byte[] data = md5Hasher.ComputeHash(sr.BaseStream);
            sr.Close();
            StringBuilder sBuilder = new StringBuilder();
            for (int i = 0; i < data.Length; i++)
            {
                sBuilder.Append(data[i].ToString("x2"));
            }
            return sBuilder.ToString();
        }
        public string GetMd5_CheckSum(byte[] arrDocument)
        {
            msgError = "";
            MD5 md5Hasher = MD5.Create();
            byte[] data = md5Hasher.ComputeHash(arrDocument);
            StringBuilder sBuilder = new StringBuilder();
            for (int i = 0; i < data.Length; i++)
            {
                sBuilder.Append(data[i].ToString("x2"));
            }
            return sBuilder.ToString();
        }



        #endregion
    }
}
