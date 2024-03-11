using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.IO;
using System.Web;

namespace eDocsDN_Get_Directory_Info
{
    public enum LockType
    {
        ReadOnly = 0, None, Comments, TrackedChanges, Forms
    }


    [DataContract]
    public class File_Data : IDisposable
    {
        File_Operations objFile_Operations = null;
        public File_Data()
        {
            Need_File_Blob = true;
            objFile_Operations = new File_Operations(null, null, null, null, null, false, false);

        }
        [DataMember]
        public int SurrKey { get; set; }
        [DataMember]
        public int Serial_Number { get; set; }
        [DataMember]
        public int Draft_Version { get; set; }

        [DataMember]
        public string Source_Directory { get; set; }
        [DataMember]
        public string Destination_Directory { get; set; }
        [DataMember]
        public string File_Name { get; set; }
        [DataMember]
        public string Directory { get; set; }
        [DataMember]
        public string Destination_File_Name { get; set; }
        [DataMember]
        public string CheckSum { get; set; }
        [DataMember]
        public string Source_File_CheckSum { get; set; }
        [DataMember]
        public byte[] Data { get; set; }
        [DataMember]
        public string SourceFilePath { get; set; }
        [DataMember]
        public byte[] ConvertedPDF_Data { get; set; }
        [DataMember]
        public string User_Id { get; set; }
        [DataMember]
        public string Type_of_User { get; set; }
        [DataMember]
        public bool TrackChanges { get; set; }
        [DataMember]
        public bool PrintFormData { get; set; }
        [DataMember]
        public bool Remove_Scan_Sign { get; set; }
        [DataMember]
        public bool Convert_To_PDF { get; set; }
        [DataMember]
        public bool Need_File_Blob { get; set; }
        [DataMember]
        public string Referance_Source_Location { get; set; }
        [DataMember]
        public string Referance_Source_Dir { get; set; }
        [DataMember]
        public int Referance_Source_SurrKey { get; set; }
        [DataMember]
        public int Referance_Source_SerialNumber { get; set; }
        [DataMember]
        public string Referance_Source_FileName { get; set; }

        [DataMember]
        public File_Operations File_Operations { get; set; }


        public void Dispose()
        {
            if (Data != null)
                Array.Clear(Data, 0, Data.Length);
            Data = null;
            if (ConvertedPDF_Data != null)
                Array.Clear(ConvertedPDF_Data, 0, ConvertedPDF_Data.Length);
            ConvertedPDF_Data = null;
            Source_Directory = string.Empty;
            Destination_Directory = string.Empty;
            File_Name = string.Empty;
            Directory = string.Empty;
            Destination_File_Name = string.Empty;
            CheckSum = string.Empty;
            Source_File_CheckSum = string.Empty;
            User_Id = string.Empty;
            Type_of_User = string.Empty;
            if (File_Operations != null)
                File_Operations.Dispose();
            File_Operations = null;
        }
    }

    //public class Source_File_Data
    //{
    //    [DataMember]
    //    public int SurrKey { get; set; }
    //    [DataMember]
    //    public int Serial_Number { get; set; }
    //    [DataMember]
    //    public string Source_Directory { get; set; }
    //    [DataMember]
    //    public string Destination_Directory { get; set; }
    //    [DataMember]
    //    public string File_Name { get; set; }
    //    [DataMember]
    //    public string Directory { get; set; }

    //}

    [DataContract]
    public class File_Operations : IDisposable
    {
        [DataMember]
        public LockUnlockFile LockUnlock { get; set; }
        [DataMember]
        public Update_Document_Custom_Variables Update_Properties { get; set; }
        [DataMember]
        public Update_Users_Comments UpdateComments { get; set; }
        [DataMember]
        public Scan_Signature ScanSignature { get; set; }
        [DataMember]
        public Print_Documents Print_Documents { get; set; }
        [DataMember]
        public bool ConvertToPdf { get; set; }
        [DataMember]
        public bool DocumentPreCheck { get; set; }



        public File_Operations(LockUnlockFile objLockUnclock, Update_Document_Custom_Variables objUpdate_Properties, Update_Users_Comments objUpdateComments, Scan_Signature objScanSign, Print_Documents objPrint_Documents, bool bConvertToPdf, bool bDocumentPreCheck)
        {
            LockUnlock = objLockUnclock;
            Update_Properties = objUpdate_Properties;
            UpdateComments = objUpdateComments;
            ScanSignature = objScanSign;
            Print_Documents = objPrint_Documents;
            ConvertToPdf = bConvertToPdf;
            DocumentPreCheck = bDocumentPreCheck;
        }

        public void Dispose()
        {
            LockUnlock = null;
            Update_Properties = null;
            UpdateComments = null;
            ScanSignature = null;
        }
    }


    [DataContract]
    public class LockUnlockFile
    {
        [DataMember]
        public bool LockFile { get; set; }
        [DataMember]
        public LockType Lock_Type { get; set; }
    }
    [DataContract]
    public class Update_Document_Custom_Variables : IDisposable
    {
        [DataMember]
        public Documents_Process eDocument_Process { get; set; }
        [DataMember]
        public Documents_Status eDocument_Status { get; set; }

        public void Dispose()
        {
        }
    }
    [DataContract]
    public class Update_Users_Comments
    {
        [DataMember]
        public string UserID { get; set; }
        [DataMember]
        public DateTime dtDateTime { get; set; }
    }
    [DataContract]
    public class Scan_Signature
    {
        [DataMember]
        public bool Remove_Scan_Sign { get; set; }
        [DataMember]
        public Dictionary<string, string> Users_Scan_Sign { get; set; }
    }
    [DataContract]
    public class Print_Documents
    {
        [DataMember]
        public bool Clear_comments { get; set; }
    }



    public enum Documents_Process
    {
        Controller_Live = 0,
        Transfer_Document,
        Controller_Publish,
        Document_Recall,
        TR4,
        Document_Issuance,
        Preview,
        PDC,
        Update_User_Comments,
        Attach_Custom_Variables,
        Attach_Custom_Variables_To_Template,
        obsolete_Document,
        Expired_Document,
        NA,
        Repaire_Document
    }

    public enum Documents_Status
    {
        Draft,
        Draft_Approved,
        Publish,
        Issued,
        Expired,
        Obsolete
    }



}
