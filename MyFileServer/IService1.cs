using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;
using eDocsDN_Get_Directory_Info;
using System.Data;
using System.ServiceModel.Activation;

namespace MyFileServer
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the interface name "IService1" in both code and config file together.
    [ServiceContract(
    SessionMode = SessionMode.Allowed)]

    public interface IService1
    {

        [OperationContract]
        [TransactionFlow(TransactionFlowOption.Allowed)]
        File_Data CopyFile(File_Data oFileData);

        [OperationContract(Name = "ListOfFiles")]
        [TransactionFlow(TransactionFlowOption.Allowed)]
        List<File_Data> CopyFile(List<File_Data> lstFileData);


        [OperationContract]
        [TransactionFlow(TransactionFlowOption.Allowed)]
        Boolean Delete_File(File_Data oFileData);


        [OperationContract]
        Boolean Check_File_Exist(File_Data oFileData);

        [OperationContract]
        File_Data Get_File_Information(File_Data oFileData);

        [OperationContract(Name = "Check_File_Is_Locked")]
        Boolean Check_File_Is_Locked(string szFileName);

        [OperationContract(Name = "Check_File_Is_Locked_WithVersion")]
        Boolean Check_File_Is_Locked(string szFileName, string szOfficeVersion);


        [OperationContract]
        Boolean Pre_Check_File(string szFileName);

        [OperationContract(Name = "Pre_Check_File_Blob")]
        Boolean Pre_Check_File(byte[] ArrFile);

        [OperationContract(Name = "Check_File")]
        Boolean Pre_Check_File(File_Data oFileData);


        [OperationContract]
        List<File_Data> Get_Documents(File_Data oFileData);

        [OperationContract]
        String Get_Server_Date_Time();

        [OperationContract]
        String Get_Server_Time();

        [OperationContract]
        String Get_Document_CheckSum(File_Data oFileData);

        [OperationContract]
        Boolean Convert_Document_To_PDF(File_Data oFileData);


        //[OperationContract]
        //DataTable Read_Excel_File_For_Migration(string szFileName);

        //[OperationContract]
        //File_Data Process_File_Physically(File_Data oFileData);

        // TODO: Add your service operations here
    }


    // Use a data contract as illustrated in the sample below to add composite types to service operations.
    //[DataContract]
    //public class CompositeType
    //{
    //    bool boolValue = true;
    //    string stringValue = "Hello ";

    //    [DataMember]
    //    public bool BoolValue
    //    {
    //        get { return boolValue; }
    //        set { boolValue = value; }
    //    }

    //    [DataMember]
    //    public string StringValue
    //    {
    //        get { return stringValue; }
    //        set { stringValue = value; }
    //    }
    //}
}
