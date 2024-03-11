using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestConsole
{

    public enum eLockUnlock
    {
        NoLcock = 0,
        Comments,
        TrackChanges,
        Form,
        ReadOnly,
    }

    public enum eTrackChanges
    {
        On = 0,
        Off

    }
    public enum ePrintFormData
    {
        On = 0,
        Off

    }


    public class ClsFileOperations
    {
        ILockUnlock _LockUnlock;
        ITrackChanges _ITrackChanges;
        IUpdateScanSignature _IUpdateScanSignature;
        ISetPrintFormData _ISetPrintFormData;
        public ClsFileOperations(ILockUnlock LockUnlock,
                                    ITrackChanges ITrackChanges,
                                    IUpdateScanSignature IUpdateScanSignature,
                                    ISetPrintFormData ISetPrintFormData)
        {
            _LockUnlock = LockUnlock;
            _ITrackChanges = ITrackChanges;
            _IUpdateScanSignature = IUpdateScanSignature;
            _ISetPrintFormData = ISetPrintFormData;
        }
        public void Process_Documents(Queue<string> Operations)
        {
            foreach (string item in Operations)
            {
                switch (item)
                {
                    case "LockUnlock":
                        _LockUnlock.LockUnclockOperation(eLockUnlock.Comments);
                        break;
                    default:
                        break;
                }
                Console.WriteLine(item);
            }
        }
    }


    public class clsLockUnlock : ILockUnlock
    {
        public void LockUnclockOperation(eLockUnlock e)
        {
            Console.WriteLine(e);
        }
    }
    public class clsSet_Track_Changes : ITrackChanges
    {
        public void SetTrackChanges(eTrackChanges e)
        {
            Console.WriteLine(e);
        }
    }

    public class clsUpdate_ScanSignature : IUpdateScanSignature
    {
        public void UpdateScanSignature()
        {
            Console.WriteLine("Update Scan Signature");
        }
    }

    public class clsPrint_Form_Data : ISetPrintFormData
    {
        public void SetPrintFormData(ePrintFormData e)
        {
            Console.WriteLine("Print Form Data : " + e);
        }
    }



    public interface ILockUnlock
    {
        void LockUnclockOperation(eLockUnlock e);
    }
    public interface ITrackChanges
    {
        void SetTrackChanges(eTrackChanges e);
    }

    public interface IUpdateScanSignature
    {
        void UpdateScanSignature();
    }
    public interface ISetPrintFormData
    {
        void SetPrintFormData(ePrintFormData e);
    }

    public interface IIsValidProcess
    {
        bool IsValid();
    }


}
