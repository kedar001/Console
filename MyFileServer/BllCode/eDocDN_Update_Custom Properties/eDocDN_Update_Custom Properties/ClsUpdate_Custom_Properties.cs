using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using eDocsDN_DocX;
using System.IO;


namespace eDocDN_Update_Custom_Properties
{
    public class ClsUpdate_Custom_Properties : clsUpdate_Document_Custom_Properties, IDisposable
    {

        #region .... Variable Declaration ....
        #endregion

        #region .... Constructor ....
        public ClsUpdate_Custom_Properties()
        {
            msgError = "";
        }

        public ClsUpdate_Custom_Properties(string szFilePath)
        {            msgError = "";
            FileName = szFilePath;
            _strmDocument = null;
            lstCustom_Properties = new List<Custom_Property>();
        }
        public ClsUpdate_Custom_Properties(Stream strmDocument)
        {
            msgError = "";
            FileName = string.Empty;
            _strmDocument = strmDocument;
            lstCustom_Properties = new List<Custom_Property>();
        }


        #endregion

        #region .... Properties ....

        public string FileName { get; set; }

        #endregion

        #region .... Public functions ....

        //public Stream Update_Document_Custom_Property(List<Custom_Property> lstCustomProperty)
        //{
        //    eDocsDN_DocX.DocX document = null;
        //    try
        //    {
        //        if (_strmDocument != null)
        //            document = eDocsDN_DocX.DocX.Load(_strmDocument);
        //        else
        //            document = eDocsDN_DocX.DocX.Load(FileName);

        //        foreach (var item in lstCustomProperty)
        //            document.AddCustomProperty(new eDocsDN_DocX.CustomProperty(item.PropertyName, Convert.ToString(item.PropertyValue)));

        //        if (_strmDocument != null)
        //            document.SaveAs(_strmDocument);
        //        else
        //            document.SaveAs(FileName);

        //    }
        //    catch (Exception ex)
        //    {
        //        msgError = ex.Message;
        //    }
        //    finally
        //    {
        //        if (document != null)
        //            document.Dispose();
        //        document = null;
        //        lstCustomProperty = null;
        //    }
        //    return _strmDocument;
        //}

        //...

        public Stream Attach_Document_Custom_Property(List<Custom_Property> lstCustomProperty)
        {
            eDocsDN_DocX.DocX document = null;
            try
            {
                if (_strmDocument != null)
                    document = eDocsDN_DocX.DocX.Load(_strmDocument);
                else
                    document = eDocsDN_DocX.DocX.Load(FileName);

                foreach (var item in lstCustomProperty)
                    document.AddCustomProperty(new eDocsDN_DocX.CustomProperty(item.PropertyName, Convert.ToString(item.PropertyValue)));

                if (_strmDocument != null)
                    document.SaveAs(_strmDocument);
                else
                    document.SaveAs(FileName);

            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            finally
            {
                if (document != null)
                    document.Dispose();
                document = null;
                lstCustomProperty = null;
            }
            return _strmDocument;
        }

        #endregion

        //#region .... IDISPOSABLE ....

        //public void Dispose()
        //{
        //    Dispose(true);
        //    GC.SuppressFinalize(this);
        //}
        //protected virtual void Dispose(bool disposing)
        //{
        //    if (disposing)
        //    {


        //    }
        //    else
        //    {
        //        if (lstCustom_Properties != null)
        //        {
        //            lstCustom_Properties.Clear();
        //            lstCustom_Properties = null;
        //        }
        //    }
        //}

        //~ClsUpdate_Custom_Properties()
        //{
        //    Dispose(false);
        //}


        //#endregion

    }
}
