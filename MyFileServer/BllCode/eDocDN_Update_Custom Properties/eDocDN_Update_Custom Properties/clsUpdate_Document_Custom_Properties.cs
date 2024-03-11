using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace eDocDN_Update_Custom_Properties
{
    public class clsUpdate_Document_Custom_Properties : IDisposable
    {
        #region .... Variable Declaration ....
        public Stream _strmDocument = null;

        #endregion

        #region .... Property ...
        public string msgError { get; set; }
        public List<Custom_Property> lstCustom_Properties { get; set; }
        #endregion

        #region .... Constructor ....
        public clsUpdate_Document_Custom_Properties()
        {
            msgError = "";
            lstCustom_Properties = new List<Custom_Property>();
        }
        #endregion

        #region .... Public functions ....
        public Stream Update_Document_Custom_Property(Stream _strmDocument, List<Custom_Property> lstCustomProperty)
        {
            eDocsDN_DocX.DocX document = null;
            try
            {
                document = eDocsDN_DocX.DocX.Load(_strmDocument);
                foreach (var item in lstCustomProperty)
                    document.AddCustomProperty(new eDocsDN_DocX.CustomProperty(item.PropertyName, Convert.ToString(item.PropertyValue)));
            }
            catch (Exception ex)
            {
                _strmDocument = null;
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
        public Stream Update_Document_Custom_Property_In_Document(Stream _strmDocument, List<Custom_Property> lstCustomProperty)
        {
            eDocsDN_DocX_IS.DocX document = null;
            try
            {
                document = eDocsDN_DocX_IS.DocX.Load(_strmDocument);
                foreach (var item in lstCustomProperty)
                    document.AddCustomProperty(new eDocsDN_DocX_IS.CustomProperty(item.PropertyName, Convert.ToString(item.PropertyValue)));
            }
            catch (Exception ex)
            {
                _strmDocument = null;
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
                if (lstCustom_Properties != null)
                    lstCustom_Properties.Clear();
                lstCustom_Properties = null;

            }
            else
            {

            }
        }

        ~clsUpdate_Document_Custom_Properties()
        {
            Dispose(false);
        }


        #endregion
    }
}
