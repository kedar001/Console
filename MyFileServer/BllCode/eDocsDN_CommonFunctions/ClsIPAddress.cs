//... Code Added by manav on 21-05-2013 for DRT-4065 ...

using System;
using System.Web;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using System.Net;

using System.Net.Sockets;

namespace eDocsDN_CommonFunctions
{
    public static class ClsIPAddress
    {
        #region .... Variable Declaration ....

        static string _szError;

        #endregion

        #region .... Property ....

        public static string msgError
        {
            get { return _szError; }
            set { _szError = value; }
        }

        #endregion

        #region .... Functions Definition ....

        public static string GetIP4Address()
        {
            msgError = "";
            string IP4Address = String.Empty;
            try
            {

                //.... For-Each block Added for IP bug Fixing (R&D) provided by Kedar DR- on 6/8/2014 by Harshad ....

                //foreach (var item in Dns.GetHostEntry(HttpContext.Current.Request.ServerVariables["REMOTE_ADDR"]).AddressList)
                //{
                //    if (item.AddressFamily == AddressFamily.InterNetwork)
                //    {
                //        IP4Address = item.ToString();
                //        break;
                //    }
                //}

                //if (IP4Address != String.Empty)
                //{
                //    return IP4Address;
                //}

                //........................

                foreach (IPAddress IPA in Dns.GetHostAddresses(HttpContext.Current.Request.UserHostAddress))
                {
                    if (IPA.AddressFamily.ToString() == "InterNetwork")
                    {
                        IP4Address = IPA.ToString();
                        break;
                    }
                }

                if (IP4Address != String.Empty)
                {
                    return IP4Address;
                }

                foreach (IPAddress IPA in Dns.GetHostAddresses(Dns.GetHostName()))
                {
                    if (IPA.AddressFamily.ToString() == "InterNetwork")
                    {
                        IP4Address = IPA.ToString();
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            return IP4Address;
        }

        public static void Dispose()
        {
            _szError = null;
            msgError = null;
        }

        #endregion
    }
}