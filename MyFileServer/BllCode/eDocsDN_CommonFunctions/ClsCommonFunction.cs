using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Security.Cryptography;
using System.Globalization;

namespace eDocsDN_CommonFunctions
{
    public class ClsCommonFunction
    {
        # region .... Variable Declaration ....

        //... Code Added & Changed by manav on 31-07-2013 for DRD-919119 ...
        DateTime _sdtConvertedDate;
        string _szError;

        #endregion

        # region .... Constructor ....

        public ClsCommonFunction()
        {
            msgError = "";
        }

        #endregion

        # region .... Property ....

        public string msgError
        {
            get { return _szError; }
            set { _szError = value; }
        }

        #endregion

        # region .... Functions Definition ....

        public string Encrypt(string szString)
        {
            string szEncString = "";
            for (int i = 0; i < szString.Length; i++)
            {
                szEncString = szEncString + (char)((int)szString[i] + 1);
            }
            return ReverseString(szEncString);
        }

        public string ReverseString(string szString)
        {
            string szReverseString = "";
            for (int i = szString.Length - 1; i >= 0; i--)
            {
                szReverseString = szReverseString + szString[i];
            }
            return szReverseString;
        }

        public string Decrypt(string szString)
        {
            string szResult = "";
            string szTemp = ReverseString(szString);

            for (int i = 0; i < szTemp.Length; i++)
            {
                szResult = szResult + (char)((int)szTemp[i] - 1);
            }
            return szResult;
        }

        public string getCurrentDate()
        {
            //return (string.Format("{0:yyyy/MM/dd}", DateTime.Now));
            return string.Format("{0:MM/dd/yyyy}", DateTime.Now);
        }

        public string getCurrentTime()
        {
            return string.Format("{0:hh:mm:ss tt}", DateTime.Now);
            //return string.Format("{0:hh:mm:ss}", DateTime.Now);
        }

        public string getCurrentDateTime()
        {
            return string.Format("{0:MM/dd/yyyy hh:mm:ss tt}", DateTime.Now);
        }

        public string getMd5Hash(string szFilePath)
        {
            msgError = "";

            // Create a new instance of the MD5CryptoServiceProvider object.
            MD5 md5Hasher = null;
            StreamReader srReader = null;
            StringBuilder strBuilder = null;
            try
            {
                md5Hasher = MD5.Create();
                srReader = new StreamReader(szFilePath);

                // compute the hash for given file.
                byte[] data = md5Hasher.ComputeHash(srReader.BaseStream);
                srReader.Close();
                srReader.Dispose();
                srReader = null;
                md5Hasher.Clear();
                md5Hasher = null;

                // Create a new Stringbuilder to collect the bytes
                // and create a string.
                strBuilder = new StringBuilder();

                // Loop through each byte of the hashed data 
                // and format each one as a hexadecimal string.
                for (int i = 0; i < data.Length; i++)
                {
                    strBuilder.Append(data[i].ToString("x2"));
                }
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            finally
            {
                md5Hasher = null;
                if (srReader != null)
                {
                    srReader.Close();
                    srReader.Dispose();
                }
                srReader = null;
            }
            // Return the hexadecimal string.
            return strBuilder.ToString();
        }

        #region ... DateTime Functions ...

        public string GetServerDateSeparator()
        {
            msgError = "";
            return CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator;
        }

        public string GetServerDateFormat()
        {
            msgError = "";
            DateTimeFormatInfo SysDateFormatInfo = null;
            string szSysDateFormat = "";
            try
            {
                SysDateFormatInfo = CultureInfo.CurrentCulture.DateTimeFormat;
                szSysDateFormat = SysDateFormatInfo.ShortDatePattern;
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            finally
            {
                SysDateFormatInfo = null;
            }
            return szSysDateFormat;
        }

        public string GetServerDateFormat(out string szSysDateFormat, out string szDateSeparator)
        {
            msgError = "";
            szSysDateFormat = "";
            szDateSeparator = "";
            DateTimeFormatInfo SysDateFormatInfo = null;
            try
            {
                SysDateFormatInfo = CultureInfo.CurrentCulture.DateTimeFormat;
                szSysDateFormat = SysDateFormatInfo.ShortDatePattern;
                szDateSeparator = SysDateFormatInfo.DateSeparator;
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            finally
            {
                SysDateFormatInfo = null;
            }
            return szSysDateFormat;
        }

        public string GetServerDateTimeFormat()
        {
            msgError = "";
            DateTimeFormatInfo SysDateFormatInfo = null;
            string szSysDateFormat = "";
            try
            {
                SysDateFormatInfo = CultureInfo.CurrentCulture.DateTimeFormat;
                szSysDateFormat = SysDateFormatInfo.ShortDatePattern + " " + SysDateFormatInfo.LongTimePattern;
            }
            catch (Exception ex)
            {
                msgError = ex.Message;
            }
            finally
            {
                SysDateFormatInfo = null;
            }
            return szSysDateFormat;
        }

        public DateTimeFormatInfo GetServerDateTimeFormatInfo()
        {
            msgError = "";
            return CultureInfo.CurrentCulture.DateTimeFormat;
        }

        public DateTime GetDateFromDatabaseDate(string szDateFromDatabase)
        {
            msgError = "";
            CultureInfo CultureProvider = null;
            DateTimeFormatInfo SysDateFormatInfo = null;
            string szSysDateFormat = "";
            string szDateFormat = "";
            _sdtConvertedDate = new DateTime();
            try
            {
                CheckDate(szDateFromDatabase);
                szDateFromDatabase = szDateFromDatabase.Trim();

                CultureProvider = new CultureInfo(CultureInfo.CurrentCulture.Name);
                SysDateFormatInfo = CultureInfo.CurrentCulture.DateTimeFormat;
                szSysDateFormat = SysDateFormatInfo.ShortDatePattern;

                if (szDateFromDatabase.Trim().Contains(" "))
                    szSysDateFormat = SysDateFormatInfo.ShortDatePattern + " " + SysDateFormatInfo.LongTimePattern;

                SysDateFormatInfo = null;

                if (DateTime.TryParseExact(szDateFromDatabase, szSysDateFormat, CultureProvider, DateTimeStyles.None, out  _sdtConvertedDate))
                {
                    szDateFormat = szSysDateFormat;
                }
                else if (DateTime.TryParseExact(szDateFromDatabase, szSysDateFormat, CultureInfo.CreateSpecificCulture("en-US"), DateTimeStyles.None, out  _sdtConvertedDate))
                {
                    szDateFormat = szSysDateFormat;
                }
                else if (DateTime.TryParse(szDateFromDatabase, CultureProvider, System.Globalization.DateTimeStyles.None, out  _sdtConvertedDate))
                {
                    szDateFormat = szSysDateFormat;
                }
                else if (DateTime.TryParse(szDateFromDatabase, CultureInfo.CreateSpecificCulture("en-US"), DateTimeStyles.None, out  _sdtConvertedDate))
                {
                    szDateFormat = szSysDateFormat;
                }
                else
                {
                    _sdtConvertedDate = Convert_UserDateToDateTime(szDateFromDatabase, szSysDateFormat);
                }
            }
            catch (Exception ex)
            {
                msgError = "Error while Converting Date(" + Convert.ToString(szDateFromDatabase) + "): " + ex.Message;
            }
            finally
            {
                szDateFormat = null;
                szSysDateFormat = null;
                SysDateFormatInfo = null;
                CultureProvider = null;
            }
            return _sdtConvertedDate;
        }

        public DateTime ConvertToDateTime(string szUserDate)
        {
            msgError = "";
            return ConvertToDateTime(szUserDate, CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern + " " + CultureInfo.CurrentCulture.DateTimeFormat.LongTimePattern);
        }

        public DateTime ConvertToDateTime(string szUserDateTime, string szUserDateTimeFormat)
        {
            msgError = "";
            CultureInfo CultureProvider = null;
            DateTimeFormatInfo SysDateFormatInfo = null;
            DateTime stcDateTime = new DateTime();
            string szSysDateFormat = "";
            string szSysDateSeparator = "";
            string szDateFormat = "";
            string szTimeFromat = "";
            try
            {
                CheckDate(szUserDateTime);

                szUserDateTime = szUserDateTime.Trim();
                szUserDateTimeFormat = Convert.ToString(szUserDateTimeFormat).Trim();

                //CultureInfo CultureProvider = new CultureInfo("en-US");
                CultureProvider = new CultureInfo(CultureInfo.CurrentCulture.Name);
                SysDateFormatInfo = CultureInfo.CurrentCulture.DateTimeFormat;
                szSysDateSeparator = SysDateFormatInfo.DateSeparator;
                szSysDateFormat = SysDateFormatInfo.ShortDatePattern;

                if (szUserDateTime.Contains(" "))
                {
                    szSysDateFormat = SysDateFormatInfo.ShortDatePattern + " " + SysDateFormatInfo.LongTimePattern;
                    if (szUserDateTimeFormat.Trim().Contains(" "))
                    {
                        szTimeFromat = szUserDateTimeFormat.Substring(szUserDateTimeFormat.IndexOf(' '));
                    }
                    else
                    {
                        szTimeFromat = " hh:mm:ss tt";
                        szUserDateTimeFormat = szUserDateTimeFormat + szTimeFromat;
                    }
                }
                else
                {
                    if (szUserDateTimeFormat.Trim().Contains(" "))
                        szUserDateTimeFormat = szUserDateTimeFormat.Substring(0, szUserDateTimeFormat.IndexOf(' '));
                }
                SysDateFormatInfo = null;
                //....

                if (DateTime.TryParseExact(szUserDateTime, szUserDateTimeFormat, CultureProvider, DateTimeStyles.None, out  stcDateTime))
                {
                    szDateFormat = szUserDateTimeFormat;
                }
                else if (DateTime.TryParseExact(szUserDateTime, szUserDateTimeFormat, CultureInfo.CreateSpecificCulture("en-US"), DateTimeStyles.None, out  stcDateTime))
                {
                    szDateFormat = szUserDateTimeFormat;
                }
                else if (DateTime.TryParseExact(szUserDateTime, szSysDateFormat, CultureProvider, DateTimeStyles.None, out  stcDateTime))
                {
                    szDateFormat = szSysDateFormat;
                }
                else if (DateTime.TryParseExact(szUserDateTime, szSysDateFormat, CultureInfo.CreateSpecificCulture("en-US"), DateTimeStyles.None, out  stcDateTime))
                {
                    szDateFormat = szSysDateFormat;
                }
                else if (DateTime.TryParseExact(szUserDateTime, CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern, CultureProvider, System.Globalization.DateTimeStyles.None, out  stcDateTime))
                {
                    szDateFormat = szSysDateFormat;
                }
                else if (DateTime.TryParse(szUserDateTime, CultureProvider, System.Globalization.DateTimeStyles.None, out  stcDateTime))
                {
                    szDateFormat = szSysDateFormat;
                }
                else if (DateTime.TryParse(szUserDateTime, CultureInfo.CreateSpecificCulture("en-US"), DateTimeStyles.None, out  stcDateTime))
                {
                    szDateFormat = szSysDateFormat;
                }
                else if (DateTime.TryParseExact(szUserDateTime, "M/d/yyyy" + szTimeFromat, CultureProvider, DateTimeStyles.None, out  stcDateTime))
                {
                    szDateFormat = "M/d/yyyy" + szTimeFromat;
                }
                else if (DateTime.TryParseExact(szUserDateTime, "M/d/yyyy" + szTimeFromat, CultureInfo.CreateSpecificCulture("en-US"), DateTimeStyles.None, out  stcDateTime))
                {
                    szDateFormat = "M/d/yyyy" + szTimeFromat;
                }
                else if (DateTime.TryParseExact(szUserDateTime, "dd/MM/yyyy" + szTimeFromat, CultureProvider, DateTimeStyles.None, out  stcDateTime))
                {
                    szDateFormat = "dd/MM/yyyy" + szTimeFromat;
                }
                else if (DateTime.TryParseExact(szUserDateTime, "dd.MM.yyyy" + szTimeFromat, CultureProvider, DateTimeStyles.None, out stcDateTime))
                {
                    szDateFormat = "dd.MM.yyyy" + szTimeFromat;
                }
                else if (DateTime.TryParseExact(szUserDateTime, "dd-MM-yyyy" + szTimeFromat, CultureProvider, DateTimeStyles.None, out stcDateTime))
                {
                    szDateFormat = "dd-MM-yyyy" + szTimeFromat;
                }
                else if (DateTime.TryParseExact(szUserDateTime, "MM/dd/yyyy" + szTimeFromat, CultureProvider, DateTimeStyles.None, out  stcDateTime))
                {
                    szDateFormat = "MM/dd/yyyy" + szTimeFromat;
                }
                else if (DateTime.TryParseExact(szUserDateTime, "MM.dd.yyyy" + szTimeFromat, CultureProvider, DateTimeStyles.None, out stcDateTime))
                {
                    szDateFormat = "MM.dd.yyyy" + szTimeFromat;
                }
                else if (DateTime.TryParseExact(szUserDateTime, "MM-dd-yyyy" + szTimeFromat, CultureProvider, DateTimeStyles.None, out stcDateTime))
                {
                    szDateFormat = "MM-dd-yyyy" + szTimeFromat;
                }
                else if (DateTime.TryParseExact(szUserDateTime, "yyyy/MM/dd" + szTimeFromat, CultureProvider, DateTimeStyles.None, out  stcDateTime))
                {
                    szDateFormat = "yyyy/MM/dd" + szTimeFromat;
                }
                else if (DateTime.TryParseExact(szUserDateTime, "yyyy.MM.dd" + szTimeFromat, CultureProvider, DateTimeStyles.None, out stcDateTime))
                {
                    szDateFormat = "yyyy.MM.dd" + szTimeFromat;
                }
                else if (DateTime.TryParseExact(szUserDateTime, "yyyy-MM-dd" + szTimeFromat, CultureProvider, DateTimeStyles.None, out stcDateTime))
                {
                    szDateFormat = "yyyy-MM-dd" + szTimeFromat;
                }
                else
                {
                    stcDateTime = Convert_UserDateToDateTime(szUserDateTime, szUserDateTimeFormat);
                }
            }
            catch (Exception ex)
            {
                msgError = "Error while Converting Date(" + Convert.ToString(szUserDateTime) + "): " + ex.Message;
            }
            finally
            {
                CultureProvider = null;
                SysDateFormatInfo = null;
                szSysDateFormat = null;
                szSysDateSeparator = null;
                szDateFormat = null;
            }
            return stcDateTime;
        }

        private DateTime Convert_UserDateToDateTime(string szUserDateTime, string szUserDateTimeFormat)
        {
            bool bStart = false;
            string szTime = "";
            string szAMDesignator = "AM";
            string szTimeFormat = "";
            string[] szDateInfo = null;
            string[] szDateFormatInfo = null;
            string[] szTimeInfo = null;
            string[] szTimeFormatInfo = null;
            DateTime stcDateTime = new DateTime();
            try
            {
                #region ... Set Date Time Formats ...

                szUserDateTime = szUserDateTime.Trim();
                if (szUserDateTime.Contains(" "))
                {
                    string[] arrDateInfo = szUserDateTime.Split(' ');
                    szUserDateTime = arrDateInfo[0];
                    szTime = arrDateInfo[1];
                    if (arrDateInfo.Length > 2)
                        szAMDesignator = arrDateInfo[2].Trim().ToUpper();

                    arrDateInfo = null;
                }

                if (szUserDateTimeFormat.Contains(" "))
                {
                    string[] arrTimeInfo = szUserDateTimeFormat.Split(' ');
                    szUserDateTimeFormat = arrTimeInfo[0];
                    szTimeFormat = arrTimeInfo[1];
                    arrTimeInfo = null;
                }

                #endregion

                #region ... Get Date Parameters ...

                //... Default DateTime Format : M/d/yyyy ....
                for (int iIndex = 0; iIndex < szUserDateTime.Length; iIndex++)
                {
                    if (!bStart)
                    {
                        switch (szUserDateTime[iIndex])
                        {
                            case '/':
                                {
                                    bStart = true;
                                    szDateInfo = szUserDateTime.Split('/');
                                    szDateFormatInfo = szUserDateTimeFormat.Split('/');
                                }
                                break;
                            case '.':
                                {
                                    bStart = true;
                                    szDateInfo = szUserDateTime.Split('.');
                                    szDateFormatInfo = szUserDateTimeFormat.Split('.');
                                }
                                break;
                            case '-':
                                {
                                    bStart = true;
                                    szDateInfo = szUserDateTime.Split('-');
                                    szDateFormatInfo = szUserDateTimeFormat.Split('-');
                                }
                                break;
                        }
                    }
                    else
                    {
                        break;
                    }
                }

                if (szDateFormatInfo == null)
                    throw new Exception("Invalid Date format !!!");

                int iDay = 0, iMonth = 0, iYear = 0;
                for (int iIndex = 0; iIndex < szDateFormatInfo.Length; iIndex++)
                {
                    if (szDateFormatInfo.GetValue(iIndex).ToString().ToLower().StartsWith("d"))
                    {
                        iDay = Convert.ToInt32(szDateInfo.GetValue(iIndex));
                    }
                    else if (szDateFormatInfo.GetValue(iIndex).ToString().ToLower().StartsWith("m"))
                    {
                        iMonth = Convert.ToInt32(szDateInfo.GetValue(iIndex));
                    }
                    else if (szDateFormatInfo.GetValue(iIndex).ToString().ToLower().StartsWith("y"))
                    {
                        iYear = Convert.ToInt32(szDateInfo.GetValue(iIndex));
                    }
                }

                if (iMonth > 12 && iDay < 13)
                {
                    int iTempDay = iMonth;
                    iMonth = iDay;
                    iDay = iTempDay;
                }

                #endregion

                #region ... Get Time Parameters ...

                int iHour = 0, iMinute = 0, iSecond = 0;
                if (szTime != "")
                {
                    szTimeInfo = szTime.Split(':');
                    szTimeFormatInfo = szTimeFormat.Split(':');
                    for (int iIndex = 0; iIndex < szTimeFormatInfo.Length; iIndex++)
                    {
                        if (szTimeInfo.Length > iIndex)
                        {
                            switch (szTimeFormatInfo[iIndex])
                            {
                                case "h":
                                case "hh":
                                case "H":
                                case "HH":
                                    {
                                        iHour = Convert.ToInt32(szTimeInfo[iIndex]);
                                        if (szAMDesignator.ToUpper().Trim() == "PM")
                                            if (iHour < 13) iHour = iHour + 12;
                                    }
                                    break;
                                case "m":
                                case "mm":
                                    {
                                        iMinute = Convert.ToInt32(szTimeInfo[iIndex]);
                                    }
                                    break;
                                case "s":
                                case "ss":
                                    {
                                        if (szTimeInfo[iIndex].Contains("."))
                                            iSecond = Convert.ToInt32(szTimeInfo[iIndex].Split('.')[0]);
                                        else
                                            iSecond = Convert.ToInt32(szTimeInfo[iIndex]);
                                    }
                                    break;
                            }
                        }
                    }
                }

                #endregion

                stcDateTime = new DateTime(iYear, iMonth, iDay, iHour, iMinute, iSecond);
            }
            finally
            {
                szAMDesignator = null;
                szTime = null;
                szTimeFormat = null;
                szTimeInfo = null;
                szTimeFormatInfo = null;
                szDateInfo = null;
                szDateFormatInfo = null;
            }
            return stcDateTime;
        }

        private void CheckDate(string szUserDate)
        {
            if (string.IsNullOrEmpty(szUserDate) || szUserDate.Trim() == "")
                throw new Exception("Date should not be null/empty !!");
        }

        public string ConvertDateToUserFormat(DateTime sdtUsertDate, string szDateTimeFormat)
        {
            msgError = "";
            string szFormattedUserDate = "";
            //DateTimeFormatInfo SysDateFormatInfo = null;
            //string[] szDateFormatInfo = null;
            try
            {
                if (string.IsNullOrEmpty(szDateTimeFormat) || szDateTimeFormat.Trim() == "")
                    throw new Exception("Date Format (" + Convert.ToString(szDateTimeFormat) + ") should not be null/empty !!");

                szDateTimeFormat = szDateTimeFormat.Trim();
                szFormattedUserDate = string.Format("{0:" + szDateTimeFormat + "}", sdtUsertDate);

                //-------- OR ---------
                //szFormattedUserDate = sdtUsertDate.ToString(szDateTimeFormat);

                //-------- OR ---------
                //SysDateFormatInfo = new DateTimeFormatInfo();
                //SysDateFormatInfo.AMDesignator = "tt";
                //if (szDateTimeFormat.Contains(" "))
                //{
                //    szDateFormatInfo = szDateTimeFormat.Split(' ');
                //    //SysDateFormatInfo.DateSeparator = "/";
                //    SysDateFormatInfo.ShortDatePattern = szDateFormatInfo[0];
                //    SysDateFormatInfo.LongTimePattern = szDateFormatInfo[1];
                //}
                //else
                //{
                //    SysDateFormatInfo.ShortDatePattern = szDateTimeFormat;
                //}
                //szFormattedUserDate = string.Format("{0:" + SysDateFormatInfo.ShortDatePattern + " " + SysDateFormatInfo.LongTimePattern + " " + SysDateFormatInfo.AMDesignator + "}", sdtUsertDate);
                //SysDateFormatInfo = null;
            }
            catch (Exception ex)
            {
                msgError = "Error while Converting Date (" + sdtUsertDate.ToString("MM/dd/yyyy hh:mm:ss tt") + ") into user format (" + szDateTimeFormat + "): " + ex.Message;
                szFormattedUserDate = "";
            }
            finally
            {
                //szDateFormatInfo = null;
                //SysDateFormatInfo = null;
            }
            return szFormattedUserDate;
        }

        #endregion

        #endregion
    }
}