using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.DirectoryServices;
using System.Security.Principal;
using Microsoft.VisualBasic.FileIO;
using System.Data.OleDb;
using System.Globalization;
using System.ComponentModel;
using System.Net.Mail;
using System.Net;

namespace SKF.RS.STB.Generic
{

    public static class Debug
    {
        public static bool DebugMode { get; set; }

        //private static Error _LastError;
        public static Error LastError { get; set; }

        public static void Log() { GenericTools.DebugMsg(LastError.Message); }
        public static void Log(string Message) { GenericTools.DebugMsg(Message); }
    }

    public class Error
    {
        public uint Code = 0;
        public uint Type = 0;
        public string Message = string.Empty;
        public object Source = null;
        public Exception Exception = null;

        public Error()
        {
            Clear();
        }
        public void Clear()
        {
            Code = 0;
            Type = 0;
            Source = null;
            Exception = null;
            Message = string.Empty;
        }
    }

    public static class GenericTools
    {
        public static bool Interactive = false;
        public static bool Log = true;
        public static bool Debug = false;
        private static string _LogFileName = string.Empty;
        public static string LogFileName 
        {
            get { return (string.IsNullOrWhiteSpace(_LogFileName) ? GetAuxFileName(".log") : _LogFileName); }
            set { _LogFileName = value; }
        }
        public static Error LastError = new Error();


        public static bool SendEmail(string[] Emails, string Subject, StringBuilder Body)
        {
            bool _return = false;

            MailMessage Email = new MailMessage();

            try
            {
                foreach (var s in Emails)
                {
                    Email.To.Add(s);
                }

                if (Emails.Length > 0)
                {
                    DebugMsg("Sending Email...");

                    Email.Subject = Subject + " - " + System.DateTime.Now.ToString();
                    Email.From = new MailAddress("no-reply@cdr-skf.com.br");
                    SmtpClient SMTPConn = new SmtpClient("SKF-WEB01", 25);
                    SMTPConn.Credentials = new NetworkCredential("cdr_services@cdr.skf", "RDCWorld2011");

                    AlternateView HTMLTextView = AlternateView.CreateAlternateViewFromString(Body.ToString(), null, "text/html");

                    Email.AlternateViews.Add(HTMLTextView);
                    try
                    {
                        SMTPConn.Send(Email);
                        DebugMsg("Email sent");
                        _return = true;
                    }
                    catch (Exception ex)
                    {
                        DebugMsg("Erro: " + ex.Message);

                        _return = false;
                    }

                }
            }
            catch (Exception ex)
            {
                DebugMsg("Erro: " + ex.Message);
            }

            return _return;
        }
        public static DataTable GetDataTableFromCsv(string path, bool isFirstRowHeader, string Delimiter = ",")
        {
            DataTable csvData = new DataTable();
            try
            {
                using (TextFieldParser csvReader = new TextFieldParser(path))
                 {
                     csvReader.SetDelimiters(new string[] { Delimiter });
                    csvReader.HasFieldsEnclosedInQuotes = true;
                    string[] colFields = csvReader.ReadFields();
                    int x = 0;
                    foreach (string column in colFields)
                    {
                        DataColumn datecolumn = new DataColumn(column);
                        datecolumn.AllowDBNull = true;
                        try
                        {
                            //csvData.Columns.Add(datecolumn);
                            csvData.Columns.Add("Coluna " + x);
                            x++;
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Erro Colocando Coluna");
                        }
                    }
                    csvData.Rows.Add(colFields);

                    while (!csvReader.EndOfData)
                    {
                        string[] fieldData = csvReader.ReadFields();
                        //Making empty value as null
                        for (int i = 0; i < fieldData.Length; i++)
                        {
                            try
                            {
                                if (fieldData[i] == "")
                                {
                                    fieldData[i] = null;
                                }
                            }
                            catch(Exception ex)
                            {
                                Console.WriteLine("Erro Colocando Coluna");
                            }
                        }
                        csvData.Rows.Add(fieldData);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERRO: " + ex.Message.ToString());
            }
            return csvData;

        }
   
        public static DataTable GetDataTableFromList<T>(IList<T> data)
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
            {
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            }
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                {
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                }
                table.Rows.Add(row);
            }
            return table;
        }
        /*
        public Generic()
        {
            GenericTools_Initiate(Log);
        }
        public Generic(bool bLog)
        {
            GenericTools_Initiate(bLog);
        }
        
        private void GenericTools_Initiate(bool bLog)
        {
            Log = bLog;
            LogFileName = GetAuxFileName(".log");
            if (Log)
            {
                WriteLog("");
                WriteLog("");
                WriteLog(Application.ProductName);
                WriteLog("Version: " + Application.ProductVersion);
                WriteLog("Computer name: " + Environment.MachineName);
                WriteLog("User name: " + Environment.UserDomainName + "\\" + Environment.UserName);
                WriteLog("Executable: " + Application.ExecutablePath);
                WriteLog("Command line: " + Environment.CommandLine);
                WriteLog("Current directory: " + Environment.CurrentDirectory);
                WriteLog("OS version: " + Environment.OSVersion);
            }
        }
        */


        public static uint ISOWeek { get { return Convert.ToUInt32(System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(System.DateTime.Now, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)); } }
        /// <summary>
        /// Return the ISO week number for current date.
        /// </summary>
        /// <returns>ISO week number</returns>
        public static int GetWeekOfYear() { return GetWeekOfYear(System.DateTime.Now); }
        /// <summary>
        /// Return the ISO week number for given date.
        /// </summary>
        /// <returns>ISO week number</returns>
        public static int GetWeekOfYear(DateTime dDateTime) { return System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(dDateTime, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday); }

        public static string WindowsGetUserFullName { get { return System.Security.Principal.WindowsIdentity.GetCurrent().Name; } }

        public static string WindowsGetUserName { get { return WindowsGetUserFullName.Split('\\')[1]; } }

        public static string WindowsGetUserDomainName { get { return WindowsGetUserFullName.Split('\\')[0]; } }

        public static string WindowsGetHostName { get { return System.Net.Dns.GetHostName(); } }

        public static string GetAuxFileName(string Extension)
        {
            if (!Extension.StartsWith(".")) Extension = "." + Extension;
            return Application.ExecutablePath.Replace(".Exe", Extension).Replace(".ExE", Extension).Replace(".EXe", Extension).Replace(".EXE", Extension).Replace(".exe", Extension).Replace(".exE", Extension).Replace(".eXe", Extension).Replace(".eXE", Extension);
        }
        public static DataTable SortDataTable(DataTable _Datatable, string Column, bool Ascendent = true)
        {
            DataTable sortedDT = new DataTable();
            if (_Datatable.Rows.Count > 0)
            {
                string By;
                if (Ascendent == true)
                    By = "ASC";
                else
                    By = "DESC";


                DataView dv = _Datatable.DefaultView;
                dv.Sort = Column + " " + By;
                sortedDT = dv.ToTable();
            }
            else
            {
                sortedDT = _Datatable;
            }
             return sortedDT;
        }


        private static string StdDateTime()
        {
            return StdDateTime(System.DateTime.Now);
        }
        private static string StdDateTime(DateTime oDateTime)
        {
            return oDateTime.ToString("yyyy/MM/dd HH:mm:ss.fff");
        }

        public static DateTime StrToDateTime(string sDateTime)
        {
            GenericTools.DebugMsg("StrToDateTime(" + sDateTime + "): Starting...");

            DateTime ReturnValue = new DateTime();

            try
            {
                if (sDateTime.Length == 14)
                {
                    int Year = Convert.ToInt16(sDateTime.Substring(0, 4));
                    int Month = Convert.ToInt16(sDateTime.Substring(4, 2));
                    int Day = Convert.ToInt16(sDateTime.Substring(6, 2));
                    int Hours = Convert.ToInt16(sDateTime.Substring(8, 2));
                    int Minutes = Convert.ToInt16(sDateTime.Substring(10, 2));
                    int Secconds = Convert.ToInt16(sDateTime.Substring(12, 2));
                    ReturnValue = new DateTime(Year, Month, Day, Hours, Minutes, Secconds);
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("StrToDateTime(" + sDateTime + ") error: " + ex.Message);
            }

            GenericTools.DebugMsg("StrToDateTime(" + sDateTime + "): " + ReturnValue.ToString()); 
            return ReturnValue;
        }




        public static void WriteLog(string sLogFileName, string sErrMsg)
        {
            try
            {
                StreamWriter sw = new StreamWriter(sLogFileName, true);
                sw.WriteLine(StdDateTime() + "\t" + Environment.UserDomainName + "\\" + Environment.UserName + "\t" + sErrMsg.Replace(Environment.NewLine, "\t"));
                //sw.WriteLine(StdDateTime() + "\t" + Environment.UserDomainName + "\\" + Environment.UserName + "\t" + cpuCounter.NextValue() + "%" + "\t" + sErrMsg.Replace(Environment.NewLine, "\t"));
                sw.Flush();
                sw.Close();
            }
            catch { }
            if (Interactive & Debug) MessageBox.Show(sErrMsg.Replace("\t", Environment.NewLine), "Debug", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
        }
        public static void WriteLog(string sErrMsg)
        {
            WriteLog(LogFileName, sErrMsg);
            Console.WriteLine(sErrMsg);
        }

        public static void GetError(string sErrMsg)
        {
            if (Log) WriteLog(sErrMsg);
            if (Interactive) MessageBox.Show(sErrMsg.Replace("\t", Environment.NewLine), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            LastError.Message = sErrMsg;
            
        }

        public static void DebugMsg(string DebugMessage)
        {
            if (Debug) WriteLog(DebugMessage);
        }
        

        public static Int32 __iif(bool Condition, Int32 TrueValue, Int32 FalseValue) { return (Condition ? TrueValue : FalseValue); }
        public static Int32 __iif(bool Condition, Int32 TrueValue)  { return (Condition ? TrueValue : int.MinValue); }
        public static float __iif(bool Condition, float TrueValue, float FalseValue) { return (Condition ? TrueValue : FalseValue); }
        public static float __iif(bool Condition, float TrueValue) { return (Condition ? TrueValue : float.NaN); }
        public static string __iif(bool Condition, string TrueValue, string FalseValue) { return (Condition ? TrueValue : FalseValue); }
        public static string __iif(bool Condition, string TrueValue) { return (Condition ? TrueValue : string.Empty); }

        public static Int32 StrToInt32(string sStrToInt_String) { return Convert.ToInt32(sStrToInt_String); }
        public static int StrToInt(string sStrToInt_String) { return (int)Convert.ToInt32(sStrToInt_String); }

        public static float StrToFloat(string sStrToFloat_String)
        {
            try
            {
                if (sStrToFloat_String.Length == 0) return 0;

                float nStrToFloat_Number = 0;

                if (float.TryParse(sStrToFloat_String, out nStrToFloat_Number)) return nStrToFloat_Number;

                int i = 0;

                while (i < (sStrToFloat_String.Length))
                {
                    if (sStrToFloat_String.Length <= 1)
                    {
                        if (float.TryParse(sStrToFloat_String, out nStrToFloat_Number)) return nStrToFloat_Number;
                        return 0;
                    }

                    if (char.IsNumber(sStrToFloat_String, i))
                    {
                        i++;
                    }
                    else if (sStrToFloat_String[i].ToString() == Microsoft.Win32.Registry.CurrentUser.OpenSubKey("HKEY_CURRENT_USER\\Control Panel\\International").GetValue("sDecimal").ToString())
                    {
                        i++;
                    }
                    else
                    {
                        sStrToFloat_String = sStrToFloat_String.Remove(i, 1);
                    }
                    if (float.TryParse(sStrToFloat_String, out nStrToFloat_Number)) return nStrToFloat_Number;
                }
            }
            catch
            {
            }
            try
            {
                return (float)Convert.ToDouble(sStrToFloat_String);
            }
            catch
            {
            }
            return float.NaN;
        }

        public static int Hex2ToValue(string sHexToValue_HexValue)
        {
            DateTime StartTime = System.DateTime.Now;
            int ReturnValue = 0;

            if (sHexToValue_HexValue.Length == 4)
            {
                try
                {
                    string sFirstByte = sHexToValue_HexValue.Substring(0, 2);
                    string sSecondByte = sHexToValue_HexValue.Substring(2, 2);
                    int nFirstByte = Convert.ToInt16(sFirstByte, 16);
                    int nSecondByte = Convert.ToInt16(sSecondByte, 16);
                    if (nSecondByte > 127)
                    {
                        nSecondByte = nSecondByte - 256;
                    }
                    ReturnValue = (int)((nSecondByte * 256) + nFirstByte);
                }
                catch (Exception ex)
                {
                    GetError(ex.Message);
                }
            }
            DebugMsg("Hex2ToValue(" + sHexToValue_HexValue + "): " + ReturnValue.ToString() + " - " + (System.DateTime.Now - StartTime).ToString());
            return ReturnValue;
        }

        public static float Hex8ToValue(string sHexToValue_HexValue)
        {
            DateTime StartTime = System.DateTime.Now;
            float ReturnValue = float.NaN;
            if (sHexToValue_HexValue.Length == 16)
            {
                try
                {
                    ReturnValue = (float)BitConverter.Int64BitsToDouble(Convert.ToInt64(sHexToValue_HexValue, 16));
                }
                catch (Exception ex)
                {
                    GetError(ex.Message);
                }
            }
            DebugMsg("Hex8ToValue(" + sHexToValue_HexValue + "): " + ReturnValue.ToString() + " - " + (System.DateTime.Now - StartTime).ToString());
            return ReturnValue;
        }

        public static string ReverseHex(string sReverseHex_HexValue)
        {
            DateTime StartTime = System.DateTime.Now;
            string sReverseHex_OriginalHexValue = sReverseHex_HexValue;

            for (int i = 0; i <= (sReverseHex_HexValue.Length - 2); i += 2)
            {
                sReverseHex_HexValue = sReverseHex_HexValue.Substring(0, i) + sReverseHex_HexValue.Substring(sReverseHex_HexValue.Length - 2, 2) + sReverseHex_HexValue.Substring(i, sReverseHex_HexValue.Length - (i + 2));
            }
            DebugMsg("ReverseHex(" + sReverseHex_OriginalHexValue + "): " + sReverseHex_HexValue + " - " + (System.DateTime.Now - StartTime).ToString());
            return sReverseHex_HexValue;
        }

        public static uint AptChangeContainer(uint ContainerType)
        {
            switch (ContainerType)
            {
                case 1: return 6;
                case 2: return 9;
                case 3: return 10;
                case 4: return 11;
                default: return 0;
            }
        }

        public static string GetRegistryString(string KeyName, string ValueName)
        {
            return GetRegistryString(KeyName, ValueName, string.Empty);
        }
        public static string GetRegistryString(string KeyName, string ValueName, string DefaultValue)
        {
            string ReturnValue = DefaultValue;

            try
            {
                ReturnValue = Registry.GetValue(KeyName.Replace("//", "/").Replace("/", "//"), ValueName, DefaultValue).ToString();
            }
            catch (Exception ex)
            {
                ReturnValue = DefaultValue;
                GetError("GetRegistryString(" + KeyName + ", " + ValueName + ", " + DefaultValue + "): " + ex.Message);
            }

            return ReturnValue;
        }

        /// <summary>
        /// Remove regional special chars from string
        /// </summary>
        /// <param name="SourceString"></param>
        /// <returns>Standardized string</returns>
        public static string RemoveSpecialChar(string SourceString)
        {
            StringBuilder ReturnValue = new StringBuilder(SourceString);

            ReturnValue.Replace('á', 'a');
            ReturnValue.Replace('à', 'a');
            ReturnValue.Replace('â', 'a');
            ReturnValue.Replace('ä', 'a');
            ReturnValue.Replace('ã', 'a');
            ReturnValue.Replace('Á', 'A');
            ReturnValue.Replace('À', 'A');
            ReturnValue.Replace('Â', 'A');
            ReturnValue.Replace('Ä', 'A');
            ReturnValue.Replace('Ã', 'A');
            ReturnValue.Replace('é', 'e');
            ReturnValue.Replace('è', 'e');
            ReturnValue.Replace('ê', 'e');
            ReturnValue.Replace('ë', 'e');
            ReturnValue.Replace('É', 'E');
            ReturnValue.Replace('È', 'E');
            ReturnValue.Replace('Ê', 'E');
            ReturnValue.Replace('Ë', 'E');
            ReturnValue.Replace('í', 'i');
            ReturnValue.Replace('ì', 'i');
            ReturnValue.Replace('î', 'i');
            ReturnValue.Replace('ï', 'i');
            ReturnValue.Replace('Í', 'I');
            ReturnValue.Replace('Ì', 'I');
            ReturnValue.Replace('Î', 'I');
            ReturnValue.Replace('Ï', 'I');
            ReturnValue.Replace('ó', 'o');
            ReturnValue.Replace('ò', 'o');
            ReturnValue.Replace('ô', 'o');
            ReturnValue.Replace('ö', 'o');
            ReturnValue.Replace('õ', 'o');
            ReturnValue.Replace('Ó', 'O');
            ReturnValue.Replace('Ò', 'O');
            ReturnValue.Replace('Ô', 'O');
            ReturnValue.Replace('Ö', 'O');
            ReturnValue.Replace('Õ', 'O');
            ReturnValue.Replace('ú', 'u');
            ReturnValue.Replace('ù', 'u');
            ReturnValue.Replace('û', 'u');
            ReturnValue.Replace('ü', 'u');
            ReturnValue.Replace('Ú', 'U');
            ReturnValue.Replace('Ù', 'U');
            ReturnValue.Replace('Û', 'U');
            ReturnValue.Replace('Ü', 'U');
            ReturnValue.Replace('ç', 'c');
            ReturnValue.Replace('Ç', 'C');
            ReturnValue.Replace('ñ', 'n');
            ReturnValue.Replace('Ñ', 'N');
            return ReturnValue.ToString();
        }

        /// <summary>
        /// Return the current date and time on "yyyyMMddHHmmss" format
        /// </summary>
        /// <returns>String with current date and time</returns>
        public static string DateTime() { return System.DateTime.Now.ToString("yyyyMMddHHmmss"); }
        /// <summary>
        /// Return the given date and time on "yyyyMMddHHmmss" format
        /// </summary>
        /// <returns>String with given date and time</returns>
        public static string DateTime(DateTime DateTimeValue) { return DateTimeValue.ToString("yyyyMMddHHmmss"); }

        /// <summary>
        /// Encode/decode @ptitude Analyst user password
        /// </summary>
        /// <param name="OriginalPass"></param>
        /// <returns>String with encoded/decoded password</returns>
        public static string PassEncode(string OriginalPass)
        {
            if (string.IsNullOrEmpty(OriginalPass)) return string.Empty;

            string EncodedPass = string.Empty;

            for (int i = 0; i < OriginalPass.Length; i++)
                EncodedPass = EncodedPass + (char)((int)OriginalPass[i] ^ (((int)((Math.Pow(2, (i+1)) - 1) % 127)==0) ? 127 : (int)((Math.Pow(2, (i+1)) - 1) % 127)));

            return EncodedPass;
        }
    }

    public static class HTMLExport_Formated
    {
        public static void Create_Html(string Title, DataTable Source, string DestFile)
        {
            if (File.Exists(DestFile)) File.Delete(DestFile);

            System.IO.StreamWriter file = new System.IO.StreamWriter(DestFile);
            StringBuilder html = Create_Html(Title, Source);
            file.WriteLine(html);
            file.Close();

        }
        public static StringBuilder Create_Html(string Title, DataTable Source)
        {
           

            StringBuilder html = new StringBuilder();
            html.Append(" <!doctype html> ");
            html.Append("<html>");
            html.Append("<head>");
            html.Append("	<title>" + GenericTools.RemoveSpecialChar(Title) + " - " + DateTime.Now + "</title>");
            html.Append("</head>");
            html.Append("<body>");
            html.Append("<p>&nbsp;</p>");
            html.Append("<h2 style='font-size:22px; color:gray; font-family:Verdana; border-bottom: solid 1px gray; text-align: center'> " + GenericTools.RemoveSpecialChar(Title) + " - " + DateTime.Now + "</h2>");
            html.Append("<p><table align='left' border='1' cellpadding='4' cellspacing='0' style='border: solid 1px #CCC; border-collapse: collapse;'>");
            html.Append("<tbody>");
            html.Append("	<tr>");

            foreach (DataColumn column in Source.Columns)
            {
                html.Append("		<td style='font-family:Verdana; font-size:12px; text-align: center; background-color: #0066CC; color:white;'>" + GenericTools.RemoveSpecialChar(column.ColumnName) + "</td>");
            }
            html.Append("	</tr>");

            foreach (DataRow dr2 in Source.Rows)
            {

                //StringBuilder HtmlItens = new StringBuilder();
                html.Append("	<tr>");

                for (int k = 0; k <= Source.Columns.Count - 1; k++)
                {
                    string Linha = GenericTools.RemoveSpecialChar(dr2[k].ToString());
                    string color = "white";

                    html.Append("		<td style='font-family:Verdana; font-size:11px; text-align: center; background-color: " + color + "'>" + Linha + "</td>");

                }
                html.Append("	</tr>");
            }

            html.Append("</tbody>");
            html.Append("</table>");
            html.Append("</p>");

       
            return html;

        }

        public static StringBuilder Create_Html(string Title, DataTable Source1, string Item2, DataTable Source2)
        {


            StringBuilder html = new StringBuilder();
            html.Append(" <!doctype html> ");
            html.Append("<html>");
            html.Append("<head>");
            html.Append("	<title>" + GenericTools.RemoveSpecialChar(Title) + " - " + DateTime.Now + "</title>");
            html.Append("</head>");
            html.Append("<body>");
            html.Append("<p>&nbsp;</p>");
            html.Append("<h2 style='font-size:22px; color:gray; font-family:Verdana; border-bottom: solid 1px gray; text-align: center'> " + GenericTools.RemoveSpecialChar(Title) + " - " + DateTime.Now + "</h2>");
            html.Append("<p><table align='left' border='1' cellpadding='4' cellspacing='0' style='border: solid 1px #CCC; border-collapse: collapse;'>");
            html.Append("<tbody>");
            html.Append("	<tr>");

            foreach (DataColumn column in Source1.Columns)
            {
                html.Append("		<td style='font-family:Verdana; font-size:12px; text-align: center; background-color: #0066CC; color:white;'>" + GenericTools.RemoveSpecialChar(column.ColumnName) + "</td>");
            }
            html.Append("	</tr>");

            foreach (DataRow dr2 in Source1.Rows)
            {

                //StringBuilder HtmlItens = new StringBuilder();
                html.Append("	<tr>");

                for (int k = 0; k <= Source1.Columns.Count - 1; k++)
                {
                    string Linha = GenericTools.RemoveSpecialChar(dr2[k].ToString());
                    string color = "white";

                    html.Append("		<td style='font-family:Verdana; font-size:11px; text-align: center; background-color: " + color + "'>" + Linha + "</td>");

                }
                html.Append("	</tr>");
            }

            html.Append("</tbody>");
            html.Append("</table>");
            html.Append("</p>");


            html.Append("<p>&nbsp;</p>");
            html.Append("<h2 style='font-size:22px; color:gray; font-family:Verdana; border-bottom: solid 1px gray; text-align: center'> " + GenericTools.RemoveSpecialChar(Item2) + "</h2>");
            html.Append("<p><table align='left' border='1' cellpadding='4' cellspacing='0' style='border: solid 1px #CCC; border-collapse: collapse;'>");
            html.Append("<tbody>");
            html.Append("	<tr>");

            foreach (DataColumn column in Source2.Columns)
            {
                html.Append("		<td style='font-family:Verdana; font-size:12px; text-align: center; background-color: #0066CC; color:white;'>" + GenericTools.RemoveSpecialChar(column.ColumnName) + "</td>");
            }
            html.Append("	</tr>");

            foreach (DataRow dr2 in Source2.Rows)
            {

                html.Append("	<tr>");

                for (int k = 0; k <= Source2.Columns.Count - 1; k++)
                {
                    string Linha = GenericTools.RemoveSpecialChar(dr2[k].ToString());
                    string color = "white";

                    html.Append("		<td style='font-family:Verdana; font-size:11px; text-align: center; background-color: " + color + "'>" + Linha + "</td>");

                }
                html.Append("	</tr>");
            }

            html.Append("</tbody>");
            html.Append("</table>");
            html.Append("</p>");


            return html;

        }

    }

    public static class CSVExport
    {
        public static void ExportToCSV(DataGridView GrdView, string DestFile, bool RemoveSpecialChars)
        {
            ExportToCSV(GrdView, DestFile, System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator, "\"", RemoveSpecialChars);
        }
        public static void ExportToCSV(DataGridView GrdView, string DestFile)
        {
            ExportToCSV(GrdView, DestFile, System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator, "\"", false);
        }
        public static void ExportToCSV(DataGridView GrdView, string DestFile, string CellDivisor, bool RemoveSpecialChars)
        {
            ExportToCSV(GrdView, DestFile, CellDivisor, "\"", RemoveSpecialChars);
        }
        public static void ExportToCSV(DataGridView GrdView, string DestFile, string CellDivisor)
        {
            ExportToCSV(GrdView, DestFile, CellDivisor, "\"", false);
        }
        public static void ExportToCSV(DataGridView GrdView, string DestFile, string CellDivisor, string TextDelimiter, bool RemoveSpecialChars)
        {
            try
            {
                // Open the file
                GenericTools.DebugMsg("Openning the file " + DestFile);
                StreamWriter fs = new StreamWriter(DestFile, false, Encoding.UTF8);

                // Creating header
                try
                {
                    GenericTools.DebugMsg("Creating header...");
                    for (Int32 i = 0; i < GrdView.ColumnCount; i++)
                    {
                        if (RemoveSpecialChars)
                        {
                            fs.Write(GenericTools.RemoveSpecialChar(GrdView.Columns[i].HeaderText));
                        }
                        else
                        {
                            fs.Write(GrdView.Columns[i].HeaderText);
                        }

                        if (i < (GrdView.ColumnCount - 1)) fs.Write(CellDivisor);
                    }
                    fs.WriteLine(string.Empty);
                }
                catch (Exception ex)
                {
                    GenericTools.GetError("CSVExport.ExportToCSV(" + GrdView.Name.ToString() + ", " + DestFile + ", " + CellDivisor + ", " + TextDelimiter + "): " + ex.Message);
                }

                GenericTools.DebugMsg("Saving data to file...");
                // Saving data
                try
                {
                    for (Int32 i = 0; i < GrdView.RowCount; i++)
                    {
                        for (Int32 j = 0; j < GrdView.ColumnCount; j++)
                        {
                            try
                            {
                                if ((GrdView.Rows[i].Cells[j].ValueType == Type.GetType("System.String")) | (GrdView.Rows[i].Cells[j].ValueType == Type.GetType("System.Object"))) fs.Write(TextDelimiter);
                            }
                            catch (Exception ex)
                            {
                                GenericTools.GetError("CSVExport.ExportToCSV(" + GrdView.Name.ToString() + ", " + DestFile + ", " + CellDivisor + ", " + TextDelimiter + "): " + ex.Message);
                            }
                            if (RemoveSpecialChars)
                            {
                                try
                                {
                                    fs.Write(GenericTools.RemoveSpecialChar(GrdView.Rows[i].Cells[j].Value.ToString()));
                                }
                                catch (Exception ex)
                                {
                                    GenericTools.GetError("CSVExport.ExportToCSV(" + GrdView.Name.ToString() + ", " + DestFile + ", " + CellDivisor + ", " + TextDelimiter + "): " + ex.Message);
                                }
                            }
                            else
                            {
                                try
                                {
                                    fs.Write(GrdView.Rows[i].Cells[j].Value.ToString());
                                }
                                catch (Exception ex)
                                {
                                    GenericTools.GetError("CSVExport.ExportToCSV(" + GrdView.Name.ToString() + ", " + DestFile + ", " + CellDivisor + ", " + TextDelimiter + "): " + ex.Message);
                                }
                            }
                            try
                            {
                                if ((GrdView.Rows[i].Cells[j].ValueType == Type.GetType("System.String")) | (GrdView.Rows[i].Cells[j].ValueType == Type.GetType("System.Object"))) fs.Write(TextDelimiter);
                            }
                            catch (Exception ex)
                            {
                                GenericTools.GetError("CSVExport.ExportToCSV(" + GrdView.Name.ToString() + ", " + DestFile + ", " + CellDivisor + ", " + TextDelimiter + "): " + ex.Message);
                            }
                            try
                            {
                                if (j < (GrdView.ColumnCount - 1)) fs.Write(CellDivisor);
                            }
                            catch (Exception ex)
                            {
                                GenericTools.GetError("CSVExport.ExportToCSV(" + GrdView.Name.ToString() + ", " + DestFile + ", " + CellDivisor + ", " + TextDelimiter + "): " + ex.Message);
                            }
                        }
                        try
                        {
                            fs.WriteLine(string.Empty);
                        }
                        catch (Exception ex)
                        {
                            GenericTools.GetError("CSVExport.ExportToCSV(" + GrdView.Name.ToString() + ", " + DestFile + ", " + CellDivisor + ", " + TextDelimiter + "): " + ex.Message);
                        }
                    }
                }
                catch (Exception ex)
                {
                    GenericTools.GetError("CSVExport.ExportToCSV(" + GrdView.Name.ToString() + ", " + DestFile + ", " + CellDivisor + ", " + TextDelimiter + "): " + ex.Message);
                }


                // Closing file
                fs.Close();
            }
            catch (Exception ex)
            {
                GenericTools.GetError("CSVExport.ExportToCSV(" + GrdView.Name.ToString() + ", " + DestFile + ", " + CellDivisor + ", " + TextDelimiter + "): " + ex.Message);
            }
        }

        public static void ExportToCSV(DataTable Table, string DestFile, bool RemoveSpecialChars)
        {
            ExportToCSV(Table, DestFile, System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator, "\"", RemoveSpecialChars);
        }
        public static void ExportToCSV(DataTable Table, string DestFile)
        {
            ExportToCSV(Table, DestFile, System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator, "\"", false);
        }
        public static void ExportToCSV(DataTable Table, string DestFile, string CellDivisor, bool RemoveSpecialChars)
        {
            ExportToCSV(Table, DestFile, CellDivisor, "\"", RemoveSpecialChars);
        }
        public static void ExportToCSV(DataTable Table, string DestFile, string CellDivisor)
        {
            ExportToCSV(Table, DestFile, CellDivisor, "\"", false);
        }
        public static void ExportToCSV(DataTable Table, string DestFile, bool Headers, bool RemoveSpecialChars)
        {
            ExportToCSV(Table, DestFile, System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator, "\"", RemoveSpecialChars, Headers);
        }
        public static void ExportToCSV(DataTable Table, string DestFile, string CellDivisor, string TextDelimiter, bool RemoveSpecialChars, bool Headers = true)
        {
            try
            {
                // Open the file
                GenericTools.DebugMsg("Openning the file " + DestFile);
                StreamWriter fs = new StreamWriter(DestFile, false);

                // Creating header
                if (Headers)
                {
                    GenericTools.DebugMsg("Creating header...");
                    for (Int32 i = 0; i < Table.Columns.Count; i++)
                    {
                        fs.Write(Table.Columns[i].ColumnName);
                        if (i < (Table.Columns.Count - 1)) fs.Write(CellDivisor);
                    }
                    fs.WriteLine(string.Empty);
                }

                GenericTools.DebugMsg("Saving data to file...");
                // Saving data
                for (Int32 i = 0; i < Table.Rows.Count; i++)
                {
                    for (Int32 j = 0; j < Table.Columns.Count; j++)
                    {
                        if (Table.Rows[i][j].GetType() == Type.GetType("System.String")) fs.Write(TextDelimiter);
                        if (RemoveSpecialChars)
                        {
                            fs.Write(GenericTools.RemoveSpecialChar(Table.Rows[i][j].ToString()));
                        }
                        else
                        {
                            fs.Write(Table.Rows[i][j].ToString());
                        }
                        if (Table.Rows[i][j].GetType() == Type.GetType("System.String")) fs.Write(TextDelimiter);
                        if (j < (Table.Columns.Count - 1)) fs.Write(CellDivisor);
                    }
                    fs.WriteLine(string.Empty);
                }

                // Closing file
                fs.Close();
            }
            catch (Exception ex)
            {
                GenericTools.GetError("CSVExport.ExportToCSV(" + Table.TableName + ", " + DestFile + ", " + CellDivisor + ", " + TextDelimiter + "): " + ex.Message);
            }
        }
    }

    public static class HTMLExport
    {
        static private string ToHTML(string text)
        {
            string ReturnValue = text.Trim();

            ReturnValue.Replace("á", "&aacute;");
            ReturnValue.Replace("é", "&eacute;");
            ReturnValue.Replace("í", "&iacute;");
            ReturnValue.Replace("ó", "&oacute;");
            ReturnValue.Replace("ú", "&uacute;");

            ReturnValue.Replace("Á", "&Aacute;");
            ReturnValue.Replace("É", "&Eacute;");
            ReturnValue.Replace("Í", "&Iacute;");
            ReturnValue.Replace("Ó", "&Oacute;");
            ReturnValue.Replace("Ú", "&Uacute;");

            ReturnValue.Replace("ç", "&ccedil;");
            ReturnValue.Replace("Ç", "&Ccedil;");

            ReturnValue.Replace("ã", "&atilde;");
            ReturnValue.Replace("õ", "&otilde;");

            ReturnValue.Replace("Ã", "&Atilde;");
            ReturnValue.Replace("Õ", "&Otilde;");

            ReturnValue.Replace("à", "&agrave;");
            ReturnValue.Replace("À", "&Agrave;");

            ReturnValue = GenericTools.RemoveSpecialChar(ReturnValue);

            if (string.IsNullOrEmpty(ReturnValue.Trim())) ReturnValue = "&nbsp;";
            return ReturnValue;
        }

        public static void DataTableToHTML(DataTable GrdView, string DestFile, string Title)
        {
            try
            {
                // Open the file
                GenericTools.DebugMsg("Openning the file " + DestFile);
                StreamWriter fs = new StreamWriter(DestFile, false);

                // Creating header
                try
                {
                    GenericTools.DebugMsg("Creating header...");
                    fs.WriteLine("<html>");
                    fs.WriteLine("    <head>");
                    fs.WriteLine("        <meta name=ProgId content=Excel.Sheet>");
                    fs.WriteLine("        <title>" + ToHTML(Title) + "</title>");
                    fs.WriteLine("    </head>");
                    fs.WriteLine(string.Empty);
                    fs.WriteLine("    <body>");
                    fs.WriteLine("        <table width=99% border=1 align=center>");
                    fs.WriteLine("            <tr>");

                    for (Int32 i = 0; i < GrdView.Columns.Count; i++)
                    {
                        fs.WriteLine("                <td>");
                        fs.Write(GrdView.Columns[i].ColumnName);
                        fs.WriteLine("                </td>");
                    }
                    fs.WriteLine("            </tr>");
                    fs.WriteLine(string.Empty);
                }
                catch (Exception ex)
                {
                    GenericTools.GetError("HTMLExport.ExportToHTML(" + GrdView.TableName.ToString() + ", " + DestFile + ", " + Title + "): " + ex.Message);
                }

                GenericTools.DebugMsg("Saving data to file...");
                // Saving data
                try
                {

                    for (Int32 i = 0; i < GrdView.Rows.Count; i++)
                    {
                        fs.WriteLine("            <tr>");
                        for (Int32 j = 0; j < GrdView.Columns.Count; j++)
                        {
                            fs.WriteLine("                <td>");
                            try
                            {
                                fs.Write(ToHTML(GrdView.Rows[i][j].ToString()));
                            }
                            catch (Exception ex)
                            {
                                GenericTools.GetError("HTMLExport.ExportToHTML(" + GrdView.TableName.ToString() + ", " + DestFile + ", " + Title + "): " + ex.Message);
                            }
                            fs.WriteLine("                </td>");
                        }
                        try
                        {
                            fs.WriteLine(string.Empty);
                        }
                        catch (Exception ex)
                        {
                            GenericTools.GetError("HTMLExport.ExportToHTML(" + GrdView.TableName.ToString() + ", " + DestFile + ", " + Title + "): " + ex.Message);
                        }
                        fs.WriteLine("            </tr>");
                    }
                }
                catch (Exception ex)
                {
                    GenericTools.GetError("HTMLExport.ExportToHTML(" + GrdView.TableName.ToString() + ", " + DestFile + ", " + Title + "): " + ex.Message);
                }


                // Closing file
                fs.WriteLine("        </table>");
                fs.WriteLine("    </body>");
                fs.WriteLine("</html>");
                fs.Close();
            }
            catch (Exception ex)
            {
                GenericTools.GetError("HTMLExport.ExportToHTML(" + GrdView.TableName.ToString() + ", " + DestFile + ", " + Title + "): " + ex.Message);
            }
        }
        public static void ExportToHTML(DataGridView GrdView, string DestFile, string Title)
        {
            try
            {
                // Open the file
                GenericTools.DebugMsg("Openning the file " + DestFile);
                StreamWriter fs = new StreamWriter(DestFile, false);

                // Creating header
                try
                {
                    GenericTools.DebugMsg("Creating header...");
                    fs.WriteLine("<html>");
                    fs.WriteLine("    <head>");
                    fs.WriteLine("        <meta name=ProgId content=Excel.Sheet>");
                    fs.WriteLine("        <title>" + ToHTML(Title) + "</title>");
                    fs.WriteLine("    </head>");
                    fs.WriteLine(string.Empty);
                    fs.WriteLine("    <body>");
                    fs.WriteLine("        <table width=99% border=1 align=center>");
                    fs.WriteLine("            <tr>");

                    for (Int32 i = 0; i < GrdView.ColumnCount; i++)
                    {
                        fs.WriteLine("                <td>");
                            fs.Write(GrdView.Columns[i].HeaderText);
                        fs.WriteLine("                </td>");
                    }
                    fs.WriteLine("            </tr>");
                    fs.WriteLine(string.Empty);
                }
                catch (Exception ex)
                {
                    GenericTools.GetError("HTMLExport.ExportToHTML(" + GrdView.Name.ToString() + ", " + DestFile + ", " + Title + "): " + ex.Message);
                }

                GenericTools.DebugMsg("Saving data to file...");
                // Saving data
                try
                {
                    for (Int32 i = 0; i < GrdView.RowCount; i++)
                    {
                    fs.WriteLine("            <tr>");
                        for (Int32 j = 0; j < GrdView.ColumnCount; j++)
                        {
                            fs.WriteLine("                <td" + (GrdView.Rows[i].Cells[j].Style.BackColor.IsEmpty ? string.Empty : " bgcolor='" + GrdView.Rows[i].Cells[j].Style.BackColor.Name + "'") + ">");
                        try
                        {
                            fs.Write(ToHTML(GrdView.Rows[i].Cells[j].Value.ToString()));
                        }
                        catch (Exception ex)
                        {
                            GenericTools.GetError("HTMLExport.ExportToHTML(" + GrdView.Name.ToString() + ", " + DestFile + ", " + Title + "): " + ex.Message);
                        }
                        fs.WriteLine("                </td>");
                        }
                        try
                        {
                            fs.WriteLine(string.Empty);
                        }
                        catch (Exception ex)
                        {
                            GenericTools.GetError("HTMLExport.ExportToHTML(" + GrdView.Name.ToString() + ", " + DestFile + ", " + Title + "): " + ex.Message);
                        }
                                            fs.WriteLine("            </tr>");
                    }
                }
                catch (Exception ex)
                {
                    GenericTools.GetError("HTMLExport.ExportToHTML(" + GrdView.Name.ToString() + ", " + DestFile + ", " + Title + "): " + ex.Message);
                }


                // Closing file
            fs.WriteLine("        </table>");
            fs.WriteLine("    </body>");
            fs.WriteLine("</html>");
            fs.Close();
            }
            catch (Exception ex)
            {
                GenericTools.GetError("HTMLExport.ExportToHTML(" + GrdView.Name.ToString() + ", " + DestFile + ", " + Title + "): " + ex.Message);
            }
        }
        public static string DataTableToHTMLTable(DataTable GrdView, uint WidthPerc = 100)
        {
            StringBuilder HTMLTable = new StringBuilder();
            try
            {
                HTMLTable.Append("<table width=" + WidthPerc.ToString() + "% border=1 align=center>");
                HTMLTable.Append("<tbody>");

                //Creating header
                HTMLTable.Append("<tr>");
                for (int i = 0; i < GrdView.Columns.Count; i++)
                    HTMLTable.Append("<td>" + GrdView.Columns[i].ColumnName + "</td>");
                HTMLTable.Append("</tr>");

                //Filling table
                for (Int32 i = 0; i < GrdView.Rows.Count; i++)
                {
                    HTMLTable.Append("<tr>");
                    for (Int32 j = 0; j < GrdView.Columns.Count; j++)
                    {
                        HTMLTable.Append("<td>");
                        HTMLTable.Append(ToHTML(GrdView.Rows[i][j].ToString()));
                        HTMLTable.Append("</td>");
                    }
                    HTMLTable.Append("</tr>");
                }

                // Ending table
                HTMLTable.Append("</tbody>");
                HTMLTable.Append("</table>");
            }
            catch (Exception ex)
            {
                GenericTools.GetError("HTMLExport.ExportToHTML(" + GrdView.TableName.ToString() + ", " + WidthPerc.ToString() + "): " + ex.Message);
                HTMLTable.Clear();
            }
            return HTMLTable.ToString();
        }
    }

    public class IniReader
    {
        // API declarations
        /// <summary>
        /// The GetPrivateProfileInt function retrieves an integer associated with a key in the specified section of an initialization file.
        /// </summary>
        /// <param name="lpApplicationName">Pointer to a null-terminated string specifying the name of the section in the initialization file.</param>
        /// <param name="lpKeyName">Pointer to the null-terminated string specifying the name of the key whose value is to be retrieved. This value is in the form of a string; the GetPrivateProfileInt function converts the string into an integer and returns the integer.</param>
        /// <param name="nDefault">Specifies the default value to return if the key name cannot be found in the initialization file.</param>
        /// <param name="lpFileName">Pointer to a null-terminated string that specifies the name of the initialization file. If this parameter does not contain a full path to the file, the system searches for the file in the Windows directory.</param>
        /// <returns>The return value is the integer equivalent of the string following the specified key name in the specified initialization file. If the key is not found, the return value is the specified default value. If the value of the key is less than zero, the return value is zero.</returns>
        [DllImport("KERNEL32.DLL", EntryPoint = "GetPrivateProfileIntA", CharSet = CharSet.Ansi)]
        private static extern int GetPrivateProfileInt(string lpApplicationName, string lpKeyName, int nDefault, string lpFileName);
        /// <summary>
        /// The WritePrivateProfileString function copies a string into the specified section of an initialization file.
        /// </summary>
        /// <param name="lpApplicationName">Pointer to a null-terminated string containing the name of the section to which the string will be copied. If the section does not exist, it is created. The name of the section is case-independent; the string can be any combination of uppercase and lowercase letters.</param>
        /// <param name="lpKeyName">Pointer to the null-terminated string containing the name of the key to be associated with a string. If the key does not exist in the specified section, it is created. If this parameter is NULL, the entire section, including all entries within the section, is deleted.</param>
        /// <param name="lpString">Pointer to a null-terminated string to be written to the file. If this parameter is NULL, the key pointed to by the lpKeyName parameter is deleted.</param>
        /// <param name="lpFileName">Pointer to a null-terminated string that specifies the name of the initialization file.</param>
        /// <returns>If the function successfully copies the string to the initialization file, the return value is nonzero; if the function fails, or if it flushes the cached version of the most recently accessed initialization file, the return value is zero.</returns>
        [DllImport("KERNEL32.DLL", EntryPoint = "WritePrivateProfileStringA", CharSet = CharSet.Ansi)]
        private static extern int WritePrivateProfileString(string lpApplicationName, string lpKeyName, string lpString, string lpFileName);
        /// <summary>
        /// The GetPrivateProfileString function retrieves a string from the specified section in an initialization file.
        /// </summary>
        /// <param name="lpApplicationName">Pointer to a null-terminated string that specifies the name of the section containing the key name. If this parameter is NULL, the GetPrivateProfileString function copies all section names in the file to the supplied buffer.</param>
        /// <param name="lpKeyName">Pointer to the null-terminated string specifying the name of the key whose associated string is to be retrieved. If this parameter is NULL, all key names in the section specified by the lpAppName parameter are copied to the buffer specified by the lpReturnedString parameter.</param>
        /// <param name="lpDefault">Pointer to a null-terminated default string. If the lpKeyName key cannot be found in the initialization file, GetPrivateProfileString copies the default string to the lpReturnedString buffer. This parameter cannot be NULL. <br>Avoid specifying a default string with trailing blank characters. The function inserts a null character in the lpReturnedString buffer to strip any trailing blanks.</br></param>
        /// <param name="lpReturnedString">Pointer to the buffer that receives the retrieved string.</param>
        /// <param name="nSize">Specifies the size, in TCHARs, of the buffer pointed to by the lpReturnedString parameter.</param>
        /// <param name="lpFileName">Pointer to a null-terminated string that specifies the name of the initialization file. If this parameter does not contain a full path to the file, the system searches for the file in the Windows directory.</param>
        /// <returns>The return value is the number of characters copied to the buffer, not including the terminating null character.</returns>
        [DllImport("KERNEL32.DLL", EntryPoint = "GetPrivateProfileStringA", CharSet = CharSet.Ansi)]
        private static extern int GetPrivateProfileString(string lpApplicationName, string lpKeyName, string lpDefault, StringBuilder lpReturnedString, int nSize, string lpFileName);
        /// <summary>
        /// The GetPrivateProfileSectionNames function retrieves the names of all sections in an initialization file.
        /// </summary>
        /// <param name="lpszReturnBuffer">Pointer to a buffer that receives the section names associated with the named file. The buffer is filled with one or more null-terminated strings; the last string is followed by a second null character.</param>
        /// <param name="nSize">Specifies the size, in TCHARs, of the buffer pointed to by the lpszReturnBuffer parameter.</param>
        /// <param name="lpFileName">Pointer to a null-terminated string that specifies the name of the initialization file. If this parameter is NULL, the function searches the Win.ini file. If this parameter does not contain a full path to the file, the system searches for the file in the Windows directory.</param>
        /// <returns>The return value specifies the number of characters copied to the specified buffer, not including the terminating null character. If the buffer is not large enough to contain all the section names associated with the specified initialization file, the return value is equal to the length specified by nSize minus two.</returns>
        [DllImport("KERNEL32.DLL", EntryPoint = "GetPrivateProfileSectionNamesA", CharSet = CharSet.Ansi)]
        private static extern int GetPrivateProfileSectionNames(byte[] lpszReturnBuffer, int nSize, string lpFileName);
        /// <summary>
        /// The WritePrivateProfileSection function replaces the keys and values for the specified section in an initialization file.
        /// </summary>
        /// <param name="lpAppName">Pointer to a null-terminated string specifying the name of the section in which data is written. This section name is typically the name of the calling application.</param>
        /// <param name="lpString">Pointer to a buffer containing the new key names and associated values that are to be written to the named section.</param>
        /// <param name="lpFileName">Pointer to a null-terminated string containing the name of the initialization file. If this parameter does not contain a full path for the file, the function searches the Windows directory for the file. If the file does not exist and lpFileName does not contain a full path, the function creates the file in the Windows directory. The function does not create a file if lpFileName contains the full path and file name of a file that does not exist.</param>
        /// <returns>If the function succeeds, the return value is nonzero.<br>If the function fails, the return value is zero.</br></returns>
        [DllImport("KERNEL32.DLL", EntryPoint = "WritePrivateProfileSectionA", CharSet = CharSet.Ansi)]
        private static extern int WritePrivateProfileSection(string lpAppName, string lpString, string lpFileName);
        /// <summary>Constructs a new IniReader instance.</summary>
        /// <param name="file">Specifies the full path to the INI file (the file doesn't have to exist).</param>
        public IniReader()
        {
            Filename = GenericTools.GetAuxFileName("ini");
        }
        public IniReader(string file)
        {
            Filename = file;
        }
        /// <summary>Gets or sets the full path to the INI file.</summary>
        /// <value>A String representing the full path to the INI file.</value>
        public string Filename
        {
            get
            {
                return m_Filename;
            }
            set
            {
                m_Filename = value;
            }
        }
        /// <summary>Gets or sets the section you're working in. (aka 'the active section')</summary>
        /// <value>A String representing the section you're working in.</value>
        public string Section
        {
            get
            {
                return m_Section;
            }
            set
            {
                m_Section = value;
            }
        }
        /// <summary>Reads an Integer from the specified key of the specified section.</summary>
        /// <param name="section">The section to search in.</param>
        /// <param name="key">The key from which to return the value.</param>
        /// <param name="defVal">The value to return if the specified key isn't found.</param>
        /// <returns>Returns the value of the specified section/key pair, or returns the default value if the specified section/key pair isn't found in the INI file.</returns>
        public int ReadInteger(string section, string key, int defVal)
        {
            return GetPrivateProfileInt(section, key, defVal, Filename);
        }
        /// <summary>Reads an Integer from the specified key of the specified section.</summary>
        /// <param name="section">The section to search in.</param>
        /// <param name="key">The key from which to return the value.</param>
        /// <returns>Returns the value of the specified section/key pair, or returns 0 if the specified section/key pair isn't found in the INI file.</returns>
        public int ReadInteger(string section, string key)
        {
            return ReadInteger(section, key, 0);
        }
        /// <summary>Reads an Integer from the specified key of the active section.</summary>
        /// <param name="key">The key from which to return the value.</param>
        /// <param name="defVal">The section to search in.</param>
        /// <returns>Returns the value of the specified Key, or returns the default value if the specified Key isn't found in the active section of the INI file.</returns>
        public int ReadInteger(string key, int defVal)
        {
            return ReadInteger(Section, key, defVal);
        }
        /// <summary>Reads an Integer from the specified key of the active section.</summary>
        /// <param name="key">The key from which to return the value.</param>
        /// <returns>Returns the value of the specified key, or returns 0 if the specified key isn't found in the active section of the INI file.</returns>
        public int ReadInteger(string key)
        {
            return ReadInteger(key, 0);
        }
        /// <summary>Reads a String from the specified key of the specified section.</summary>
        /// <param name="section">The section to search in.</param>
        /// <param name="key">The key from which to return the value.</param>
        /// <param name="defVal">The value to return if the specified key isn't found.</param>
        /// <returns>Returns the value of the specified section/key pair, or returns the default value if the specified section/key pair isn't found in the INI file.</returns>
        public string ReadString(string section, string key, string defVal)
        {
            StringBuilder sb = new StringBuilder(MAX_ENTRY);
            int Ret = GetPrivateProfileString(section, key, defVal, sb, MAX_ENTRY, Filename);
            return sb.ToString();
        }
        /// <summary>Reads a String from the specified key of the specified section.</summary>
        /// <param name="section">The section to search in.</param>
        /// <param name="key">The key from which to return the value.</param>
        /// <returns>Returns the value of the specified section/key pair, or returns an empty String if the specified section/key pair isn't found in the INI file.</returns>
        public string ReadString(string section, string key)
        {
            return ReadString(section, key, "");
        }
        /// <summary>Reads a String from the specified key of the active section.</summary>
        /// <param name="key">The key from which to return the value.</param>
        /// <returns>Returns the value of the specified key, or returns an empty String if the specified key isn't found in the active section of the INI file.</returns>
        public string ReadString(string key)
        {
            return ReadString(Section, key);
        }
        /// <summary>Reads a Long from the specified key of the specified section.</summary>
        /// <param name="section">The section to search in.</param>
        /// <param name="key">The key from which to return the value.</param>
        /// <param name="defVal">The value to return if the specified key isn't found.</param>
        /// <returns>Returns the value of the specified section/key pair, or returns the default value if the specified section/key pair isn't found in the INI file.</returns>
        public long ReadLong(string section, string key, long defVal)
        {
            return long.Parse(ReadString(section, key, defVal.ToString()));
        }
        /// <summary>Reads a Long from the specified key of the specified section.</summary>
        /// <param name="section">The section to search in.</param>
        /// <param name="key">The key from which to return the value.</param>
        /// <returns>Returns the value of the specified section/key pair, or returns 0 if the specified section/key pair isn't found in the INI file.</returns>
        public long ReadLong(string section, string key)
        {
            return ReadLong(section, key, 0);
        }
        /// <summary>Reads a Long from the specified key of the active section.</summary>
        /// <param name="key">The key from which to return the value.</param>
        /// <param name="defVal">The section to search in.</param>
        /// <returns>Returns the value of the specified key, or returns the default value if the specified key isn't found in the active section of the INI file.</returns>
        public long ReadLong(string key, long defVal)
        {
            return ReadLong(Section, key, defVal);
        }
        /// <summary>Reads a Long from the specified key of the active section.</summary>
        /// <param name="key">The key from which to return the value.</param>
        /// <returns>Returns the value of the specified Key, or returns 0 if the specified Key isn't found in the active section of the INI file.</returns>
        public long ReadLong(string key)
        {
            return ReadLong(key, 0);
        }
        /// <summary>Reads a Byte array from the specified key of the specified section.</summary>
        /// <param name="section">The section to search in.</param>
        /// <param name="key">The key from which to return the value.</param>
        /// <returns>Returns the value of the specified section/key pair, or returns null (Nothing in VB.NET) if the specified section/key pair isn't found in the INI file.</returns>
        public byte[] ReadByteArray(string section, string key)
        {
            try
            {
                return Convert.FromBase64String(ReadString(section, key));
            }
            catch { }
            return null;
        }
        /// <summary>Reads a Byte array from the specified key of the active section.</summary>
        /// <param name="key">The key from which to return the value.</param>
        /// <returns>Returns the value of the specified key, or returns null (Nothing in VB.NET) if the specified key pair isn't found in the active section of the INI file.</returns>
        public byte[] ReadByteArray(string key)
        {
            return ReadByteArray(Section, key);
        }
        /// <summary>Reads a Boolean from the specified key of the specified section.</summary>
        /// <param name="section">The section to search in.</param>
        /// <param name="key">The key from which to return the value.</param>
        /// <param name="defVal">The value to return if the specified key isn't found.</param>
        /// <returns>Returns the value of the specified section/key pair, or returns the default value if the specified section/key pair isn't found in the INI file.</returns>
        public bool ReadBoolean(string section, string key, bool defVal)
        {
            return Boolean.Parse(ReadString(section, key, defVal.ToString()));
        }
        /// <summary>Reads a Boolean from the specified key of the specified section.</summary>
        /// <param name="section">The section to search in.</param>
        /// <param name="key">The key from which to return the value.</param>
        /// <returns>Returns the value of the specified section/key pair, or returns false if the specified section/key pair isn't found in the INI file.</returns>
        public bool ReadBoolean(string section, string key)
        {
            return ReadBoolean(section, key, false);
        }
        /// <summary>Reads a Boolean from the specified key of the specified section.</summary>
        /// <param name="key">The key from which to return the value.</param>
        /// <param name="defVal">The value to return if the specified key isn't found.</param>
        /// <returns>Returns the value of the specified key pair, or returns the default value if the specified key isn't found in the active section of the INI file.</returns>
        public bool ReadBoolean(string key, bool defVal)
        {
            return ReadBoolean(Section, key, defVal);
        }
        /// <summary>Reads a Boolean from the specified key of the specified section.</summary>
        /// <param name="key">The key from which to return the value.</param>
        /// <returns>Returns the value of the specified key, or returns false if the specified key isn't found in the active section of the INI file.</returns>
        public bool ReadBoolean(string key)
        {
            return ReadBoolean(Section, key);
        }
        /// <summary>Writes an Integer to the specified key in the specified section.</summary>
        /// <param name="section">The section to write in.</param>
        /// <param name="key">The key to write to.</param>
        /// <param name="value">The value to write.</param>
        /// <returns>Returns true if the function succeeds, false otherwise.</returns>
        public bool Write(string section, string key, int value)
        {
            return Write(section, key, value.ToString());
        }
        /// <summary>Writes an Integer to the specified key in the active section.</summary>
        /// <param name="key">The key to write to.</param>
        /// <param name="value">The value to write.</param>
        /// <returns>Returns true if the function succeeds, false otherwise.</returns>
        public bool Write(string key, int value)
        {
            return Write(Section, key, value);
        }
        /// <summary>Writes a String to the specified key in the specified section.</summary>
        /// <param name="section">Specifies the section to write in.</param>
        /// <param name="key">Specifies the key to write to.</param>
        /// <param name="value">Specifies the value to write.</param>
        /// <returns>Returns true if the function succeeds, false otherwise.</returns>
        public bool Write(string section, string key, string value)
        {
            return (WritePrivateProfileString(section, key, value, Filename) != 0);
        }
        /// <summary>Writes a String to the specified key in the active section.</summary>
        ///	<param name="key">The key to write to.</param>
        /// <param name="value">The value to write.</param>
        /// <returns>Returns true if the function succeeds, false otherwise.</returns>
        public bool Write(string key, string value)
        {
            return Write(Section, key, value);
        }
        /// <summary>Writes a Long to the specified key in the specified section.</summary>
        /// <param name="section">The section to write in.</param>
        /// <param name="key">The key to write to.</param>
        /// <param name="value">The value to write.</param>
        /// <returns>Returns true if the function succeeds, false otherwise.</returns>
        public bool Write(string section, string key, long value)
        {
            return Write(section, key, value.ToString());
        }
        /// <summary>Writes a Long to the specified key in the active section.</summary>
        /// <param name="key">The key to write to.</param>
        /// <param name="value">The value to write.</param>
        /// <returns>Returns true if the function succeeds, false otherwise.</returns>
        public bool Write(string key, long value)
        {
            return Write(Section, key, value);
        }
        /// <summary>Writes a Byte array to the specified key in the specified section.</summary>
        /// <param name="section">The section to write in.</param>
        /// <param name="key">The key to write to.</param>
        /// <param name="value">The value to write.</param>
        /// <returns>Returns true if the function succeeds, false otherwise.</returns>
        public bool Write(string section, string key, byte[] value)
        {
            if (value == null)
                return Write(section, key, (string)null);
            else
                return Write(section, key, value, 0, value.Length);
        }
        /// <summary>Writes a Byte array to the specified key in the active section.</summary>
        /// <param name="key">The key to write to.</param>
        /// <param name="value">The value to write.</param>
        /// <returns>Returns true if the function succeeds, false otherwise.</returns>
        public bool Write(string key, byte[] value)
        {
            return Write(Section, key, value);
        }
        /// <summary>Writes a Byte array to the specified key in the specified section.</summary>
        /// <param name="section">The section to write in.</param>
        /// <param name="key">The key to write to.</param>
        /// <param name="value">The value to write.</param>
        /// <param name="offset">An offset in <i>value</i>.</param>
        /// <param name="length">The number of elements of <i>value</i> to convert.</param>
        /// <returns>Returns true if the function succeeds, false otherwise.</returns>
        public bool Write(string section, string key, byte[] value, int offset, int length)
        {
            if (value == null)
                return Write(section, key, (string)null);
            else
                return Write(section, key, Convert.ToBase64String(value, offset, length));
        }
        /// <summary>Writes a Boolean to the specified key in the specified section.</summary>
        /// <param name="section">The section to write in.</param>
        /// <param name="key">The key to write to.</param>
        /// <param name="value">The value to write.</param>
        /// <returns>Returns true if the function succeeds, false otherwise.</returns>
        public bool Write(string section, string key, bool value)
        {
            return Write(section, key, value.ToString());
        }
        /// <summary>Writes a Boolean to the specified key in the active section.</summary>
        /// <param name="key">The key to write to.</param>
        /// <param name="value">The value to write.</param>
        /// <returns>Returns true if the function succeeds, false otherwise.</returns>
        public bool Write(string key, bool value)
        {
            return Write(Section, key, value);
        }
        /// <summary>Deletes a key from the specified section.</summary>
        /// <param name="section">The section to delete from.</param>
        /// <param name="key">The key to delete.</param>
        /// <returns>Returns true if the function succeeds, false otherwise.</returns>
        public bool DeleteKey(string section, string key)
        {
            return (WritePrivateProfileString(section, key, null, Filename) != 0);
        }
        /// <summary>Deletes a key from the active section.</summary>
        /// <param name="key">The key to delete.</param>
        /// <returns>Returns true if the function succeeds, false otherwise.</returns>
        public bool DeleteKey(string key)
        {
            return (WritePrivateProfileString(Section, key, null, Filename) != 0);
        }
        /// <summary>Deletes a section from an INI file.</summary>
        /// <param name="section">The section to delete.</param>
        /// <returns>Returns true if the function succeeds, false otherwise.</returns>
        public bool DeleteSection(string section)
        {
            return WritePrivateProfileSection(section, null, Filename) != 0;
        }
        /// <summary>Retrieves a list of all available sections in the INI file.</summary>
        /// <returns>Returns an ArrayList with all available sections.</returns>
        public ArrayList GetSectionNames()
        {
            try
            {
                byte[] buffer = new byte[MAX_ENTRY];
                GetPrivateProfileSectionNames(buffer, MAX_ENTRY, Filename);
                string[] parts = Encoding.ASCII.GetString(buffer).Trim('\0').Split('\0');
                return new ArrayList(parts);
            }
            catch { }
            return null;
        }
        //Private variables and constants
        /// <summary>
        /// Holds the full path to the INI file.
        /// </summary>
        private string m_Filename;
        /// <summary>
        /// Holds the active section name
        /// </summary>
        private string m_Section;
        /// <summary>
        /// The maximum number of bytes in a section buffer.
        /// </summary>
        private const int MAX_ENTRY = 32768;
    }

    public class Log
    {
        public static string evSource = "SKF @STB";
        static string evLog = "SKF_Application";

        public static void log(string message)
        {
            log(true, message);
        }
        public static void log(bool LogEvent, string message)
        {
            GenericTools.DebugMsg("log(" + LogEvent.ToString() + ", " + message + ")");
            System.Console.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff ") + message);
            GenericTools.WriteLog(message);
            try
            {
                if (LogEvent)
                {
                    if (!EventLog.SourceExists(evSource))
                        EventLog.CreateEventSource(evSource, evLog);
                    EventLog.WriteEntry(evSource, message);
                }
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("log(" + LogEvent.ToString() + ", " + message + "): ERROR - " + ex.Message);
            }
        }
    }

    public class DirectoryUser
    {
        /// <summary>
        /// User's directory entry.
        /// </summary>
        private DirectoryEntry DirectoryEntry;

        /// <summary>
        /// User's login name.
        /// </summary>
        public string LoginName;

        /// <summary>
        /// AD Organization Unit of given user
        /// </summary>
        public string OU;

        /// <summary>
        /// Account expiration date
        /// </summary>
        public DateTime ExpiresAt;

        /// <summary>
        /// Last password changed date
        /// </summary>
        public DateTime PassLastChange;

        /// <summary>
        /// User's first name
        /// </summary>
        public string FirstName;

        /// <summary>
        /// User's last name
        /// </summary>
        public string LastName;

        /// <summary>
        /// User's e-mail address
        /// </summary>
        public string Email;

        /// <summary>
        /// User's phone number
        /// </summary>
        public string Phone;

        /// <summary>
        /// User's mobile phone number
        /// </summary>
        public string Mobile;

        /// <summary>
        /// User's country of residence
        /// </summary>
        public string Country;

        /// <summary>
        /// Last account update date
        /// </summary>
        public DateTime LastUpdate;

        /// <summary>
        /// Last logon date
        /// </summary>
        public DateTime LastLogon;

        /// <summary>
        /// Is user data loaded from Active Directory?
        /// </summary>
        private bool IsLoaded = false;

        /// <summary>
        /// Clear data and reset it to default values
        /// </summary>
        public void Clear()
        {
            DirectoryEntry = new System.DirectoryServices.DirectoryEntry();
            LoginName = string.Empty;
            OU = string.Empty;
            ExpiresAt = new DateTime();
            PassLastChange = new DateTime();
            FirstName = string.Empty;
            LastName = string.Empty;
            Email = string.Empty;
            Phone = string.Empty;
            Mobile = string.Empty;
            Country = string.Empty;
            LastUpdate = new DateTime();
            LastLogon = new DateTime();

            IsLoaded = false;
        }

        /// <summary>
        /// Is user data loadad from AD
        /// </summary>
        /// <returns>Is loaded</returns>
        public bool Loaded() { return IsLoaded; }

        /// <summary>
        /// Loads current user data from Active Directory
        /// </summary>
        /// <returns>Successfuly loaded</returns>
        public bool Load() { return Load(WindowsIdentity.GetCurrent().Name.Substring(WindowsIdentity.GetCurrent().Name.IndexOf("\\") + 1)); }
        /// <summary>
        /// Loads user data from Active Directory
        /// </summary>
        /// <param name="sLoginName">User's login name</param>
        /// <returns>Successfuly loaded</returns>
        public bool Load(string sLoginName)
        {
            Clear();

            DirectorySearcher ds = new DirectorySearcher();

            try
            {
                string userFilter = "(&(objectCategory=user)(sAMAccountName={0}))";

                ds.SearchScope = SearchScope.Subtree;
                ds.PropertiesToLoad.Add("distinguishedName");
                ds.PropertiesToLoad.Add("mail");
                ds.PropertiesToLoad.Add("givenName");
                ds.PropertiesToLoad.Add("sn");
                ds.PropertiesToLoad.Add("telephoneNumber");
                ds.PropertiesToLoad.Add("mobile");
                ds.PropertiesToLoad.Add("pwdLastSet");
                ds.PropertiesToLoad.Add("accountExpires");
                ds.PropertiesToLoad.Add("co");
                ds.PropertiesToLoad.Add("countryCode");
                ds.PropertiesToLoad.Add("c");
                ds.PropertiesToLoad.Add("lastLogonTimestamp");
                ds.PropertiesToLoad.Add("whenChanged");

                ds.PageSize = 1;
                ds.ServerPageTimeLimit = TimeSpan.FromSeconds(2);
                ds.Filter = string.Format(userFilter, sLoginName);

                SearchResult sr = ds.FindOne();

                if (sr != null)
                {
                    IsLoaded = true;
                    LoginName = sLoginName;

                    DirectoryEntry = sr.GetDirectoryEntry();

                    /*
                    string PropertyText = string.Empty;
                    foreach (System.DirectoryServices.PropertyValueCollection Property in DirectoryEntry.Properties)
                        PropertyText += Property.PropertyName + "=" + Property.Value.ToString() + Environment.NewLine;
                    MessageBox.Show(PropertyText);
                    */

                    if (sr.Properties.Contains("mail"))
                        Email = sr.Properties["mail"][0].ToString();

                    if (sr.Properties.Contains("givenName"))
                        FirstName = sr.Properties["givenName"][0].ToString();

                    if (sr.Properties.Contains("sn"))
                        LastName = sr.Properties["sn"][0].ToString();

                    if (sr.Properties.Contains("telephoneNumber"))
                        Phone = sr.Properties["telephoneNumber"][0].ToString();

                    if (sr.Properties.Contains("mobile"))
                        Mobile = sr.Properties["mobile"][0].ToString();

                    if (sr.Properties.Contains("co"))
                        Country = sr.Properties["co"][0].ToString();

                    if (sr.Properties.Contains("accountExpires"))
                        if (Convert.ToInt64(sr.Properties["accountExpires"][0]) == 0)
                            ExpiresAt = DateTime.MaxValue;
                        else
                            ExpiresAt = new DateTime(Convert.ToInt64(sr.Properties["accountExpires"][0])).AddYears(1600);

                    if (sr.Properties.Contains("pwdLastSet"))
                        if (Convert.ToInt64(sr.Properties["pwdLastSet"][0]) == 0)
                            PassLastChange = DateTime.MinValue;
                        else
                            PassLastChange = new DateTime(Convert.ToInt64(sr.Properties["pwdLastSet"][0])).AddYears(1600);

                    if (sr.Properties.Contains("lastLogonTimestamp"))
                        if (Convert.ToInt64(sr.Properties["lastLogonTimestamp"][0]) == 0)
                            LastLogon = DateTime.MinValue;
                        else
                            LastLogon = new DateTime(Convert.ToInt64(sr.Properties["lastLogonTimestamp"][0])).AddYears(1600);

                    LastUpdate = DateTime.MinValue;
                    if (sr.Properties.Contains("whenChanged"))
                        if (!string.IsNullOrEmpty(sr.Properties["whenChanged"][0].ToString()))
                            LastUpdate = Convert.ToDateTime(sr.Properties["whenChanged"][0]);

                    string userDistinguishedName = sr.Properties["distinguishedName"][0].ToString();
                    string tOU = userDistinguishedName.Substring(userDistinguishedName.IndexOf("OU="));
                    OU = tOU.Substring(tOU.IndexOf("OU=") + 3, tOU.IndexOf(",") - (tOU.IndexOf("OU=") + 3));
                    //string OUcoded = OUshort;
                    //if (OUshort.IndexOf(" ") > 0) OUcoded = OUshort.Substring(0, OUshort.IndexOf(" "));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }

            ds.Dispose();

            return IsLoaded;
        }

        public bool Save()
        {
            bool IsSaved = false;

            if (IsLoaded)
            {
                try
                {
                    DirectoryEntry EntryToUpdate = DirectoryEntry;
                    EntryToUpdate.Properties["mail"].Value = Email;
                    EntryToUpdate.Properties["givenName"].Value = FirstName;
                    EntryToUpdate.Properties["sn"].Value = LastName;
                    EntryToUpdate.Properties["telephoneNumber"].Value = Phone;
                    EntryToUpdate.Properties["mobile"].Value = Mobile;

                    DataRow[] Countries = CountryCodes().Select("Name='" + Country + "'");

                    if (Countries.GetLength(0) > 0)
                    {
                        EntryToUpdate.Properties["co"].Value = Countries[0]["Name"];
                        EntryToUpdate.Properties["countryCode"].Value = Countries[0]["Number"];
                        EntryToUpdate.Properties["c"].Value = Countries[0]["Code"];
                    }
                    EntryToUpdate.CommitChanges();
                    EntryToUpdate.Dispose();
                    IsSaved = Load(LoginName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
            }

            return IsSaved;
        }

        public DataTable CountryCodes()
        {
            DataTable Countries = new DataTable();
            Countries.Columns.Add("Name");
            Countries.Columns.Add("Code");
            Countries.Columns.Add("Number");

            Countries.Rows.Add("Afghanistan", "AF", 4);
            Countries.Rows.Add("Åland Islands", "AX", 248);
            Countries.Rows.Add("Albania", "AL", 8);
            Countries.Rows.Add("Algeria", "DZ", 12);
            Countries.Rows.Add("American Samoa", "AS", 16);
            Countries.Rows.Add("Andorra", "AD", 20);
            Countries.Rows.Add("Angola", "AO", 24);
            Countries.Rows.Add("Anguilla", "AI", 660);
            Countries.Rows.Add("Antarctica", "AQ", 10);
            Countries.Rows.Add("Antigua and Barbuda", "AG", 28);
            Countries.Rows.Add("Argentina", "AR", 32);
            Countries.Rows.Add("Armenia", "AM", 51);
            Countries.Rows.Add("Aruba", "AW", 533);
            Countries.Rows.Add("Australia", "AU", 36);
            Countries.Rows.Add("Austria", "AT", 40);
            Countries.Rows.Add("Azerbaijan", "AZ", 31);
            Countries.Rows.Add("Bahamas", "BS", 44);
            Countries.Rows.Add("Bahrain", "BH", 48);
            Countries.Rows.Add("Bangladesh", "BD", 50);
            Countries.Rows.Add("Barbados", "BB", 52);
            Countries.Rows.Add("Belarus", "BY", 112);
            Countries.Rows.Add("Belgium", "BE", 56);
            Countries.Rows.Add("Belize", "BZ", 84);
            Countries.Rows.Add("Benin", "BJ", 204);
            Countries.Rows.Add("Bermuda", "BM", 60);
            Countries.Rows.Add("Bhutan", "BT", 64);
            Countries.Rows.Add("Bolivia, Plurinational State of", "BO", 68);
            Countries.Rows.Add("Bonaire, Sint Eustatius and Saba", "BQ", 535);
            Countries.Rows.Add("Bosnia and Herzegovina", "BA", 70);
            Countries.Rows.Add("Botswana", "BW", 72);
            Countries.Rows.Add("Bouvet Island", "BV", 74);
            Countries.Rows.Add("Brazil", "BR", 76);
            Countries.Rows.Add("British Indian Ocean Territory", "IO", 86);
            Countries.Rows.Add("Brunei Darussalam", "BN", 96);
            Countries.Rows.Add("Bulgaria", "BG", 100);
            Countries.Rows.Add("Burkina Faso", "BF", 854);
            Countries.Rows.Add("Burundi", "BI", 108);
            Countries.Rows.Add("Cambodia", "KH", 116);
            Countries.Rows.Add("Cameroon", "CM", 120);
            Countries.Rows.Add("Canada", "CA", 124);
            Countries.Rows.Add("Cape Verde", "CV", 132);
            Countries.Rows.Add("Cayman Islands", "KY", 136);
            Countries.Rows.Add("Central African Republic", "CF", 140);
            Countries.Rows.Add("Chad", "TD", 148);
            Countries.Rows.Add("Chile", "CL", 152);
            Countries.Rows.Add("China", "CN", 156);
            Countries.Rows.Add("Christmas Island", "CX", 162);
            Countries.Rows.Add("Cocos (Keeling) Islands", "CC", 166);
            Countries.Rows.Add("Colombia", "CO", 170);
            Countries.Rows.Add("Comoros", "KM", 174);
            Countries.Rows.Add("Congo", "CG", 178);
            Countries.Rows.Add("Congo, the Democratic Republic of the", "CD", 180);
            Countries.Rows.Add("Cook Islands", "CK", 184);
            Countries.Rows.Add("Costa Rica", "CR", 188);
            Countries.Rows.Add("Côte d'Ivoire", "CI", 384);
            Countries.Rows.Add("Croatia", "HR", 191);
            Countries.Rows.Add("Cuba", "CU", 192);
            Countries.Rows.Add("Curaçao", "CW", 531);
            Countries.Rows.Add("Cyprus", "CY", 196);
            Countries.Rows.Add("Czech Republic", "CZ", 203);
            Countries.Rows.Add("Denmark", "DK", 208);
            Countries.Rows.Add("Djibouti", "DJ", 262);
            Countries.Rows.Add("Dominica", "DM", 212);
            Countries.Rows.Add("Dominican Republic", "DO", 214);
            Countries.Rows.Add("Ecuador", "EC", 218);
            Countries.Rows.Add("Egypt", "EG", 818);
            Countries.Rows.Add("El Salvador", "SV", 222);
            Countries.Rows.Add("Equatorial Guinea", "GQ", 226);
            Countries.Rows.Add("Eritrea", "ER", 232);
            Countries.Rows.Add("Estonia", "EE", 233);
            Countries.Rows.Add("Ethiopia", "ET", 231);
            Countries.Rows.Add("Falkland Islands (Malvinas)", "FK", 238);
            Countries.Rows.Add("Faroe Islands", "FO", 234);
            Countries.Rows.Add("Fiji", "FJ", 242);
            Countries.Rows.Add("Finland", "FI", 246);
            Countries.Rows.Add("France", "FR", 250);
            Countries.Rows.Add("French Guiana", "GF", 254);
            Countries.Rows.Add("French Polynesia", "PF", 258);
            Countries.Rows.Add("French Southern Territories", "TF", 260);
            Countries.Rows.Add("Gabon", "GA", 266);
            Countries.Rows.Add("Gambia", "GM", 270);
            Countries.Rows.Add("Georgia", "GE", 268);
            Countries.Rows.Add("Germany", "DE", 276);
            Countries.Rows.Add("Ghana", "GH", 288);
            Countries.Rows.Add("Gibraltar", "GI", 292);
            Countries.Rows.Add("Greece", "GR", 300);
            Countries.Rows.Add("Greenland", "GL", 304);
            Countries.Rows.Add("Grenada", "GD", 308);
            Countries.Rows.Add("Guadeloupe", "GP", 312);
            Countries.Rows.Add("Guam", "GU", 316);
            Countries.Rows.Add("Guatemala", "GT", 320);
            Countries.Rows.Add("Guernsey", "GG", 831);
            Countries.Rows.Add("Guinea", "GN", 324);
            Countries.Rows.Add("Guinea-Bissau", "GW", 624);
            Countries.Rows.Add("Guyana", "GY", 328);
            Countries.Rows.Add("Haiti", "HT", 332);
            Countries.Rows.Add("Heard Island and McDonald Islands", "HM", 334);
            Countries.Rows.Add("Holy See (Vatican City State)", "VA", 336);
            Countries.Rows.Add("Honduras", "HN", 340);
            Countries.Rows.Add("Hong Kong", "HK", 344);
            Countries.Rows.Add("Hungary", "HU", 348);
            Countries.Rows.Add("Iceland", "IS", 352);
            Countries.Rows.Add("India", "IN", 356);
            Countries.Rows.Add("Indonesia", "ID", 360);
            Countries.Rows.Add("Iran, Islamic Republic of", "IR", 364);
            Countries.Rows.Add("Iraq", "IQ", 368);
            Countries.Rows.Add("Ireland", "IE", 372);
            Countries.Rows.Add("Isle of Man", "IM", 833);
            Countries.Rows.Add("Israel", "IL", 376);
            Countries.Rows.Add("Italy", "IT", 380);
            Countries.Rows.Add("Jamaica", "JM", 388);
            Countries.Rows.Add("Japan", "JP", 392);
            Countries.Rows.Add("Jersey", "JE", 832);
            Countries.Rows.Add("Jordan", "JO", 400);
            Countries.Rows.Add("Kazakhstan", "KZ", 398);
            Countries.Rows.Add("Kenya", "KE", 404);
            Countries.Rows.Add("Kiribati", "KI", 296);
            Countries.Rows.Add("Korea, Democratic People's Republic of", "KP", 408);
            Countries.Rows.Add("Korea, Republic of", "KR", 410);
            Countries.Rows.Add("Kuwait", "KW", 414);
            Countries.Rows.Add("Kyrgyzstan", "KG", 417);
            Countries.Rows.Add("Lao People's Democratic Republic", "LA", 418);
            Countries.Rows.Add("Latvia", "LV", 428);
            Countries.Rows.Add("Lebanon", "LB", 422);
            Countries.Rows.Add("Lesotho", "LS", 426);
            Countries.Rows.Add("Liberia", "LR", 430);
            Countries.Rows.Add("Libya", "LY", 434);
            Countries.Rows.Add("Liechtenstein", "LI", 438);
            Countries.Rows.Add("Lithuania", "LT", 440);
            Countries.Rows.Add("Luxembourg", "LU", 442);
            Countries.Rows.Add("Macao", "MO", 446);
            Countries.Rows.Add("Macedonia, The Former Yugoslav Republic of", "MK", 807);
            Countries.Rows.Add("Madagascar", "MG", 450);
            Countries.Rows.Add("Malawi", "MW", 454);
            Countries.Rows.Add("Malaysia", "MY", 458);
            Countries.Rows.Add("Maldives", "MV", 462);
            Countries.Rows.Add("Mali", "ML", 466);
            Countries.Rows.Add("Malta", "MT", 470);
            Countries.Rows.Add("Marshall Islands", "MH", 584);
            Countries.Rows.Add("Martinique", "MQ", 474);
            Countries.Rows.Add("Mauritania", "MR", 478);
            Countries.Rows.Add("Mauritius", "MU", 480);
            Countries.Rows.Add("Mayotte", "YT", 175);
            Countries.Rows.Add("Mexico", "MX", 484);
            Countries.Rows.Add("Micronesia, Federated States of", "FM", 583);
            Countries.Rows.Add("Moldova, Republic of", "MD", 498);
            Countries.Rows.Add("Monaco", "MC", 492);
            Countries.Rows.Add("Mongolia", "MN", 496);
            Countries.Rows.Add("Montenegro", "ME", 499);
            Countries.Rows.Add("Montserrat", "MS", 500);
            Countries.Rows.Add("Morocco", "MA", 504);
            Countries.Rows.Add("Mozambique", "MZ", 508);
            Countries.Rows.Add("Myanmar", "MM", 104);
            Countries.Rows.Add("Namibia", "NA", 516);
            Countries.Rows.Add("Nauru", "NR", 520);
            Countries.Rows.Add("Nepal", "NP", 524);
            Countries.Rows.Add("Netherlands", "NL", 528);
            Countries.Rows.Add("New Caledonia", "NC", 540);
            Countries.Rows.Add("New Zealand", "NZ", 554);
            Countries.Rows.Add("Nicaragua", "NI", 558);
            Countries.Rows.Add("Niger", "NE", 562);
            Countries.Rows.Add("Nigeria", "NG", 566);
            Countries.Rows.Add("Niue", "NU", 570);
            Countries.Rows.Add("Norfolk Island", "NF", 574);
            Countries.Rows.Add("Northern Mariana Islands", "MP", 580);
            Countries.Rows.Add("Norway", "NO", 578);
            Countries.Rows.Add("Oman", "OM", 512);
            Countries.Rows.Add("Pakistan", "PK", 586);
            Countries.Rows.Add("Palau", "PW", 585);
            Countries.Rows.Add("Palestinian Territory, Occupied", "PS", 275);
            Countries.Rows.Add("Panama", "PA", 591);
            Countries.Rows.Add("Papua New Guinea", "PG", 598);
            Countries.Rows.Add("Paraguay", "PY", 600);
            Countries.Rows.Add("Peru", "PE", 604);
            Countries.Rows.Add("Philippines", "PH", 608);
            Countries.Rows.Add("Pitcairn", "PN", 612);
            Countries.Rows.Add("Poland", "PL", 616);
            Countries.Rows.Add("Portugal", "PT", 620);
            Countries.Rows.Add("Puerto Rico", "PR", 630);
            Countries.Rows.Add("Qatar", "QA", 634);
            Countries.Rows.Add("Réunion", "RE", 638);
            Countries.Rows.Add("Romania", "RO", 642);
            Countries.Rows.Add("Russian Federation", "RU", 643);
            Countries.Rows.Add("Rwanda", "RW", 646);
            Countries.Rows.Add("Saint Barthélemy", "BL", 652);
            Countries.Rows.Add("Saint Helena, Ascension and Tristan da Cunha", "SH", 654);
            Countries.Rows.Add("Saint Kitts and Nevis", "KN", 659);
            Countries.Rows.Add("Saint Lucia", "LC", 662);
            Countries.Rows.Add("Saint Martin (French part)", "MF", 663);
            Countries.Rows.Add("Saint Pierre and Miquelon", "PM", 666);
            Countries.Rows.Add("Saint Vincent and the Grenadines", "VC", 670);
            Countries.Rows.Add("Samoa", "WS", 882);
            Countries.Rows.Add("San Marino", "SM", 674);
            Countries.Rows.Add("Sao Tome and Principe", "ST", 678);
            Countries.Rows.Add("Saudi Arabia", "SA", 682);
            Countries.Rows.Add("Senegal", "SN", 686);
            Countries.Rows.Add("Serbia", "RS", 688);
            Countries.Rows.Add("Seychelles", "SC", 690);
            Countries.Rows.Add("Sierra Leone", "SL", 694);
            Countries.Rows.Add("Singapore", "SG", 702);
            Countries.Rows.Add("Sint Maarten (Dutch part)", "SX", 534);
            Countries.Rows.Add("Slovakia", "SK", 703);
            Countries.Rows.Add("Slovenia", "SI", 705);
            Countries.Rows.Add("Solomon Islands", "SB", 90);
            Countries.Rows.Add("Somalia", "SO", 706);
            Countries.Rows.Add("South Africa", "ZA", 710);
            Countries.Rows.Add("South Georgia and the South Sandwich Islands", "GS", 239);
            Countries.Rows.Add("South Sudan", "SS", 728);
            Countries.Rows.Add("Spain", "ES", 724);
            Countries.Rows.Add("Sri Lanka", "LK", 144);
            Countries.Rows.Add("Sudan", "SD", 729);
            Countries.Rows.Add("Suriname", "SR", 740);
            Countries.Rows.Add("Svalbard and Jan Mayen", "SJ", 744);
            Countries.Rows.Add("Swaziland", "SZ", 748);
            Countries.Rows.Add("Sweden", "SE", 752);
            Countries.Rows.Add("Switzerland", "CH", 756);
            Countries.Rows.Add("Syrian Arab Republic", "SY", 760);
            Countries.Rows.Add("Taiwan, Province of China", "TW", 158);
            Countries.Rows.Add("Tajikistan", "TJ", 762);
            Countries.Rows.Add("Tanzania, United Republic of", "TZ", 834);
            Countries.Rows.Add("Thailand", "TH", 764);
            Countries.Rows.Add("Timor-Leste", "TL", 626);
            Countries.Rows.Add("Togo", "TG", 768);
            Countries.Rows.Add("Tokelau", "TK", 772);
            Countries.Rows.Add("Tonga", "TO", 776);
            Countries.Rows.Add("Trinidad and Tobago", "TT", 780);
            Countries.Rows.Add("Tunisia", "TN", 788);
            Countries.Rows.Add("Turkey", "TR", 792);
            Countries.Rows.Add("Turkmenistan", "TM", 795);
            Countries.Rows.Add("Turks and Caicos Islands", "TC", 796);
            Countries.Rows.Add("Tuvalu", "TV", 798);
            Countries.Rows.Add("Uganda", "UG", 800);
            Countries.Rows.Add("Ukraine", "UA", 804);
            Countries.Rows.Add("United Arab Emirates", "AE", 784);
            Countries.Rows.Add("United Kingdom", "GB", 826);
            Countries.Rows.Add("United States", "US", 840);
            Countries.Rows.Add("United States Minor Outlying Islands", "UM", 581);
            Countries.Rows.Add("Uruguay", "UY", 858);
            Countries.Rows.Add("Uzbekistan", "UZ", 860);
            Countries.Rows.Add("Vanuatu", "VU", 548);
            Countries.Rows.Add("Venezuela, Bolivarian Republic of", "VE", 862);
            Countries.Rows.Add("Viet Nam", "VN", 704);
            Countries.Rows.Add("Virgin Islands, British", "VG", 92);
            Countries.Rows.Add("Virgin Islands, U.S.", "VI", 850);
            Countries.Rows.Add("Wallis and Futuna", "WF", 876);
            Countries.Rows.Add("Western Sahara", "EH", 732);
            Countries.Rows.Add("Yemen", "YE", 887);
            Countries.Rows.Add("Zambia", "ZM", 894);
            Countries.Rows.Add("Zimbabwe", "ZW", 716);

            return Countries;
        }
    }
}
