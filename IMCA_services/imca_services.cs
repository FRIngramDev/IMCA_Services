using Microsoft.VisualBasic;
using Newtonsoft.Json;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.ServiceProcess;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;
using Task = System.Threading.Tasks.Task;


namespace IMCA_Services
{
    public partial class imca_services : ServiceBase
    {
        private string service_path;
        private System.Timers.Timer timer_check_TODO;
        private int nb_sec = 10;
        private string sql_con;
        private string session_name = "";
        private string EXECUTE_IMCA_ACTION_USING_THREAD = "";
        private string Number_of_errors_allowed_per_query = "";
        private string NUMBER_OF_THREAD_MAX = "";
        private Int16 nb_created_thread = 0;
        private Boolean timer_check_TODO_is_running = false;
        private string logs_folder = "";
        private string temp_folder = "";

        public class JSON_file_config_services
        {
            public int timer_interval { get; set; }
            public string logs_folder { get; set; }
            public string session_name { get; set; }
            public string SQL_CONNECTION_STRING { get; set; }
            public string Number_of_errors_allowed_per_query { get; set; }
            public string EXECUTE_IMCA_ACTION_USING_THREAD { get; set; }
            public string NUMBER_OF_THREAD_MAX { get; set; }
            public string temp_folder { get; set; }

        }
        public imca_services()
        {
            InitializeComponent();
        }

        public void OnDebug()
        {
            OnStart(null);
        }


        private string Encrypt(string plainText)
        {
            // Secret Key.
            string secretKey = "QHiv6qHSSZ6VeunQ1UB+xw==";

            // Secret Bytes.
            byte[] secretBytes = Encoding.UTF8.GetBytes(secretKey);

            // Plain Text Bytes.
            byte[] plainTextBytes = Encoding.UTF8.GetBytes(plainText);

            // Encrypt with AES Alogorithm using Secret Key.
            using (var aes = Aes.Create())
            {
                aes.Key = secretBytes;
                aes.Mode = CipherMode.ECB;
                aes.Padding = PaddingMode.PKCS7;
                byte[] encryptedBytes = null;
                using (var encryptor = aes.CreateEncryptor())
                {
                    encryptedBytes = encryptor.TransformFinalBlock(plainTextBytes, 0, plainTextBytes.Length);
                }
                return Convert.ToBase64String(encryptedBytes);
            }
        }

        private string Decrypt(string encryptedText)
        {
            // Secret Key.
            string secretKey = "QHiv6qHSSZ6VeunQ1UB+xw==";

            // Secret Bytes.
            byte[] secretBytes = Encoding.UTF8.GetBytes(secretKey);

            // Encrypted Bytes.
            byte[] encryptedBytes = Convert.FromBase64String(encryptedText);

            // Decrypt with AES Alogorithm using Secret Key.
            using (var aes = Aes.Create())
            {
                aes.Key = secretBytes;
                aes.Mode = CipherMode.ECB;
                aes.Padding = PaddingMode.PKCS7;
                byte[] decryptedBytes = null;
                using (var decryptor = aes.CreateDecryptor())
                {
                    decryptedBytes = decryptor.TransformFinalBlock(encryptedBytes, 0, encryptedBytes.Length);
                }
                return Encoding.UTF8.GetString(decryptedBytes);
            }
        }
        
        protected override void OnStart(string[] args)
        {
            var param = new JSON_file_config_services();
            string json_file = "";

            service_path = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
            //param.logs_folder = "logs";

            try
            {
                json_file = System.IO.File.ReadAllText(service_path + "\\IMCA_services.json");

                // We read the config file

                param = JsonConvert.DeserializeObject<JSON_file_config_services>(json_file);
                sql_con = param.SQL_CONNECTION_STRING;
                nb_sec = param.timer_interval;
                session_name = param.session_name;
                Number_of_errors_allowed_per_query = param.Number_of_errors_allowed_per_query;
                EXECUTE_IMCA_ACTION_USING_THREAD = param.EXECUTE_IMCA_ACTION_USING_THREAD;
                NUMBER_OF_THREAD_MAX = param.NUMBER_OF_THREAD_MAX;
                logs_folder = param.logs_folder;
                temp_folder = param.temp_folder;


                if (session_name is null)
                {
                    session_name = "";
                }

                if (EXECUTE_IMCA_ACTION_USING_THREAD is null)
                {
                    EXECUTE_IMCA_ACTION_USING_THREAD = "FALSE";
                }

                if (Number_of_errors_allowed_per_query is null)
                {
                    Number_of_errors_allowed_per_query = "10";
                }

                if (NUMBER_OF_THREAD_MAX is null)
                {
                    NUMBER_OF_THREAD_MAX = "10";
                }


                // Delete old logs files
                class_dev_tools.Fonction.DeleteFichiers(service_path + "\\" + logs_folder, -7);

                WriteToFile("IMCA Services started at                   : " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), 0, logs_folder);
                WriteToFile("   Session Name                            : " + session_name.ToUpper(), 0, logs_folder);
                WriteToFile("   Timer Interval                          : " + nb_sec + " second(s)", 0, logs_folder);
                WriteToFile("   Connection String                       : " + sql_con, 0, logs_folder);
                WriteToFile("   Number of Errors Allowed per query      : " + Number_of_errors_allowed_per_query, 0, logs_folder);
                WriteToFile("   Execute IMCA Action using Thread        : " + EXECUTE_IMCA_ACTION_USING_THREAD, 0, logs_folder);
                WriteToFile("   Mumber of Thread Max                    : " + NUMBER_OF_THREAD_MAX, 0, logs_folder);


                if (session_name.ToString().Trim().ToUpper() != "ONLY_TO_START_SESSIONS")
                {
                    try
                    {
                        start_timer(logs_folder, temp_folder);
                    }
                    catch (Exception e)
                    {
                        WriteToFile("IMCA Services not started (Unable to start Thread) - Error : " + e.Message + " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), 0, logs_folder);
                        this.Stop();
                    }
                }
                else
                {
                    // Wait 1 minute before stopping the service
                    System.Threading.Thread.Sleep(60000);
                    System.Windows.Forms.Application.DoEvents();
                    this.Stop();
                }

            }
            catch (Exception ex)
            {
                WriteToFile("IMCA Services not started - OnStart procedure Issue -  " + ex.Message + " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), 0, logs_folder);
                this.Stop();
            }

        }
        protected override void OnStop()
        {
            while (timer_check_TODO_is_running == true)
            {
                System.Threading.Thread.Sleep(500);
                System.Windows.Forms.Application.DoEvents();
            }

            if (timer_check_TODO != null)
            {
                timer_check_TODO.Stop();
                timer_check_TODO.Dispose();
            }

            if (EXECUTE_IMCA_ACTION_USING_THREAD.Trim().ToUpper() == "TRUE")
            {
                WriteToFile("Before stopping the service we check if threads are running at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));

                Boolean bol_thread = true;
                int nb_running_thread = 0;

                using (SqlConnection con = new SqlConnection(sql_con))
                {
                    con.Open();

                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.Connection = con;
                        cmd.CommandTimeout = 0;

                        while (bol_thread)
                        {
                            cmd.CommandText = "select count(*) from [IMCA_BACKOFFICE].[dbo].PCM_TAB_IMCA_ACTION where TODO_BY='THREAD_" + session_name.ToUpper() + "'";
                            nb_running_thread = int.Parse(cmd.ExecuteScalar().ToString());

                            if (nb_running_thread > 0)
                            {
                                WriteToFile(nb_running_thread.ToString() + " Thread(s) is(are) running. We wait 10 secs at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                                System.Threading.Thread.Sleep(10000);
                                System.Windows.Forms.Application.DoEvents();
                            }
                            else
                            {
                                bol_thread = false;
                            }
                        }
                    }

                    con.Close();
                }

            }

            WriteToFile("IMCA Services stopped at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
        }
        public void start_timer(string logs, string temp_folder)
        {
            timer_check_TODO = new System.Timers.Timer();
            timer_check_TODO.Elapsed += delegate { OnElaspedTime(logs, temp_folder); };
            timer_check_TODO.Interval = nb_sec * 1000;
            timer_check_TODO.Enabled = true;
        }
        private void OnElaspedTime(string logs, string temp_folder)
        {
            WriteToFile("   Start Checking TODO was ran at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));

            try
            {
                //timer_check_TODO.Enabled = false;
                timer_check_TODO.Stop();

                check_if_TODO(logs, temp_folder);

                timer_check_TODO.Start();

                //timer_check_TODO.Enabled = true;
            }
            catch (Exception ex)
            {
                WriteToFile("Exception : " + ex + " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
            }

            WriteToFile("   End Checking TODO was ran at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));

        }

        public void InvokeOnNewThread(MethodInfo mi, object target, params object[] parameters)
        {
            ThreadStart threadMain = delegate () { mi.Invoke(target, parameters); };
            new System.Threading.Thread(threadMain).Start();
        }

        private void caller(String Namespace_myclass_mymethod, object[] param)
        {
            string Namespace = Namespace_myclass_mymethod.Split('.')[0];
            string myclass = Namespace_myclass_mymethod.Split('.')[1];
            string mymethod = Namespace_myclass_mymethod.Split('.')[2];

            string item = Namespace + "." + myclass;
            Type className = Type.GetType(item);
            MethodInfo m = className.GetMethod(mymethod);
            Object obj = Activator.CreateInstance(className);

            //InvokeOnNewThread(m, obj, param);

            object value = m.Invoke(obj, param);

        }

        private string get_IMCA_paramters(string param_name)
        {
            string ret = "";

            using (SqlConnection con = new SqlConnection(sql_con))
            {
                con.Open();

                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.Connection = con;
                    cmd.CommandTimeout = 0;

                    cmd.CommandText = "SELECT isnull(VALUE,'') as VALUE from [PCM_TAB_IMCA_PARAMETER_GLOBAL] where SK_VALID=0 AND PARAMETER='" + param_name + "'";
                    ret = cmd.ExecuteScalar().ToString();
                }

                con.Close();
            }

            return ret;

        }

        public bool TodoDateInWorkTime(DateTime dteTodoDate, DateTime dteSys, string START_TIME, string END_TIME)
        {
            bool TodoDateInWorkTime;

            DateTime WORKBEGIN = DateTime.Parse(dteSys.Year + "-" + Strings.Right("00" + dteSys.Month.ToString(), 2) + "-" + Strings.Right("00" + dteSys.Day.ToString(), 2) + " " + START_TIME + ":00.000");
            DateTime WORKEND = DateTime.Parse(dteSys.Year + "-" + Strings.Right("00" + dteSys.Month.ToString(), 2) + "-" + Strings.Right("00" + dteSys.Day.ToString(), 2) + " " + END_TIME + ":00.000");


            if (DateTime.Compare(dteTodoDate, dteSys) >= 0)
            {
                if (TimeSpan.Compare(WORKBEGIN.TimeOfDay, dteTodoDate.TimeOfDay) <= 0)
                {
                    if (TimeSpan.Compare(WORKEND.TimeOfDay, dteTodoDate.TimeOfDay) >= 0)
                    {
                        TodoDateInWorkTime = true;
                    }
                    else
                    {
                        TodoDateInWorkTime = false;
                    }
                }
                else
                {
                    TodoDateInWorkTime = false;
                }
            }
            else
            {
                TodoDateInWorkTime = false;
            }

            return TodoDateInWorkTime;
        }
        public long CurrentSKTime(string strType = "CD", string strISODate = "")
        {
            long ret = 0;

            List<string> list = new List<string>() { "CD", "FD", "FW", "CM", "FM", "CQ", "FQ", "CY", "FY" };


            if (!list.Contains(strType))
            {
                strType = "CD";
            }

            if ((!strISODate.All(char.IsNumber)) || strISODate == "")
            {
                strISODate = DateAndTime.Now.ToString("yyyyMMdd");
            }


            using (SqlConnection con = new SqlConnection(sql_con))
            {
                con.Open();

                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.Connection = con;
                    cmd.CommandTimeout = 0;

                    cmd.CommandText = "SELECT SKT_" + strType.ToString().ToUpper() + " FROM IMCA_BACKOFFICE.dbo.DSSIMPRT_TAB_SK_TIME WHERE DATESTART = CONVERT(DATETIME, '" + strISODate + "', 112) AND TIMETYPE = 'CD'";
                    ret = (int)cmd.ExecuteScalar();

                    if (ret == 0)
                    {
                        ret = -1;
                    }

                    cmd.Dispose();

                }

                con.Close();

            }

            return ret;

        }
        public void AddIMCAAction(DateTime dtSysSql, long lngSK_VALID, string strAction, long lngPriority = -9999, DateTime dteTodo = default(DateTime), string strCreatedFrom = "", string strTodoBy = "", string strSK_TYPE = "", string strParam01 = "", string strParam02 = "", string strParam03 = "", string strParam04 = "", string strParam05 = "", string strParam06 = "", string strParam07 = "", string strParam08 = "", string strParam09 = "", string strParam10 = "", string strParam11 = "", string strParam12 = "", string strParam13 = "", string strParam14 = "", string strParam15 = "", string strParam16 = "", string strParam17 = "", string strParam18 = "", string strParam19 = "", string strParam20 = "", string strParamMemo = "", string lngCreatedFromAuto = "0", long lngSK_Value = 0)
        {
            try
            {
                if (strTodoBy == "" || lngPriority == -9999)
                {
                    using (SqlConnection con = new SqlConnection(sql_con))
                    {
                        con.Open();

                        using (SqlCommand cmd = new SqlCommand())
                        {
                            cmd.Connection = con;
                            cmd.CommandTimeout = 0;
                            cmd.CommandText = "SELECT NEXT_RUN_TODO_BY,DEFAULT_PRIORITY FROM IMCA_BACKOFFICE.dbo.PCM_TAB_IMCA_ACTION_ADMIN WHERE ACTION='" + strAction + "'";

                            SqlDataReader dr;

                            dr = cmd.ExecuteReader();
                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    if (strTodoBy == "")
                                    {
                                        strTodoBy = dr["NEXT_RUN_TODO_BY"].ToString();
                                    }

                                    if (lngPriority == -9999)
                                    {
                                        lngPriority = (int)dr["DEFAULT_PRIORITY"];
                                    }
                                }

                            }
                            else
                            {
                                if (strTodoBy == "")
                                {
                                    strTodoBy = "TODO";
                                }

                                if (lngPriority == -9999)
                                {
                                    lngPriority = 999;
                                }
                            }

                            dr.Close();


                            if (dteTodo == default(DateTime))
                            {
                                dteTodo = dtSysSql;
                            }

                            if (strCreatedFrom == "")
                            {
                                strCreatedFrom = Environment.UserName;
                            }

                            if (lngSK_Value == 0)
                            {
                                switch (strSK_TYPE)
                                {
                                    case "SK_CUST":

                                        break;

                                    case "SK_SKU":

                                        break;

                                    case "SK_VEND":

                                        break;

                                }

                                strParam01 = strParam01.ToUpper();

                            }

                            strParam01 = strParam01.Replace("'", "''");
                            strParam02 = strParam02.Replace("'", "''");
                            strParam03 = strParam03.Replace("'", "''");
                            strParam04 = strParam04.Replace("'", "''");
                            strParam05 = strParam05.Replace("'", "''");
                            strParam06 = strParam06.Replace("'", "''");
                            strParam07 = strParam07.Replace("'", "''");
                            strParam08 = strParam08.Replace("'", "''");
                            strParam09 = strParam09.Replace("'", "''");
                            strParam10 = strParam10.Replace("'", "''");
                            strParam11 = strParam11.Replace("'", "''");
                            strParam12 = strParam12.Replace("'", "''");
                            strParam13 = strParam13.Replace("'", "''");
                            strParam14 = strParam14.Replace("'", "''");
                            strParam15 = strParam15.Replace("'", "''");
                            strParam16 = strParam16.Replace("'", "''");
                            strParam17 = strParam17.Replace("'", "''");
                            strParam18 = strParam18.Replace("'", "''");
                            strParam19 = strParam19.Replace("'", "''");
                            strParam20 = strParam20.Replace("'", "''");


                            strParamMemo = strParamMemo.Replace("'", "''");
                            strTodoBy = strTodoBy.Replace("'", "''");
                            strCreatedFrom = strCreatedFrom.Replace("'", "''");


                            string strSQL;

                            strSQL = "INSERT INTO dbo.PCM_TAB_IMCA_ACTION (SK_VALID, SK_TYPE, SK_VALUE, PRIORITY, ACTION, PARAM01, PARAM02, PARAM03, PARAM04, PARAM05, PARAM06, PARAM07, PARAM08, PARAM09, PARAM10, PARAM11, PARAM12, PARAM13, PARAM14, PARAM15, PARAM16, PARAM17, PARAM18, PARAM19, PARAM20, PARAM_MEMO, TODO_DATE, TODO_DATE_INT, TODO_BY, CREATED_FROM, CREATED, SK_CREATE_DATE, SK_FINISH_DATE, CREATED_FROM_AUTO) " +
                                           "VALUES (" + lngSK_VALID + ", '" + strSK_TYPE + "', " + lngSK_Value + ", " + lngPriority + ", '" + strAction + "', '" + strParam01 + "', '" + strParam02 + "', '" + strParam03 + "', '" + strParam04 + "', '" + strParam05 + "', '" + strParam06 + "', '" + strParam07 + "', '" + strParam08 + "', '" + strParam09 + "', '" + strParam10 + "', '" + strParam11 + "', '" + strParam12 + "', '" + strParam13 + "', '" + strParam14 + "', " +
                                           "'" + strParam15 + "', '" + strParam16 + "', '" + strParam17 + "', '" + strParam18 + "', '" + strParam19 + "', '" + strParam20 + "', '" + strParamMemo + "', '" + dteTodo.ToString("yyyy-MM-dd HH:mm:ss") + "', " + dteTodo.ToString("yyyyMMddHHmmss") + ", '" + strTodoBy + "', '" + strCreatedFrom + "',getdate()," + CurrentSKTime().ToString() + ", 0, " + lngCreatedFromAuto.ToString() + ")";

                            cmd.CommandText = strSQL;
                            cmd.ExecuteNonQuery();

                        }

                        con.Close();

                    }
                }
            }
            catch (Exception ex)
            {
                Debug.Print(ex.Message);
            }


        }

        private string get_IMCA_paramters(string sql_con, string param_name)
        {
            string ret = "";

            using (SqlConnection con = new SqlConnection(sql_con))
            {
                con.Open();

                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.Connection = con;
                    cmd.CommandTimeout = 0;

                    cmd.CommandText = "SELECT isnull(VALUE,'') as VALUE from [PCM_TAB_IMCA_PARAMETER_GLOBAL] where SK_VALID=0 AND PARAMETER='" + param_name + "'";
                    try
                    {
                        ret = cmd.ExecuteScalar().ToString();
                    }
                    catch (Exception)
                    {
                        ret = "";
                    }

                }

                con.Close();
            }

            return ret;

        }


        public void SendMail(String strMailTo = "", String strMailSubject = "", String strMailBody = "", String strMailCc = "", String strMailBcc = "", String strMailAttachments = "", Int32 lngMailPriority = -9999, String strMailFrom = "", DateTime? dteTodo = null, String strMailType = "MAIL", String strErrorMailTo = "", Int16 lngMailBodyFormat = 1, String strDefaultMailBody = "", String strTodoBy = "TODO", Int32 lngSkValid = -1, String strSKType = "", Int32 lngSKValue = 0)
        {
            String tmp;
            tmp = strMailTo;
        }

        protected string Execute_SendMail(string sendmail_function, long ACTION_ID, long FUNCTION_ID, Int16 sk_valid)
        {
            string ret = "";
            List<string> list_args = new List<string>();
            RegexOptions options = RegexOptions.Multiline;

            var func = Regex.Match(sendmail_function, @"\b[^()]+\((.*)\)$");

            string list_parameters = func.Groups[1].Value;

            while (list_parameters.IndexOf(",,") != -1)
            {
                list_parameters = list_parameters.Replace(",,", ",\"OPTIONAL_VALUE_NOT_PROVIDED\",");
            }

            ///
            /// Get parameters
            /// 
            var paramTags = Regex.Matches(list_parameters, @"([^,]+\(.+?\))|([^,]+)", options);

            foreach (var item in paramTags)
            {
                list_args.Add(item.ToString().Trim().TrimStart('\"').TrimEnd('\"'));
            }
            if (list_parameters != "")

            {
                try
                {

                    //
                    ////  Parameters List
                    //
                    //  Parameter 1 : strMailTo As String
                    //  Parameter 2 : strMailSubject As String
                    //  Parameter 3 : strMailBody As String
                    //  Parameter 4 : Optional strMailCc As String
                    //  Parameter 5 : Optional strMailBcc As String
                    //  Parameter 6 : Optional strMailAttachments As String,
                    //  Parameter 7 : Optional lngMailPriority As Long = -9999
                    //  Parameter 8 : Optional strMailFrom As String,
                    //  Parameter 9 : Optional dteTodo As Date
                    //  Parameter 10 :Optional strMailType As String = "MAIL"
                    //  Parameter 11 :Optional strErrorMailTo As String
                    //  Parameter 12 :Optional lngMailBodyFormat As lngEMailBodyFormat = lngFormatPlain
                    //  Parameter 13 :Optional strDefaultMailBody As String = ""
                    //  Parameter 14 :Optional strTodoBy As String = "TODO"
                    //  Parameter 15 :Optional lngSkValid As Long = -1
                    //  Parameter 16 :Optional strSKType As String = ""
                    //  Parameter 17 :Optional lngSKValue As Long = 0

                    string strMailTo = "";
                    string strMailSubject = "";
                    string strMailBody = "";
                    string strMailCc = "";
                    string strMailBcc = "";
                    string strMailAttachments = "";
                    long lngMailPriority = -9999;
                    string strMailFrom = "";
                    DateTime dteTodo = DateTime.Now;
                    string strMailType = "MAIL";
                    string strErrorMailTo = "";
                    Int16 lngMailBodyFormat = 1;
                    string strDefaultMailBody = "";
                    string strTodoBy = "TODO";
                    long lngSkValid = sk_valid;
                    string strSKType = "";
                    long lngSKValue = 0;

                    string CurrentAppID = "IMCA";
                    string CurrentUser = session_name.ToUpper();

                    for (int i = 0; i < list_args.Count; i++)
                    {

                        if (list_args[i].ToString().Trim() == "OPTIONAL_VALUE_NOT_PROVIDED")
                        {
                            list_args[i] = "";
                        }

                        switch (i)
                        {
                            case 0:
                                strMailTo = list_args[i].ToString().Trim();
                                break;
                            case 1:
                                strMailSubject = list_args[i].ToString().Trim().Replace("\" &", "").Replace("Now()", DateAndTime.Now.ToString());
                                break;
                            case 2:
                                strMailBody = list_args[i].ToString().Trim().Replace("\" &", "").Replace("Now()", DateAndTime.Now.ToString());
                                break;
                            case 3:
                                strMailCc = list_args[i].ToString().Trim();
                                break;
                            case 4:
                                strMailBcc = list_args[i].ToString().Trim();
                                break;
                            case 5:
                                strMailAttachments = list_args[i].ToString();
                                break;
                            case 6:
                                lngMailPriority = long.Parse(list_args[i].ToString());
                                break;
                            case 7:
                                strMailFrom = list_args[i].ToString();
                                break;
                            case 8:
                                dteTodo = DateTime.Parse(list_args[i].ToString());
                                break;
                            case 09:
                                strMailType = list_args[i].ToString();
                                break;
                            case 10:
                                strErrorMailTo = list_args[i].ToString();
                                break;
                            case 11:
                                lngMailBodyFormat = Int16.Parse(list_args[i].ToString());
                                break;
                            case 12:
                                strDefaultMailBody = list_args[i].ToString();
                                break;
                            case 13:
                                strTodoBy = list_args[i].ToString();
                                break;
                            case 14:
                                lngSkValid = long.Parse(list_args[i].ToString());
                                break;
                            case 15:
                                strSKType = list_args[i].ToString();
                                break;
                            case 16:
                                lngSKValue = long.Parse(list_args[i].ToString());
                                break;
                            default:
                                break;
                        }
                    }

                    if (lngMailPriority == -9999)
                    {
                        using (SqlConnection con_sql = new SqlConnection(sql_con))
                        {
                            con_sql.Open();

                            using (SqlCommand cmd = new SqlCommand())
                            {
                                cmd.Connection = con_sql;
                                cmd.CommandTimeout = 0;
                                cmd.CommandText = "SELECT TOP 1 VALUE FROM IMCA_BACKOFFICE.dbo.PCM_TAB_IMCA_PARAMETER_GLOBAL WHERE SK_VALID IN (0, " + lngSkValid.ToString() + ") AND PARAMETER IN ('MAIL_PRIORITY', 'MAIL_PRIORITY_" + strMailType + "') ORDER BY PARAMETER DESC, SK_VALID DESC";
                                lngMailPriority = long.Parse(cmd.ExecuteScalar().ToString());
                            }

                            con_sql.Close();
                        }
                    }

                    if (String.IsNullOrEmpty(strTodoBy))
                    {
                        strTodoBy = "TODO";
                    }


                    using (SqlConnection con_sql = new SqlConnection(sql_con))
                    {
                        con_sql.Open();

                        using (SqlCommand cmd = new SqlCommand())
                        {
                            cmd.Connection = con_sql;
                            cmd.CommandTimeout = 0;

                            cmd.CommandText = "INSERT INTO IMCA_BACKOFFICE.dbo.PCM_TAB_IMCA_AUTOMAIL (SK_VALID, PRIORITY, TYPE, MAIL_FROM, MAIL_TO, MAIL_CC, MAIL_BCC, MAIL_SUBJECT, MAIL_BODY, MAIL_BODY_FORMAT, DEFAULT_MAIL_BODY, MAIL_ATTACHMENTS, ERROR_MAIL_TO, TODO_DATE, TODO_DATE_INT, TODO_BY, CREATED_FROM, CREATED, ERROR, SK_CREATE_DATE, SK_FINISH_DATE, SK_TYPE, SK_VALUE) VALUES (" +
                                              lngSkValid.ToString() + ", " + lngMailPriority.ToString() + ", '" + strMailType + "', '" + strMailFrom.Replace("'", "''") + "', '" + strMailTo.Replace("'", "''") + "', '" + strMailCc.Replace("'", "''") + "', '" + strMailBcc.Replace("'", "''") + "', '" + strMailSubject.Replace("'", "''") + "', '" + strMailBody.Replace("'", "''") + "', " +
                                              lngMailBodyFormat + ", '" + strDefaultMailBody + "', '" + strMailAttachments.Replace("'", "''") + "', '" + strErrorMailTo.Replace("'", "''") + "', '" + dteTodo.ToString("yyyy-MM-dd HH:mm:ss") + "', " + dteTodo.ToString("yyyyMMddHHmmss") + ", '" + strTodoBy + "', '" + CurrentAppID.ToString() + " " + CurrentUser.ToString() + "', GETDATE(), 0, " + CurrentSKTime() + ", 0, '" + strSKType + "', " + lngSKValue.ToString() + ")";
                            cmd.ExecuteNonQuery();
                        }

                        con_sql.Close();
                    }


                }
                catch (Exception ex)
                {
                    using (SqlConnection con_sql = new SqlConnection(sql_con))
                    {
                        con_sql.Open();

                        using (SqlCommand cmd_error = new SqlCommand())
                        {
                            cmd_error.Connection = con_sql;
                            cmd_error.CommandTimeout = 0;
                            cmd_error.CommandText = "update [IMCA_BACKOFFICE].[dbo].PCM_TAB_IMCA_ACTION set TODO_BY=@TODO_BY,ERROR=@ERROR_NUMBER,ERROR_TEXT=@ERROR_MSG where ID=@ID";
                            cmd_error.Parameters.Clear();
                            cmd_error.Parameters.AddWithValue("@TODO_BY", "ERROR_" + session_name.ToUpper());
                            cmd_error.Parameters.AddWithValue("@ID", FUNCTION_ID.ToString());
                            cmd_error.Parameters.AddWithValue("@ERROR_NUMBER", 0);
                            cmd_error.Parameters.AddWithValue("@ERROR_MSG", Microsoft.VisualBasic.Strings.Left("FUNCTION_ID=" + FUNCTION_ID.ToString() + " - ERROR MSG=" + ex.Message, 1024));
                            cmd_error.ExecuteNonQuery();
                        }

                        con_sql.Close();
                    }

                    ret = " Unable to execute the function SendMail (error : " + ex.Message + ") - " + list_args[1].ToString().Trim('"').Trim();
                }
            }
            else
            {
                ret = " Too few parameters for SendMail function - " + sendmail_function;
            }

            return ret;

        }

        protected string Execute_Query(string strSQL, long ACTION_ID, long FUNCTION_ID)
        {
            string ret = "";
            Int32 nb_error = 1;
            Int32 lngMaxErrors = Int32.Parse(Number_of_errors_allowed_per_query);
            Boolean bol_error = false;

            List<string> list_args = new List<string>();
            RegexOptions options = RegexOptions.Multiline;


            string pattern = @"([\w.$]+|""[^""]+""|'[^']+')";


            while (strSQL.IndexOf(",,") != -1)
            {
                strSQL = strSQL.Replace(",,", ",\"OPTIONAL_VALUE_NOT_PROVIDED\",");
            }

            string input = strSQL;

            foreach (Match m in Regex.Matches(input, pattern, options))
            {
                list_args.Add(m.Value);
            }


            if (list_args.Count >= 3)
            {
                // First Args contains Function name. We remove it
                list_args.RemoveAt(0);

                try
                {
                    // 2nd Parameter of EvalOpenQuery contains the name of the connexion

                    string connection_string = get_IMCA_paramters(sql_con, list_args[1].ToString().Replace(Microsoft.VisualBasic.Strings.Chr(34), ' ').Trim());
                    OracleConnectionStringBuilder builder = new OracleConnectionStringBuilder(connection_string);

                    if (connection_string != "")
                    {
                        switch (connection_string.ToUpper().Trim())
                        {
                            case string x when Strings.Replace(x, " ", "").StartsWith("DATASOURCE=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST"):

                                // ORACLE QUERY

                                OracleConnection conn_oracle = new OracleConnection(connection_string);
                                OracleTransaction transaction;

                                conn_oracle.Open();

                                using (OracleCommand command = new OracleCommand())
                                {
                                    command.Connection = conn_oracle;

                                    if (list_args.Count >= 7)
                                    {
                                        command.CommandTimeout = Int32.Parse(list_args[6].ToString());
                                    }
                                    else
                                    {
                                        command.CommandTimeout = 120;
                                    }

                                    if (list_args.Count >= 8)
                                    {
                                        lngMaxErrors = Int32.Parse(list_args[7].ToString());
                                    }

                                    transaction = conn_oracle.BeginTransaction();
                                    command.Transaction = transaction;
                                    // First Parameter of EvalOpenQuery contains the query
                                    command.CommandText = list_args[0].ToString().Trim('"');
                                    command.CommandType = CommandType.Text;


                                    try
                                    {
                                        command.ExecuteNonQuery();
                                        transaction.Commit();

                                    }
                                    catch (OracleException ex)
                                    {
                                        WriteToFile("                                   : ConnectionState : " + conn_oracle.State.ToString());

                                        if (conn_oracle.State == ConnectionState.Closed)
                                        {
                                            WriteToFile("                                   : We reopen Oracle Connection");
                                            conn_oracle.OpenWithNewPassword(builder.Password);
                                        }

                                        while (nb_error < lngMaxErrors)
                                        {
                                            try
                                            {
                                                WriteToFile("                                   : ERROR : " + ex.Number.ToString() + " - WAIT 1 MINUTE BEFORE TRYING AGAIN - " + nb_error.ToString() + " in " + lngMaxErrors.ToString() + " tries");
                                                // Wait 1 minute - 60 * 1 sec
                                                for (int sec = 1; sec <= 60; sec++)
                                                {
                                                    System.Threading.Thread.Sleep(1000);
                                                    System.Windows.Forms.Application.DoEvents();
                                                }


                                                command.ExecuteNonQuery();
                                                transaction.Commit();
                                                // We exit the loop
                                                bol_error = false;
                                                nb_error = lngMaxErrors;
                                            }
                                            catch (OracleException)
                                            {
                                                bol_error = true;
                                                nb_error += 1;
                                            }

                                        }

                                        if (bol_error == true)
                                        {
                                            WriteToFile("                                   : ERROR : " + ex.Number.ToString() + " - " + ex.Message);
                                            //transaction.Rollback();
                                            using (SqlConnection con = new SqlConnection(sql_con))
                                            {
                                                con.Open();

                                                using (SqlCommand cmd = new SqlCommand())
                                                {
                                                    cmd.Connection = con;
                                                    cmd.CommandTimeout = 0;
                                                    cmd.CommandText = "update [IMCA_BACKOFFICE].[dbo].PCM_TAB_IMCA_ACTION set TODO_BY=@TODO_BY,ERROR=@ERROR_NUMBER,ERROR_TEXT=@ERROR_MSG where ID=@ID";
                                                    cmd.Parameters.Clear();
                                                    cmd.Parameters.AddWithValue("@TODO_BY", "ERROR_" + session_name.ToUpper());
                                                    cmd.Parameters.AddWithValue("@ID", ACTION_ID);
                                                    cmd.Parameters.AddWithValue("@ERROR_NUMBER", Strings.Replace(ex.Number.ToString(), "ORA-", ""));
                                                    cmd.Parameters.AddWithValue("@ERROR_MSG", Microsoft.VisualBasic.Strings.Left("FUNCTION_ID=" + FUNCTION_ID + " - ORACLE ERROR MSG=" + ex.Message, 1024));
                                                    cmd.ExecuteNonQuery();
                                                }

                                                con.Close();
                                            }

                                            ret = " Oracle Query error (" + ex.Number.ToString() + " - " + ex.Message + ") -  function - " + strSQL;
                                        }
                                        else
                                        {
                                            ret = "";
                                        }
                                        throw;
                                    }

                                }


                                transaction.Dispose();
                                conn_oracle.Close();

                                break;

                            default:

                                // SQL QUERY

                                using (SqlConnection con = new SqlConnection(sql_con))
                                {
                                    con.Open();
                                    using (SqlCommand command = new SqlCommand())
                                    {
                                        command.Connection = con;
                                        if (list_args.Count >= 7)
                                        {
                                            command.CommandTimeout = Int32.Parse(list_args[6].ToString());
                                        }
                                        else
                                        {
                                            command.CommandTimeout = 120;
                                        }

                                        if (list_args.Count >= 8)
                                        {
                                            lngMaxErrors = Int32.Parse(list_args[7].ToString());
                                        }

                                        command.CommandText = "";
                                        command.CommandType = CommandType.Text;

                                        try
                                        {
                                            command.ExecuteNonQuery();
                                        }
                                        catch (SqlException ex)
                                        {
                                            while (nb_error < lngMaxErrors)
                                            {
                                                try
                                                {
                                                    WriteToFile("                                   : WAIT 1 MINUTE BEFORE TRYING AGAIN - " + nb_error.ToString() + " in " + lngMaxErrors.ToString() + " tries");
                                                    // Wait 1 minute - 60 * 1 sec
                                                    for (int sec = 1; sec <= 60; sec++)
                                                    {
                                                        System.Threading.Thread.Sleep(1000);
                                                        System.Windows.Forms.Application.DoEvents();
                                                    }
                                                    command.ExecuteNonQuery();
                                                    // We exit the loop
                                                    bol_error = false;
                                                    nb_error = lngMaxErrors;
                                                }
                                                catch (OracleException)
                                                {
                                                    bol_error = true;
                                                    nb_error += 1;
                                                }

                                            }

                                            if (bol_error == true)
                                            {
                                                WriteToFile("                                   : ERROR : " + ex.Number.ToString() + " - " + ex.Message);

                                                using (SqlConnection con_sql = new SqlConnection(sql_con))
                                                {
                                                    con_sql.Open();

                                                    using (SqlCommand cmd = new SqlCommand())
                                                    {
                                                        cmd.Connection = con_sql;
                                                        cmd.CommandTimeout = 0;
                                                        cmd.CommandText = "update [IMCA_BACKOFFICE].[dbo].PCM_TAB_IMCA_ACTION set TODO_BY=@TODO_BY,ERROR=@ERROR_NUMBER,ERROR_TEXT=@ERROR_MSG where ID=@ID";
                                                        cmd.Parameters.Clear();
                                                        cmd.Parameters.AddWithValue("@TODO_BY", "ERROR_" + session_name.ToUpper());
                                                        cmd.Parameters.AddWithValue("@ID", ACTION_ID);
                                                        cmd.Parameters.AddWithValue("@ERROR_NUMBER", ex.Number);
                                                        cmd.Parameters.AddWithValue("@ERROR_MSG", Microsoft.VisualBasic.Strings.Left("FUNCTION_ID=" + FUNCTION_ID + " - SQL ERROR MSG=" + ex.Message, 1024));
                                                        cmd.ExecuteNonQuery();
                                                    }

                                                    con_sql.Close();
                                                }

                                                ret = " SQL Query error (" + ex.Number.ToString() + " - " + ex.Message + ") -  function - " + strSQL;
                                            }
                                            else
                                            {
                                                ret = "";
                                            }
                                            throw;
                                        }

                                    }

                                    con.Close();
                                }

                                break;
                        }

                        ret = "";
                    }
                    else
                    {
                        ret = "Unable to find (Check the Global Parameters Table) the connection String for " + list_args[1].ToString().Replace(Microsoft.VisualBasic.Strings.Chr(34), ' ').Trim();
                    }

                }
                catch (OracleException ex)
                {
                    ret = " Unable to execute the query (error : " + ex.Message + ") - " + list_args[0].ToString().Trim('"').Trim();
                }
            }
            else
            {
                ret = " Too few parameters for EvalOpenQuery function - " + strSQL;
            }

            return ret;

        }


        protected string Execute_Query_old(string strSQL, long ACTION_ID, long FUNCTION_ID)
        {
            string ret = "";

            string pattern = @"([\w.$]+|""[^""]+""|'[^']+')";


            while (strSQL.IndexOf(",,") != -1)
            {
                strSQL = strSQL.Replace(",,", ",\"\",");
            }


            string input = strSQL;
            RegexOptions options = RegexOptions.Multiline;

            List<string> list_args = new List<string>();

            foreach (Match m in Regex.Matches(input, pattern, options))
            {
                if (m.Value != "\",\"")
                {
                    list_args.Add(m.Value);
                }
                else
                {
                    list_args.Add("");
                }
            }

            if (list_args.Count >= 2)
            {
                try
                {
                    // 2nd Parameter of EvalOpenQuery contains the name of the connexion

                    string connection_string = get_IMCA_paramters(sql_con, list_args[2].ToString().Replace(Microsoft.VisualBasic.Strings.Chr(34), ' ').Trim());

                    if (connection_string != "")
                    {
                        switch (connection_string.ToUpper().Trim())
                        {
                            case string x when Strings.Replace(x, " ", "").StartsWith("DATASOURCE=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST"):

                                // ORACLE QUERY

                                OracleConnection conn_oracle = new OracleConnection(connection_string);
                                OracleTransaction transaction;
                                conn_oracle.Open();

                                using (OracleCommand command = new OracleCommand())
                                {
                                    command.Connection = conn_oracle;

                                    if (list_args.Count >= 7)
                                    {
                                        command.CommandTimeout = Int32.Parse(list_args[7].ToString());
                                    }
                                    else
                                    {
                                        command.CommandTimeout = 120;
                                    }


                                    transaction = conn_oracle.BeginTransaction();
                                    command.Transaction = transaction;
                                    // First Parameter of EvalOpenQuery contains the query
                                    command.CommandText = list_args[1].ToString().Trim('"');
                                    command.CommandType = CommandType.Text;


                                    try
                                    {
                                        command.ExecuteNonQuery();
                                        transaction.Commit();

                                    }
                                    catch (OracleException ex)
                                    {
                                        //transaction.Rollback();
                                        using (SqlConnection con = new SqlConnection(sql_con))
                                        {
                                            con.Open();

                                            using (SqlCommand cmd = new SqlCommand())
                                            {
                                                cmd.Connection = con;
                                                cmd.CommandTimeout = 0;
                                                cmd.CommandText = "update [IMCA_BACKOFFICE].[dbo].PCM_TAB_IMCA_ACTION set TODO_BY=@TODO_BY,ERROR=@ERROR_NUMBER,ERROR_TEXT=@ERROR_MSG where ID=@ID";
                                                cmd.Parameters.Clear();
                                                cmd.Parameters.AddWithValue("@TODO_BY", "ERROR_" + session_name.ToUpper());
                                                cmd.Parameters.AddWithValue("@ID", ACTION_ID);
                                                cmd.Parameters.AddWithValue("@ERROR_NUMBER", Strings.Replace(ex.Number.ToString(), "ORA-", ""));
                                                cmd.Parameters.AddWithValue("@ERROR_MSG", Microsoft.VisualBasic.Strings.Left("FUNCTION_ID=" + FUNCTION_ID + " - ORACLE ERROR MSG=" + ex.Message, 1024));
                                                cmd.ExecuteNonQuery();
                                            }

                                            con.Close();
                                        }

                                        ret = " Oracle Query error (" + ex.Number.ToString() + " - " + ex.Message + ") -  function - " + strSQL;

                                        throw;
                                    }

                                }

                                conn_oracle.Close();

                                break;

                            default:

                                // SQL QUERY

                                using (SqlConnection con = new SqlConnection(sql_con))
                                {
                                    con.Open();

                                    using (SqlCommand command = new SqlCommand())
                                    {
                                        command.Connection = con;
                                        if (list_args.Count >= 7)
                                        {
                                            command.CommandTimeout = Int32.Parse(list_args[7].ToString());
                                        }
                                        else
                                        {
                                            command.CommandTimeout = 120;
                                        }

                                        command.CommandText = "";
                                        command.CommandType = CommandType.Text;

                                        try
                                        {
                                            command.ExecuteNonQuery();
                                        }
                                        catch (SqlException ex)
                                        {
                                            using (SqlConnection con_sql = new SqlConnection(sql_con))
                                            {
                                                con_sql.Open();

                                                using (SqlCommand cmd = new SqlCommand())
                                                {
                                                    cmd.Connection = con_sql;
                                                    cmd.CommandTimeout = 0;
                                                    cmd.CommandText = "update [IMCA_BACKOFFICE].[dbo].PCM_TAB_IMCA_ACTION set TODO_BY=@TODO_BY,ERROR=@ERROR_NUMBER,ERROR_TEXT=@ERROR_MSG where ID=@ID";
                                                    cmd.Parameters.Clear();
                                                    cmd.Parameters.AddWithValue("@TODO_BY", "ERROR_" + session_name.ToUpper());
                                                    cmd.Parameters.AddWithValue("@ID", ACTION_ID);
                                                    cmd.Parameters.AddWithValue("@ERROR_NUMBER", ex.Number);
                                                    cmd.Parameters.AddWithValue("@ERROR_MSG", Microsoft.VisualBasic.Strings.Left("FUNCTION_ID=" + FUNCTION_ID + " - SQL ERROR MSG=" + ex.Message, 1024));
                                                    cmd.ExecuteNonQuery();
                                                }

                                                con_sql.Close();
                                            }

                                            ret = " SQL Query error (" + ex.Number.ToString() + " - " + ex.Message + ") -  function - " + strSQL;
                                            throw;
                                        }

                                    }

                                    con.Close();
                                }

                                break;
                        }

                        ret = "";
                    }
                    else
                    {
                        ret = "Unable to find (Check the Global Parameters Table) the connection String for " + list_args[2].ToString().Replace(Microsoft.VisualBasic.Strings.Chr(34), ' ').Trim();
                    }

                }
                catch (OracleException ex)
                {
                    ret = " Unable to execute the query (error : " + ex.Message + ") - " + list_args[1].ToString().Trim('"').Trim();
                }
            }
            else
            {
                ret = " Too few parameters for EvalOpenQuery function - " + strSQL;
            }

            return ret;

        }

        protected string resolve_parameters(string func_name, long ACTION_ID, long FUNCTION_ID, DataRow dt)
        {
            string field_name = "";

            for (int i = 1; i <= 20; i++)
            {
                field_name = "PARAM" + i.ToString("00");
                func_name = Strings.Replace(func_name, "[" + field_name + "]", dt[field_name].ToString());
            }

            func_name = Strings.Replace(func_name, "[ACTION]", dt["ACTION"].ToString());

            return func_name;
        }


        public void execute_IMCA_Action(DataRow dr, string logs, string temp_folder, Boolean using_thread)
        {
            SqlDataReader dt2;

            if (using_thread == true)
                nb_created_thread += 1;

            long id;
            switch (using_thread)
            {
                case true:
                    id = long.Parse(dr["ID"].ToString());
                    break;

                default:
                    id = 0;
                    break;
            }

            using (SqlConnection con = new SqlConnection(sql_con))
            {
                con.Open();
                using (SqlCommand cmd2 = new SqlCommand())
                {
                    cmd2.Connection = con;
                    cmd2.CommandTimeout = 0;

                    // Check the flags if all the "Other(s)" component(s) (ORACLE etc...) are available

                    cmd2.CommandText = "SELECT FLAG from [PCM_TAB_IMCA_ACTION_FLAG] " +
                                       "INNER JOIN (SELECT * FROM [PCM_TAB_IMCA_PARAMETER_GLOBAL] where PARAMETER like 'AVAIL_%' AND [VALUE]='0' and (SK_VALID=0 or SK_VALID=" + dr["SK_VALID"].ToString() + ")) [parameters] " +
                                       "on PCM_TAB_IMCA_ACTION_FLAG.FLAG=[parameters].PARAMETER" +
                                       " WHERE [ACTION_ID]=" + dr["ADMIN_ID"].ToString();
                    dt2 = cmd2.ExecuteReader();

                    Boolean all_is_available = true;
                    if (dt2.HasRows)
                    {
                        all_is_available = false;
                    }

                    dt2.Close();

                    if (all_is_available == true)
                    {
                        if (using_thread == true)
                        {

                            // We Change the TODO_BY column (adding a prefix THREAD_) to avoid the creation of a new TREAD with the same ACTION ID (if the previous tread is not completed before the next tick of the timer)
                            cmd2.CommandText = "UPDATE [IMCA_BACKOFFICE].[dbo].PCM_TAB_IMCA_ACTION set TODO_BY='THREAD_'+TODO_BY WHERE ID=" + dr["ID"].ToString();
                            cmd2.ExecuteNonQuery();

                            WriteToFile("       ACTION : " + dr["ACTION"].ToString() + " (ID : " + dr["ID"].ToString() + ") has been reserved. We can start it (using a THREAD) at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + " - Please check the associated log file of the task for details");
                            WriteToFile("       ACTION : " + dr["ACTION"].ToString() + " (ID : " + dr["ID"].ToString() + ") has been reserved. We can start it at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), long.Parse(dr["ID"].ToString()));
                        }
                        else
                        {
                            WriteToFile("       ACTION : " + dr["ACTION"].ToString() + " (ID : " + dr["ID"].ToString() + ") has been reserved. We can start it at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                        }

                        cmd2.CommandText = "UPDATE [IMCA_BACKOFFICE].[dbo].PCM_TAB_IMCA_ACTION set STARTED=getdate() WHERE ID=" + dr["ID"].ToString();
                        cmd2.ExecuteNonQuery();

                        // Get all the functions (NameSpace / ClassName / Method) for this action
                        cmd2.CommandText = "SELECT [ID], [FUNCTION] as [ClassName_Method],FUNCTION_ORDER,[SK_VALID] FROM [IMCA_BACKOFFICE].[dbo].PCM_TAB_IMCA_ACTION_FUNCTION where SK_VALID<>99 and ACTION='" + dr["ACTION"].ToString() + "' ORDER BY FUNCTION_ORDER";
                        dt2 = cmd2.ExecuteReader();

                        bool at_least_one_error = false;

                        while (dt2.Read())
                        {
                            string ret = "";
                            string func_to_execute = "";

                            switch (dt2["ClassName_Method"].ToString().ToUpper().Trim())
                            {
                                case string x when x.StartsWith("SENDMAIL"):

                                    WriteToFile("               FUNCTION to execute : SENDMAIL", id);

                                    func_to_execute = dt2["ClassName_Method"].ToString().Trim();
                                    func_to_execute = resolve_parameters(func_to_execute, long.Parse(dr["ID"].ToString()), long.Parse(dt2["ID"].ToString()), dr);

                                    ret = Execute_SendMail(func_to_execute, long.Parse(dr["ID"].ToString()), long.Parse(dt2["ID"].ToString()), Int16.Parse(dt2["SK_VALID"].ToString()));

                                    if (ret != "")
                                    {
                                        WriteToFile("       ACTION : " + dr["ACTION"].ToString() + " - Unable to execute the SendMail Function : " + func_to_execute.Trim(), id);
                                        at_least_one_error = true;
                                    }

                                    break;

                                case string x when x.StartsWith("EVALOPENQUERY"):

                                    WriteToFile("               FUNCTION to execute : EVALOPENQUERY (ID : " + dt2["ID"].ToString() + ")", id);

                                    func_to_execute = dt2["ClassName_Method"].ToString().ToUpper().Trim();
                                    func_to_execute = resolve_parameters(func_to_execute, long.Parse(dr["ID"].ToString()), long.Parse(dt2["ID"].ToString()), dr);

                                    ret = Execute_Query(func_to_execute, long.Parse(dr["ID"].ToString()), long.Parse(dt2["ID"].ToString()));

                                    if (ret != "")
                                    {
                                        WriteToFile("       ACTION : " + dr["ACTION"].ToString() + " - Unable to execute the query : " + func_to_execute.Trim(), id);
                                        at_least_one_error = true;
                                    }

                                    break;

                                default:

                                    int count = dt2["ClassName_Method"].ToString().Split('.').Length - 1; // Count the number of point (.)

                                    if (count == 2)
                                    {
                                        WriteToFile("               FUNCTION to execute : " + dt2["ClassName_Method"].ToString() + " (ID : " + dt2["ID"].ToString() + ")", id);
                                        WriteToFile("               NameSpace           : " + dt2["ClassName_Method"].ToString().Split('.')[0], id);
                                        WriteToFile("               ClassName           : " + dt2["ClassName_Method"].ToString().Split('.')[1], id);
                                        WriteToFile("               Method              : " + dt2["ClassName_Method"].ToString().Split('.')[2], id);

                                        try
                                        {
                                            caller(dt2["ClassName_Method"].ToString(), new object[] { sql_con, logs, temp_folder, session_name });

                                        }
                                        catch (Exception ex)
                                        {
                                            WriteToFile("Unable to execute the function " + dt2["ClassName_Method"].ToString() + " - ACTION :  " + dr["ACTION"].ToString() + " - " + ex.Message, id);

                                            at_least_one_error = true;

                                            using (SqlConnection con_sql = new SqlConnection(sql_con))
                                            {
                                                con_sql.Open();

                                                using (SqlCommand cmd_error = new SqlCommand())
                                                {
                                                    cmd_error.Connection = con_sql;
                                                    cmd_error.CommandTimeout = 0;
                                                    cmd_error.CommandText = "update [IMCA_BACKOFFICE].[dbo].PCM_TAB_IMCA_ACTION set TODO_BY=@TODO_BY,ERROR=@ERROR_NUMBER,ERROR_TEXT=@ERROR_MSG where ID=@ID";
                                                    cmd_error.Parameters.Clear();
                                                    cmd_error.Parameters.AddWithValue("@TODO_BY", "ERROR_" + session_name.ToUpper());
                                                    cmd_error.Parameters.AddWithValue("@ID", dr["ID"].ToString());
                                                    cmd_error.Parameters.AddWithValue("@ERROR_NUMBER", 0);
                                                    cmd_error.Parameters.AddWithValue("@ERROR_MSG", Microsoft.VisualBasic.Strings.Left("FUNCTION_ID=" + dt2["ID"].ToString() + " - ERROR MSG=" + ex.Message, 1024));
                                                    cmd_error.ExecuteNonQuery();
                                                }

                                                con_sql.Close();
                                            }
                                        }
                                    }
                                    else
                                    {
                                        WriteToFile("   [FUNCTION] value (" + dt2["ClassName_Method"].ToString() + ") doesn't respect the following syntax : NameSpace.ClassName.MethodName", id);

                                        at_least_one_error = true;
                                        using (SqlConnection con_sql = new SqlConnection(sql_con))
                                        {
                                            con_sql.Open();

                                            using (SqlCommand cmd_error = new SqlCommand())
                                            {
                                                cmd_error.Connection = con_sql;
                                                cmd_error.CommandTimeout = 0;
                                                cmd_error.CommandText = "update [IMCA_BACKOFFICE].[dbo].PCM_TAB_IMCA_ACTION set TODO_BY=@TODO_BY,ERROR=@ERROR_NUMBER,ERROR_TEXT=@ERROR_MSG where ID=@ID";
                                                cmd_error.Parameters.Clear();
                                                cmd_error.Parameters.AddWithValue("@TODO_BY", "ERROR_" + session_name.ToUpper());
                                                cmd_error.Parameters.AddWithValue("@ID", dr["ID"].ToString());
                                                cmd_error.Parameters.AddWithValue("@ERROR_NUMBER", 0);
                                                cmd_error.Parameters.AddWithValue("@ERROR_MSG", Microsoft.VisualBasic.Strings.Left("FUNCTION_ID=" + dt2["ID"].ToString() + " - ERROR MSG=[FUNCTION] field doesn't respect the following syntax : NameSpace.ClassName.MethodName", 1024));
                                                cmd_error.ExecuteNonQuery();
                                            }

                                            con_sql.Close();
                                        }
                                    }

                                    break;
                            }

                            if (at_least_one_error == true)
                            {
                                break;
                            }

                        }
                        dt2.Close();

                        if (at_least_one_error == false)
                        {
                            cmd2.CommandText = "UPDATE [IMCA_BACKOFFICE].[dbo].PCM_TAB_IMCA_ACTION set TODO_BY=replace(TODO_BY,'THREAD_',''),FINISHED=getdate(),SK_FINISH_DATE=DATEDIFF(d,'19951229',getdate()) where ID=" + dr["ID"].ToString();
                            cmd2.ExecuteNonQuery();

                            WriteToFile("       ACTION : " + dr["ACTION"].ToString() + " (ID : " + dr["ID"].ToString() + ") finished at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), id);

                            // 
                            /// We add a new record for the next RUN
                            //

                            cmd2.CommandText = "select getdate()";

                            DateTime dteSysSQL = (DateTime)cmd2.ExecuteScalar();
                            dteSysSQL = new DateTime(dteSysSQL.Year, dteSysSQL.Month, dteSysSQL.Day, dteSysSQL.Hour, dteSysSQL.Minute, 0, 0);
                            DateTime dteNextToDodate = dteSysSQL;

                            int DateInterval_value = 0;

                            if (dr["NEXT_RUN_DELAY_TYPE"].ToString() != "-" && dr["CREATED_FROM_AUTO"].ToString() != "-1")
                            {
                                switch (dr["NEXT_RUN_DELAY_TYPE"].ToString())
                                {
                                    case "N":
                                        {
                                            dteNextToDodate = DateAndTime.DateAdd(DateInterval.Minute, Convert.ToInt32(dr["NEXT_RUN_DELAY"].ToString()), dteNextToDodate);
                                            DateInterval_value = (int)DateInterval.Minute;
                                            break;
                                        }
                                    case "S":
                                        {
                                            dteNextToDodate = DateAndTime.DateAdd(DateInterval.Second, Convert.ToInt32(dr["NEXT_RUN_DELAY"].ToString()), dteNextToDodate);
                                            DateInterval_value = (int)DateInterval.Second;
                                            break;
                                        }

                                    case "YYYY":
                                    case "M":
                                    case "D":
                                    case "H":
                                        {
                                            switch (dr["NEXT_RUN_DELAY_TYPE"].ToString())
                                            {
                                                case "YYYY":
                                                    DateInterval_value = (int)DateInterval.Year;
                                                    break;
                                                case "M":
                                                    DateInterval_value = (int)DateInterval.Month;
                                                    break;
                                                case "D":
                                                    DateInterval_value = (int)DateInterval.Day;
                                                    break;
                                                case "H":
                                                    DateInterval_value = (int)DateInterval.Hour;
                                                    break;
                                            }

                                            dteNextToDodate = DateAndTime.DateAdd((DateInterval)DateInterval_value, Convert.ToInt32(dr["NEXT_RUN_DELAY"].ToString()), (DateTime)(dr["TODO_DATE"]));

                                            //while (DateAndTime.DateDiff((DateInterval)DateInterval_value, dteSysSQL, dteNextToDodate) < 1)
                                            //{
                                            //    dteNextToDodate = DateAndTime.DateAdd((DateInterval)DateInterval_value, Convert.ToInt32(dr["NEXT_RUN_DELAY"].ToString()), (DateTime)dteNextToDodate);
                                            //}

                                            // Check if the new date is in the future
                                            while (DateTime.Compare(dteNextToDodate, dteSysSQL) <= 0)
                                            {
                                                dteNextToDodate = DateAndTime.DateAdd((DateInterval)DateInterval_value, Convert.ToInt32(dr["NEXT_RUN_DELAY"].ToString()), (DateTime)dteNextToDodate);
                                            }

                                            break;
                                        }

                                    case "FM":
                                        {
                                            dteNextToDodate = (DateTime)(dr["TODO_DATE"]);
                                            Int32 lngl;
                                            int lngFMDays;

                                            while (DateAndTime.DateDiff(DateInterval.Day, dteSysSQL, dteNextToDodate) < 1)
                                            {
                                                lngl = 0;
                                                do
                                                {
                                                    cmd2.CommandText = "SELECT COUNT(*) FROM [IMCA_BACKOFFICE].[dbo].DSSIMPRT_TAB_SK_TIME where TIMETYPE='CD' and SKT_FM=" + CurrentSKTime("FM", dteNextToDodate.ToString("yyyyMMdd"));
                                                    lngFMDays = (int)cmd2.ExecuteScalar();

                                                    dteNextToDodate = DateAndTime.DateAdd(DateInterval.Day, lngFMDays, dteNextToDodate);

                                                    lngl = lngl + 1;

                                                } while (lngl < (int)dr["NEXT_RUN_DELAY"]);

                                            }

                                            break;
                                        }

                                }

                                //if (dr["NEXT_RUN_DELAY_TYPE"].ToString() != "FM")
                                //{
                                //    while (TodoDateInWorkTime(dteNextToDodate, dteSysSQL, dr["START_TIME"].ToString(), dr["END_TIME"].ToString()) == false)
                                //    {
                                //        dteNextToDodate = DateAndTime.DateAdd((DateInterval)DateInterval_value, Convert.ToInt32(dr["NEXT_RUN_DELAY"].ToString()), (DateTime)dteNextToDodate);

                                //    }
                                //}

                                if (dr["CREATED_FROM"].ToString() == "SYSTEM")
                                {

                                    // Get the details of the ACTION                                        

                                    cmd2.CommandText = "SELECT * FROM [IMCA_BACKOFFICE].[dbo].PCM_TAB_IMCA_ACTION where ID=" + dr["ID"].ToString();
                                    dt2 = cmd2.ExecuteReader();

                                    using (SqlConnection con3 = new SqlConnection(sql_con))
                                    {
                                        con3.Open();

                                        using (SqlCommand cmd3 = new SqlCommand())
                                        {
                                            cmd3.Connection = con3;
                                            cmd3.CommandTimeout = 0;

                                            while (dt2.Read())
                                            {
                                                cmd3.CommandText = "SELECT COUNT(*) FROM [IMCA_BACKOFFICE].[dbo].PCM_TAB_IMCA_ACTION where CREATED_FROM_AUTO=" + dr["ID"].ToString() + " AND [ACTION]='" + dt2["ACTION"] + "'";
                                                do
                                                {
                                                    // Add a new ACTION
                                                    AddIMCAAction(dteSysSQL, long.Parse(dt2["SK_VALID"].ToString()), dt2["ACTION"].ToString(), -9999, dteNextToDodate, dt2["created_from"].ToString(), dr["NEXT_RUN_TODO_BY"].ToString(), dt2["SK_Type"].ToString(), dt2["param01"].ToString(), dt2["Param02"].ToString(), dt2["PARAM03"].ToString(), dt2["PARAM04"].ToString(), dt2["PARAM05"].ToString(), dt2["PARAM06"].ToString(), dt2["PARAM07"].ToString(), dt2["PARAM08"].ToString(), dt2["param09"].ToString(), (string)dt2["param10"], dt2["PARAM11"].ToString(), dt2["PARAM12"].ToString(), dt2["PARAM13"].ToString(), dt2["PARAM14"].ToString(), dt2["PARAM15"].ToString(), dt2["PARAM16"].ToString(), dt2["PARAM17"].ToString(), dt2["PARAM18"].ToString(), dt2["PARAM19"].ToString(), dt2["PARAM20"].ToString(), dt2["PARAM_MEMO"].ToString(), dt2["ID"].ToString(), long.Parse(dt2["sk_value"].ToString()));
                                                } while ((int)cmd3.ExecuteScalar() < 0);
                                            }

                                            cmd3.Dispose();
                                        }

                                    }
                                    dt2.Close();
                                }
                            }
                        }
                    }
                    else
                    {
                        WriteToFile("       ACTION : " + dr["ACTION"].ToString() + " (ID : " + dr["ID"].ToString() + ") has not been reserved (Other(s) Component(s) are not available)  at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));

                    }
                }

                con.Close();
            }

            if (using_thread == true)
                nb_created_thread -= 1;
        }

        protected void check_if_TODO(string logs, string temp_folder)
        {
            timer_check_TODO_is_running = true;

            using (SqlConnection con = new SqlConnection(sql_con))
            {
                con.Open();

                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.Connection = con;
                    cmd.CommandTimeout = 0;

                    Boolean bol = true;
                    do
                    {
                        // Check if a action has been reserved (with the session name) and not started/finished
                        cmd.CommandText = " SET LANGUAGE FRENCH;SELECT action.ID,admin.ACTION,admin.ID as [ADMIN_ID],admin.SK_VALID," +
                                          " CASE WHEN isnull(admin.NEXT_RUN_DELAY_TYPE,'')='Y' THEN 'YYYY' ELSE isnull(admin.NEXT_RUN_DELAY_TYPE,'') END as NEXT_RUN_DELAY_TYPE,NEXT_RUN_DELAY, " +
                                          " action.[CREATED_FROM_AUTO],action.[TODO_DATE], CASE WHEN isnull(admin.NEXT_RUN_TODO_BY,'')='' THEN 'TODO' ELSE isnull(admin.NEXT_RUN_TODO_BY,'') END as NEXT_RUN_TODO_BY, " +
                                          " upper(action.CREATED_FROM) as CREATED_FROM, " +
                                          " convert(varchar(5),START_TIME,108) as START_TIME,convert(varchar(5),END_TIME,108) as END_TIME, " +
                                          " PARAM01,PARAM02,PARAM03,PARAM04,PARAM05,PARAM06,PARAM07,PARAM08,PARAM09,PARAM10,PARAM11,PARAM12,PARAM13,PARAM14,PARAM15,PARAM16,PARAM17,PARAM18,PARAM19,PARAM20 " +
                                          " FROM [IMCA_BACKOFFICE].[dbo].[PCM_TAB_IMCA_ACTION_ADMIN] admin " +
                                          " INNER JOIN [IMCA_BACKOFFICE].[dbo].[PCM_TAB_IMCA_ACTION] action on admin.ACTION=action.ACTION " +
                                          " WHERE admin.ACTION in (SELECT ACTION FROM [IMCA_BACKOFFICE].[dbo].PCM_TAB_IMCA_ACTION_USERVALIDATION WHERE USERID LIKE 'SERVICE%' AND VALID = 1) AND TODO_BY='" + session_name.ToUpper() + "'" +
                                          " AND action.SK_FINISH_DATE=0 AND TODO_DATE_INT<cast(replace(convert(varchar(10),GETDATE(),102),'.','')+replace(convert(varchar(10),GETDATE(),108),':','') as bigint) " +
                                          " ORDER BY action.PRIORITY, action.ID ";
                        SqlDataReader dt;

                        dt = cmd.ExecuteReader();
                        DataTable row = new DataTable();
                        row.Load(dt);
                        dt.Close();

                        foreach (DataRow dr in row.Rows)
                        {
                            if (EXECUTE_IMCA_ACTION_USING_THREAD.Trim().ToUpper() == "TRUE")
                            {
                                if (nb_created_thread <= Int16.Parse(NUMBER_OF_THREAD_MAX))
                                {
                                    Thread new_IMCA_action_Thread = new Thread(() => execute_IMCA_Action(dr, logs, temp_folder, true));
                                    new_IMCA_action_Thread.IsBackground = true;
                                    new_IMCA_action_Thread.SetApartmentState(ApartmentState.STA);
                                    new_IMCA_action_Thread.Start();
                                }
                            }
                            else
                            {
                                execute_IMCA_Action(dr, logs, temp_folder, false);
                            }
                            System.Threading.Thread.Sleep(500);
                        }


                        // Check if a TODO action must be reserved (SELECT TOP 1)
                        cmd.CommandText = "SET LANGUAGE FRENCH;SELECT TOP 1 action.ID,admin.ACTION,admin.RUN_ON_1,admin.RUN_ON_2,admin.RUN_ON_3,admin.RUN_ON_4,admin.RUN_ON_5,admin.RUN_ON_6,admin.RUN_ON_7,isnull(list_valid_user,'') as  VALID_USERID " +
                                          " FROM [IMCA_BACKOFFICE].[dbo].[PCM_TAB_IMCA_ACTION_ADMIN] admin " +
                                          " INNER JOIN [IMCA_BACKOFFICE].[dbo].[PCM_TAB_IMCA_ACTION] action on admin.ACTION=action.ACTION " +
                                          " LEFT JOIN(SELECT distinct ACTION, STUFF((SELECT  distinct tn2.[USERID] + '#' FROM   PCM_TAB_IMCA_ACTION_USERVALIDATION tn2 inner join  PCM_TAB_IMCA_ACTION_USERVALIDATION " +
                                          " on tn2.ACTION = PCM_TAB_IMCA_ACTION_USERVALIDATION.ACTION  WHERE tn1.ACTION = tn2.ACTION  FOR XML PATH('')), 1, 0, '') as list_valid_user  from PCM_TAB_IMCA_ACTION_USERVALIDATION tn1 " +
                                          " ) PCM_TAB_IMCA_ACTION_USERVALIDATION on PCM_TAB_IMCA_ACTION_USERVALIDATION.ACTION = action.ACTION " +
                                          " WHERE admin.ACTION in (SELECT ACTION FROM [IMCA_BACKOFFICE].[dbo].PCM_TAB_IMCA_ACTION_USERVALIDATION WHERE USERID LIKE 'SERVICE%' AND VALID = 1) AND TODO_BY='TODO' and action.SK_FINISH_DATE=0 " +
                                          " AND TODO_DATE_INT<cast(replace(convert(varchar(10),GETDATE(),102),'.','')+replace(convert(varchar(10),GETDATE(),108),':','') as bigint) " +
                                          " AND cast(replace(convert(varchar(5), GETDATE(), 108), ':', '') as int) BETWEEN cast(replace(CONVERT(varchar(5), admin.START_TIME, 108),':','') as int)  and cast(replace(CONVERT(varchar(5), admin.END_TIME, 108),':','') as int) " +
                                          " ORDER BY action.PRIORITY, action.ID ";

                        dt = cmd.ExecuteReader();


                        if (dt.HasRows)
                        {
                            while (dt.Read())
                            {
                                Boolean to_execute = false;
                                Boolean valid_user = true;

                                using (SqlCommand cmd2 = new SqlCommand())
                                {

                                    switch (DateAndTime.Now.DayOfWeek)
                                    {
                                        case DayOfWeek.Monday:
                                            if (int.Parse(dt["RUN_ON_1"].ToString()) == 1)
                                                to_execute = true;
                                            break;

                                        case DayOfWeek.Tuesday:
                                            if (int.Parse(dt["RUN_ON_2"].ToString()) == 1)
                                                to_execute = true;
                                            break;

                                        case DayOfWeek.Wednesday:
                                            if (int.Parse(dt["RUN_ON_3"].ToString()) == 1)
                                                to_execute = true;
                                            break;

                                        case DayOfWeek.Thursday:
                                            if (int.Parse(dt["RUN_ON_4"].ToString()) == 1)
                                                to_execute = true;
                                            break;

                                        case DayOfWeek.Friday:
                                            if (int.Parse(dt["RUN_ON_5"].ToString()) == 1)
                                                to_execute = true;
                                            break;

                                        case DayOfWeek.Saturday:
                                            if (int.Parse(dt["RUN_ON_6"].ToString()) == 1)
                                                to_execute = true;
                                            break;

                                        case DayOfWeek.Sunday:
                                            if (int.Parse(dt["RUN_ON_7"].ToString()) == 1)
                                                to_execute = true;
                                            break;

                                    }


                                    //
                                    // Check the VALID User
                                    // 

                                    if (dt["VALID_USERID"].ToString().Contains("#") == true)
                                    {
                                        string[] list_valid_users = dt["VALID_USERID"].ToString().Split('#');

                                        foreach (String str in list_valid_users)
                                        {
                                            if (str.Trim() != "")
                                            {
                                                if (str.Contains("%") == false)
                                                {
                                                    if (str.ToUpper().Trim() != session_name.ToUpper().Trim())
                                                    {
                                                        to_execute = false;
                                                        valid_user = false;
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                    }


                                    if (to_execute == true)
                                    {
                                        cmd2.Connection = con;
                                        cmd2.CommandTimeout = 0;
                                        cmd2.CommandText = "UPDATE [IMCA_BACKOFFICE].[dbo].[PCM_TAB_IMCA_ACTION] set TODO_BY='" + session_name.ToUpper() + "' where ID=" + dt["ID"].ToString() + " AND TODO_BY='TODO'";
                                        cmd2.ExecuteNonQuery();

                                        WriteToFile("       " + session_name.ToUpper() + " reserved ACTION  " + dt["ACTION"].ToString() + " (ID : " + dt["ID"].ToString() + ") at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                                    }
                                    else
                                    {
                                        bol = false;

                                        if (valid_user == true)
                                        {
                                            WriteToFile("       ACTION  " + dt["ACTION"].ToString() + " (ID : " + dt["ID"].ToString() + ") not allowed to start on " + DateAndTime.Now.DayOfWeek.ToString() + " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                                        }

                                    }

                                }
                            }
                            dt.Close();
                        }
                        else
                        {
                            dt.Close();
                            bol = false; // We exit the loop and 
                        }

                    } while (bol);

                    // Update SURVEYER status. Call the Stored Procedure

                    cmd.CommandText = "SURVEYER.dbo.USP_IMCA_SURVEYER_UPDATE";
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@CompanyCd", "EMEA");
                    cmd.Parameters.AddWithValue("@Alias", session_name.ToUpper());
                    cmd.Parameters.AddWithValue("@Description", "");
                    int rowAffected = cmd.ExecuteNonQuery();


                }
                con.Close();
            }

            timer_check_TODO_is_running = false;
        }
        public void WriteToFile(string message, long id = 0, string logs_folder = "logs")
        {
            logs_folder = service_path + "\\" + logs_folder;

            if (!Directory.Exists(logs_folder))
            {
                Directory.CreateDirectory(logs_folder);
            }

            string filepath;

            switch (id)
            {
                case 0:
                    filepath = logs_folder + "\\IMCA_" + session_name.ToUpper() + "_" + DateTime.Now.Date.ToString("dd_MM_yyyy") + ".txt";
                    break;

                default:
                    filepath = logs_folder + "\\IMCA_" + session_name.ToUpper() + "_" + DateTime.Now.Date.ToString("dd_MM_yyyy") + "_ACTION_ID_" + id.ToString() + ".txt";
                    break;
            }



            if (!File.Exists(filepath))
            {
                using (StreamWriter fic = File.CreateText(filepath))
                {
                    fic.WriteLine(message);
                }
            }
            else
            {
                using (StreamWriter fic = File.AppendText(filepath))
                {
                    fic.WriteLine(message);
                }
            }
        }
    }

}
