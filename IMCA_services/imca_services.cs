using Microsoft.VisualBasic;
using Newtonsoft.Json;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.ServiceProcess;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;



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
        // Number of action threads currently running in this service instance.
        // Access this field only through Interlocked/Volatile operations.
        private int nb_created_thread = 0;
        private Boolean timer_check_TODO_is_running = false;
        // Set to 1 while Windows is stopping the service.
        // The timer callback checks this value before starting a new cycle.
        private int service_is_stopping = 0;
        private string logs_folder = "";
        private string temp_folder = "";

        // Synchronizes log file writes performed by concurrent action threads.
        private static readonly object logLock = new object();

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
            Interlocked.Exchange(ref service_is_stopping, 0);
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
                try
                {
                    SqlConnectionStringBuilder safeConnection = new SqlConnectionStringBuilder(sql_con);
                    WriteToFile(
                        "   SQL Server                              : " + safeConnection.DataSource +
                        " / Database : " + safeConnection.InitialCatalog,
                        0, logs_folder);
                }
                catch
                {
                    WriteToFile("   SQL connection configured               : TRUE", 0, logs_folder);
                }
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
        //protected override void OnStop()
        //{
        //    while (timer_check_TODO_is_running == true)
        //    {
        //        System.Threading.Thread.Sleep(500);
        //        System.Windows.Forms.Application.DoEvents();
        //    }

        //    if (timer_check_TODO != null)
        //    {
        //        timer_check_TODO.Stop();
        //        timer_check_TODO.Dispose();

        //        RequestAdditionalTime(120000);

        //        timer_check_TODO?.Stop();

        //        while (timer_check_TODO_is_running ||
        //               Volatile.Read(ref nb_created_thread) > 0)
        //        {
        //            RequestAdditionalTime(120000);
        //            Thread.Sleep(500);
        //        }

        //        timer_check_TODO?.Dispose();


        //    }

        //    if (Volatile.Read(ref nb_created_thread) > 0 || HasDatabaseThreadedActions())
        //    {
        //        WriteToFile("Before stopping the service we check if threads are running at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));

        //        Boolean bol_thread = true;
        //        int nb_running_thread = 0;

        //        using (SqlConnection con = new SqlConnection(sql_con))
        //        {
        //            con.Open();

        //            using (SqlCommand cmd = new SqlCommand())
        //            {
        //                cmd.Connection = con;
        //                cmd.CommandTimeout = 300;

        //                while (bol_thread)
        //                {
        //                    cmd.CommandText = "select count(*) from [IMCA_BACKOFFICE].[dbo].PCM_TAB_IMCA_ACTION where TODO_BY='THREAD_" + session_name.ToUpper() + "'";
        //                    nb_running_thread = int.Parse(cmd.ExecuteScalar().ToString());

        //                    if (nb_running_thread > 0)
        //                    {
        //                        WriteToFile(nb_running_thread.ToString() + " Thread(s) is(are) running. We wait 10 secs at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
        //                        System.Threading.Thread.Sleep(10000);
        //                        System.Windows.Forms.Application.DoEvents();
        //                    }
        //                    else
        //                    {
        //                        bol_thread = false;
        //                    }
        //                }
        //            }

        //            con.Close();
        //        }

        //    }

        //    WriteToFile("IMCA Services stopped at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
        //}

        protected override void OnStop()
        {
            // Prevent any new timer cycle from being started during shutdown.
            Interlocked.Exchange(ref service_is_stopping, 1);

            WriteToFile(
                "IMCA Services stopping at " +
                DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));

            // Prevent the timer from starting another polling cycle.
            timer_check_TODO?.Stop();

            DateTime stopDeadline = DateTime.Now.AddMinutes(10);

            // Wait for the timer callback and local action threads.
            while ((timer_check_TODO_is_running ||
                    Volatile.Read(ref nb_created_thread) > 0) &&
                   DateTime.Now < stopDeadline)
            {
                // Inform Windows that the stop operation is still progressing.
                RequestAdditionalTime(30000);

                WriteToFile(
                    "Waiting for service tasks to stop. Timer running: " +
                    timer_check_TODO_is_running +
                    ", active threads: " +
                    Volatile.Read(ref nb_created_thread));

                Thread.Sleep(1000);
            }

            timer_check_TODO?.Dispose();
            timer_check_TODO = null;

            // Do not wait forever if an application or macro is blocked.
            if (timer_check_TODO_is_running ||
                Volatile.Read(ref nb_created_thread) > 0 ||
                HasDatabaseThreadedActions())
            {
                WriteToFile(
                    "Service stop timeout reached. Some actions are still running: " + timer_check_TODO_is_running + " , active threads: " + Volatile.Read(ref nb_created_thread));
            }

            WriteToFile(
                "IMCA Services stopped at " +
                DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
        }


        public void start_timer(string logs, string temp_folder)
        {
            timer_check_TODO = new System.Timers.Timer();
            timer_check_TODO.Elapsed += delegate { OnElaspedTime(logs, temp_folder); };
            timer_check_TODO.Interval = nb_sec * 1000;
            timer_check_TODO.AutoReset = false;
            timer_check_TODO.Enabled = true;
        }
        private void OnElaspedTime(string logs, string temp_folder)
        {
            // Ignore timer ticks received while the service is stopping.
            if (Volatile.Read(ref service_is_stopping) == 1)
            {
                return;
            }

            WriteToFile("   Start Checking TODO was ran at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
            timer_check_TODO?.Stop();

            try
            {
                check_if_TODO(logs, temp_folder);
            }
            catch (Exception ex)
            {
                WriteToFile("Exception while checking TODO : " + ex + " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
            }
            finally
            {
                // This flag must always be released, including when SQL Server is unavailable.
                Volatile.Write(ref timer_check_TODO_is_running, false);

                // Restart polling only when Windows has not requested service shutdown.
                if (Volatile.Read(ref service_is_stopping) == 0)
                {
                    timer_check_TODO?.Start();
                }

                WriteToFile("   End Checking TODO was ran at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
            }
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
            List<string> allowedTypes = new List<string>() { "CD", "FD", "FW", "CM", "FM", "CQ", "FQ", "CY", "FY" };
            if (!allowedTypes.Contains(strType)) strType = "CD";
            if (string.IsNullOrWhiteSpace(strISODate) || !strISODate.All(char.IsNumber)) strISODate = DateTime.Now.ToString("yyyyMMdd");
            string sql = "SELECT SKT_" + strType.ToUpperInvariant() + " FROM IMCA_BACKOFFICE.dbo.DSSIMPRT_TAB_SK_TIME " +
                         "WHERE DATESTART=CONVERT(DATETIME,@ISO_DATE,112) AND TIMETYPE='CD'";
            try
            {
                using (SqlConnection connection = new SqlConnection(sql_con))
                using (SqlCommand command = new SqlCommand(sql, connection))
                {
                    command.CommandTimeout = 300;
                    command.Parameters.Add("@ISO_DATE", SqlDbType.VarChar, 8).Value = strISODate;
                    connection.Open();
                    object result = command.ExecuteScalar();
                    long value = result == null || result == DBNull.Value ? 0 : Convert.ToInt64(result);
                    return value == 0 ? -1 : value;
                }
            }
            catch (Exception ex)
            {
                LogDatabaseError(nameof(CurrentSKTime), sql, ex);
                throw;
            }
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
                            cmd.CommandTimeout = 300;
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
                LogDatabaseError(nameof(AddIMCAAction), "Create next IMCA action", ex);
                throw;
            }


        }

        private string get_IMCA_paramters(string sqlCon, string paramName)
        {
            const string sql = @"SELECT ISNULL(VALUE, '')
FROM [PCM_TAB_IMCA_PARAMETER_GLOBAL]
WHERE SK_VALID=0 AND PARAMETER=@PARAMETER";
            try
            {
                using (SqlConnection connection = new SqlConnection(sqlCon))
                using (SqlCommand command = new SqlCommand(sql, connection))
                {
                    command.CommandTimeout = 300;
                    command.Parameters.Add("@PARAMETER", SqlDbType.VarChar, 255).Value = paramName ?? "";
                    connection.Open();
                    object result = command.ExecuteScalar();
                    return result == null || result == DBNull.Value ? "" : Convert.ToString(result);
                }
            }
            catch (Exception ex)
            {
                LogDatabaseError(nameof(get_IMCA_paramters), sql, ex);
                throw;
            }
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
                                cmd.CommandTimeout = 300;
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
                            cmd.CommandTimeout = 300;

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
                    LogDatabaseError(nameof(Execute_SendMail), "Insert email into PCM_TAB_IMCA_AUTOMAIL", ex, ACTION_ID);
                    TryMarkActionAsError(ACTION_ID, FUNCTION_ID, ex);
                    ret = " Unable to execute the function SendMail (error : " + ex.Message + ")";
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
            int maxErrors;
            if (!int.TryParse(Number_of_errors_allowed_per_query, out maxErrors) || maxErrors < 1)
            {
                maxErrors = 1;
            }

            List<string> list_args = new List<string>();
            RegexOptions options = RegexOptions.Multiline;
            string pattern = @"([\w.$]+|""[^""]+""|'[^']+')";

            while (strSQL.IndexOf(",,") != -1)
            {
                strSQL = strSQL.Replace(",,", ",\"OPTIONAL_VALUE_NOT_PROVIDED\",");
            }

            foreach (Match match in Regex.Matches(strSQL, pattern, options))
            {
                list_args.Add(match.Value);
            }

            if (list_args.Count < 3)
            {
                return " Too few parameters for EvalOpenQuery function - " + strSQL;
            }

            // The first argument contains the function name.
            list_args.RemoveAt(0);

            string queryText = list_args[0].ToString().Trim('"');
            string connectionParameter = list_args[1].ToString()
                .Replace(Microsoft.VisualBasic.Strings.Chr(34), ' ')
                .Trim();

            try
            {
                string connectionString = get_IMCA_paramters(sql_con, connectionParameter);

                if (string.IsNullOrWhiteSpace(connectionString))
                {
                    return "Unable to find (Check the Global Parameters Table) the connection String for " +
                           connectionParameter;
                }

                int commandTimeout = 120;
                if (list_args.Count >= 7)
                {
                    int.TryParse(list_args[6].ToString(), out commandTimeout);
                    if (commandTimeout <= 0)
                    {
                        commandTimeout = 120;
                    }
                }

                if (list_args.Count >= 8)
                {
                    int configuredMaxErrors;
                    if (int.TryParse(list_args[7].ToString(), out configuredMaxErrors) && configuredMaxErrors > 0)
                    {
                        maxErrors = configuredMaxErrors;
                    }
                }

                bool isOracle = Strings.Replace(connectionString.ToUpperInvariant(), " ", "")
                    .StartsWith("DATASOURCE=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST");

                if (isOracle)
                {
                    // Oracle-specific objects are created only after the connection type is identified.
                    OracleConnectionStringBuilder builder = new OracleConnectionStringBuilder(connectionString);

                    using (OracleConnection connection = new OracleConnection(connectionString))
                    {
                        connection.Open();

                        for (int attempt = 1; attempt <= maxErrors; attempt++)
                        {
                            using (OracleTransaction transaction = connection.BeginTransaction())
                            using (OracleCommand command = new OracleCommand(queryText, connection))
                            {
                                command.CommandType = CommandType.Text;
                                command.CommandTimeout = commandTimeout;
                                command.Transaction = transaction;

                                try
                                {
                                    command.ExecuteNonQuery();
                                    transaction.Commit();
                                    return "";
                                }
                                catch (OracleException ex)
                                {
                                    try
                                    {
                                        transaction.Rollback();
                                    }
                                    catch
                                    {
                                        // The connection may already be unavailable.
                                    }

                                    WriteToFile(
                                        "                                   : ORACLE ERROR : " + ex.Number +
                                        " - attempt " + attempt + " of " + maxErrors +
                                        " - " + ex.Message);

                                    if (attempt >= maxErrors)
                                    {
                                        LogDatabaseError(nameof(Execute_Query), queryText, ex, ACTION_ID);
                                        TryMarkActionAsError(ACTION_ID, FUNCTION_ID, ex);
                                        return " Oracle Query error (" + ex.Number + " - " + ex.Message +
                                               ") - function - " + strSQL;
                                    }

                                    if (connection.State == ConnectionState.Closed)
                                    {
                                        connection.OpenWithNewPassword(builder.Password);
                                    }

                                    Thread.Sleep(60000);
                                }
                            }
                        }
                    }
                }
                else
                {
                    // SQL commands use the dynamically resolved connection string.
                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {
                        connection.Open();

                        for (int attempt = 1; attempt <= maxErrors; attempt++)
                        {
                            using (SqlCommand command = new SqlCommand(queryText, connection))
                            {
                                command.CommandType = CommandType.Text;
                                command.CommandTimeout = commandTimeout;

                                try
                                {
                                    command.ExecuteNonQuery();
                                    return "";
                                }
                                catch (SqlException ex)
                                {
                                    WriteToFile(
                                        "                                   : SQL ERROR : " + ex.Number +
                                        " - attempt " + attempt + " of " + maxErrors +
                                        " - " + ex.Message);

                                    if (attempt >= maxErrors)
                                    {
                                        LogDatabaseError(nameof(Execute_Query), queryText, ex, ACTION_ID);
                                        TryMarkActionAsError(ACTION_ID, FUNCTION_ID, ex);
                                        return " SQL Query error (" + ex.Number + " - " + ex.Message +
                                               ") - function - " + strSQL;
                                    }

                                    Thread.Sleep(60000);
                                }
                            }
                        }
                    }
                }

                return "";
            }
            catch (OracleException ex)
            {
                LogDatabaseError(nameof(Execute_Query), queryText, ex, ACTION_ID);
                TryMarkActionAsError(ACTION_ID, FUNCTION_ID, ex);
                return " Unable to execute the Oracle query (error : " + ex.Message + ") - " + queryText;
            }
            catch (SqlException ex)
            {
                LogDatabaseError(nameof(Execute_Query), queryText, ex, ACTION_ID);
                TryMarkActionAsError(ACTION_ID, FUNCTION_ID, ex);
                return " Unable to execute the SQL query (error : " + ex.Message + ") - " + queryText;
            }
            catch (Exception ex)
            {
                LogDatabaseError(nameof(Execute_Query), queryText, ex, ACTION_ID);
                TryMarkActionAsError(ACTION_ID, FUNCTION_ID, ex);
                return " Unable to execute the query (error : " + ex.Message + ") - " + queryText;
            }
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
            long actionId = 0;
            string actionName = "";
            try
            {
                actionId = Convert.ToInt64(dr["ID"]);
                actionName = Convert.ToString(dr["ACTION"]);

                SqlDataReader dt2;

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
                        cmd2.CommandTimeout = 300;

                        // Check only availability flags. USE_THREAD is a control flag and must never block an action.
                        cmd2.CommandText = "SELECT action_flag.FLAG FROM [PCM_TAB_IMCA_ACTION_FLAG] action_flag " +
                                           "INNER JOIN (SELECT * FROM [PCM_TAB_IMCA_PARAMETER_GLOBAL] " +
                                           "WHERE PARAMETER LIKE 'AVAIL_%' AND [VALUE]='0' " +
                                           "AND (SK_VALID=0 OR SK_VALID=" + dr["SK_VALID"].ToString() + ")) [parameters] " +
                                           "ON action_flag.FLAG=[parameters].PARAMETER " +
                                           "WHERE action_flag.[ACTION_ID]=" + dr["ADMIN_ID"].ToString() + " " +
                                           "AND UPPER(action_flag.FLAG)<>'USE_THREAD'";
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
                                WriteToFile("       ACTION : " + dr["ACTION"].ToString() + " (ID : " + dr["ID"].ToString() + ") has been reserved. We can start it at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), long.Parse(dr["ID"].ToString()), logs_folder, dr["ACTION"].ToString());
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

                                        WriteToFile("               FUNCTION to execute : SENDMAIL", id, logs_folder, dr["ACTION"].ToString());

                                        func_to_execute = dt2["ClassName_Method"].ToString().Trim();
                                        func_to_execute = resolve_parameters(func_to_execute, long.Parse(dr["ID"].ToString()), long.Parse(dt2["ID"].ToString()), dr);

                                        ret = Execute_SendMail(func_to_execute, long.Parse(dr["ID"].ToString()), long.Parse(dt2["ID"].ToString()), Int16.Parse(dt2["SK_VALID"].ToString()));

                                        if (ret != "")
                                        {
                                            WriteToFile("       ACTION : " + dr["ACTION"].ToString() + " - Unable to execute the SendMail Function : " + func_to_execute.Trim(), id, logs_folder, dr["ACTION"].ToString());
                                            at_least_one_error = true;
                                        }

                                        break;

                                    case string x when x.StartsWith("EVALOPENQUERY"):

                                        WriteToFile("               FUNCTION to execute : EVALOPENQUERY (ID : " + dt2["ID"].ToString() + ")", id, logs_folder, dr["ACTION"].ToString());

                                        func_to_execute = dt2["ClassName_Method"].ToString().ToUpper().Trim();
                                        func_to_execute = resolve_parameters(func_to_execute, long.Parse(dr["ID"].ToString()), long.Parse(dt2["ID"].ToString()), dr);

                                        ret = Execute_Query(func_to_execute, long.Parse(dr["ID"].ToString()), long.Parse(dt2["ID"].ToString()));

                                        if (ret != "")
                                        {
                                            WriteToFile("       ACTION : " + dr["ACTION"].ToString() + " - Unable to execute the query : " + func_to_execute.Trim(), id, logs_folder, dr["ACTION"].ToString());
                                            at_least_one_error = true;
                                        }

                                        break;

                                    default:

                                        int count = dt2["ClassName_Method"].ToString().Split('.').Length - 1; // Count the number of point (.)

                                        if (count == 2)
                                        {
                                            WriteToFile("               FUNCTION to execute : " + dt2["ClassName_Method"].ToString() + " (ID : " + dt2["ID"].ToString() + ")", id, logs_folder, dr["ACTION"].ToString());
                                            WriteToFile("               NameSpace           : " + dt2["ClassName_Method"].ToString().Split('.')[0], id, logs_folder, dr["ACTION"].ToString());
                                            WriteToFile("               ClassName           : " + dt2["ClassName_Method"].ToString().Split('.')[1], id, logs_folder, dr["ACTION"].ToString());
                                            WriteToFile("               Method              : " + dt2["ClassName_Method"].ToString().Split('.')[2], id, logs_folder, dr["ACTION"].ToString());

                                            try
                                            {
                                                caller(dt2["ClassName_Method"].ToString(), new object[] { sql_con, logs, temp_folder, session_name });

                                            }
                                            catch (Exception ex)
                                            {
                                                WriteToFile("Unable to execute the function " + dt2["ClassName_Method"].ToString() + " - ACTION :  " + dr["ACTION"].ToString() + " - " + ex.Message, id, logs_folder, dr["ACTION"].ToString());

                                                at_least_one_error = true;

                                                LogDatabaseError(nameof(execute_IMCA_Action), dt2["ClassName_Method"].ToString(), ex, actionId, actionName);


                                                TryMarkActionAsError(actionId, Convert.ToInt64(dt2["ID"]), ex);
                                            }
                                        }
                                        else
                                        {
                                            WriteToFile("   [FUNCTION] value (" + dt2["ClassName_Method"].ToString() + ") doesn't respect the following syntax : NameSpace.ClassName.MethodName", id, logs_folder, dr["ACTION"].ToString());

                                            at_least_one_error = true;
                                            Exception functionFormatException = new FormatException("[FUNCTION] field does not respect NameSpace.ClassName.MethodName syntax");

                                            LogDatabaseError(nameof(execute_IMCA_Action), dt2["ClassName_Method"].ToString(), functionFormatException, actionId, actionName);

                                            TryMarkActionAsError(actionId, Convert.ToInt64(dt2["ID"]), functionFormatException);
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

                                WriteToFile("       ACTION : " + dr["ACTION"].ToString() + " (ID : " + dr["ID"].ToString() + ") finished at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), id, logs_folder, dr["ACTION"].ToString());

                                // 
                                /// We add a new record for the next RUN
                                //

                                cmd2.CommandText = "select getdate()";

                                DateTime dteSysSQL = (DateTime)cmd2.ExecuteScalar();
                                dteSysSQL = new DateTime(dteSysSQL.Year, dteSysSQL.Month, dteSysSQL.Day, dteSysSQL.Hour, dteSysSQL.Minute, dteSysSQL.Second, 0);
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
                                                cmd3.CommandTimeout = 300;

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

            }
            catch (Exception ex)
            {
                LogDatabaseError(nameof(execute_IMCA_Action), "IMCA action execution", ex, actionId, actionName);
                TryMarkActionAsError(actionId, 0, ex);
                throw;
            }
        }

        /// <summary>
        /// Starts one IMCA action on a background STA thread when a thread slot is available.
        /// The counter is always released in finally, including when the action throws an exception.
        /// </summary>
        private bool StartActionThread(DataRow sourceRow, string logs, string tempFolder)
        {
            int maxThreads;
            if (!int.TryParse(NUMBER_OF_THREAD_MAX, out maxThreads) || maxThreads < 1)
            {
                maxThreads = 10;
            }

            if (!TryAcquireThreadSlot(maxThreads))
            {
                return false;
            }

            // Copy the row so the worker does not depend on the lifetime of the source DataTable.
            DataTable actionTable = sourceRow.Table.Clone();
            DataRow actionRow = actionTable.NewRow();
            actionRow.ItemArray = (object[])sourceRow.ItemArray.Clone();
            actionTable.Rows.Add(actionRow);

            Thread actionThread = new Thread(() =>
            {
                try
                {
                    WriteToFile(
                        "       ACTION : " + actionRow["ACTION"] +
                        " (ID : " + actionRow["ID"] + ")" + " - Starting action thread. Active threads : " +
                        Volatile.Read(ref nb_created_thread) + "/" + maxThreads);

                    execute_IMCA_Action(actionRow, logs, tempFolder, true);
                }
                catch (Exception ex)
                {
                    WriteToFile(
                      "       ACTION : " + actionRow["ACTION"] +
                      " (ID : " + actionRow["ID"] + ")" + " - Unhandled action thread error. " + ex, Convert.ToInt64(actionRow["ID"]));

                }
                finally
                {
                    int remainingThreads = Interlocked.Decrement(ref nb_created_thread);

                    WriteToFile(
                         "       ACTION : " + actionRow["ACTION"] +
                         " (ID : " + actionRow["ID"] + ")" + " - Action thread released. Active threads : " + remainingThreads);

                    actionTable.Dispose();
                }
            });

            actionThread.IsBackground = true;
            actionThread.SetApartmentState(ApartmentState.STA);

            try
            {
                actionThread.Start();
                return true;
            }
            catch
            {
                // Release the slot if the CLR cannot start the thread.
                Interlocked.Decrement(ref nb_created_thread);
                actionTable.Dispose();
                throw;
            }
        }

        /// <summary>
        /// Atomically reserves one thread slot without exceeding NUMBER_OF_THREAD_MAX.
        /// </summary>
        private bool TryAcquireThreadSlot(int maxThreads)
        {
            while (true)
            {
                int current = Volatile.Read(ref nb_created_thread);
                if (current >= maxThreads)
                {
                    return false;
                }

                if (Interlocked.CompareExchange(ref nb_created_thread, current + 1, current) == current)
                {
                    return true;
                }
            }
        }

        /// <summary>
        /// Checks whether this service still owns database actions marked as threaded.
        /// </summary>
        private bool HasDatabaseThreadedActions()
        {
            const string sql = @"SELECT COUNT(*)
                FROM [IMCA_BACKOFFICE].[dbo].[PCM_TAB_IMCA_ACTION]
                WHERE TODO_BY=@TODO_BY AND SK_FINISH_DATE=0";

            try
            {
                using (SqlConnection connection = new SqlConnection(sql_con))
                using (SqlCommand command = new SqlCommand(sql, connection))
                {
                    command.CommandTimeout = 300;
                    command.Parameters.Add("@TODO_BY", SqlDbType.VarChar, 100).Value =
                        "THREAD_" + session_name.ToUpper();
                    connection.Open();
                    return Convert.ToInt32(command.ExecuteScalar()) > 0;
                }
            }
            catch (Exception ex)
            {
                // SQL may be unavailable during shutdown. Do not block Windows service termination.
                WriteToFile("Unable to check threaded actions during shutdown : " + ex.Message);
                return Volatile.Read(ref nb_created_thread) > 0;
            }
        }

        protected void check_if_TODO(string logs, string temp_folder)
        {
            // Publish the callback state before opening SQL connections.
            Volatile.Write(ref timer_check_TODO_is_running, true);

            try
            {


                using (SqlConnection con = new SqlConnection(sql_con))
                {
                    con.Open();

                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.Connection = con;
                        cmd.CommandTimeout = 300;

                        Boolean bol = true;
                        do
                        {
                            // Check if a action has been reserved (with the session name) and not started/finished
                            cmd.CommandText = " SET LANGUAGE FRENCH;SELECT action.ID,admin.ACTION,admin.ID as [ADMIN_ID],admin.SK_VALID," +
                                              " CASE WHEN EXISTS (SELECT 1 FROM [IMCA_BACKOFFICE].[dbo].[PCM_TAB_IMCA_ACTION_FLAG] thread_flag " +
                                              " WHERE thread_flag.ACTION_ID=admin.ID AND UPPER(thread_flag.FLAG)='USE_THREAD') " +
                                              " THEN 1 ELSE 0 END AS USE_THREAD, " +
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
                                // USE_THREAD is configured per action in PCM_TAB_IMCA_ACTION_FLAG.
                                // The legacy global switch remains a fallback for actions without the flag.
                                bool actionRequestsThread =
                                    dr.Table.Columns.Contains("USE_THREAD") &&
                                    Convert.ToInt32(dr["USE_THREAD"]) == 1;

                                bool useThread = actionRequestsThread ||
                                    EXECUTE_IMCA_ACTION_USING_THREAD.Trim().Equals("TRUE", StringComparison.OrdinalIgnoreCase);

                                if (useThread)
                                {
                                    if (!StartActionThread(dr, logs, temp_folder))
                                    {
                                        WriteToFile(
                                            "       Thread limit reached. Action remains reserved for the next timer cycle : " +
                                            dr["ACTION"] + " (ID : " + dr["ID"] + ")");
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
                                            cmd2.CommandTimeout = 300;
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


            }
            catch (Exception ex)
            {
                LogDatabaseError(nameof(check_if_TODO), "IMCA action polling cycle", ex);
                throw;
            }
            finally
            {
                // Never leave the service stuck in a running state after a SQL/network error.
                Volatile.Write(ref timer_check_TODO_is_running, false);
            }
        }
        // Writes database failures to a local file and never depends on SQL Server.
        private void LogDatabaseError(string methodName, string commandText, Exception exception, long actionId = 0, string actionName = "")
        {
            try
            {
                const string indent = "       ";
                WriteToFile(
                    indent + "DATABASE ERROR" + Environment.NewLine +
                    indent + "   Method  : " + methodName + Environment.NewLine +
                    indent + "   Type    : " + exception.GetType().FullName + Environment.NewLine +
                    indent + "   Message : " + exception.Message + Environment.NewLine +
                    indent + "   Command : " + (string.IsNullOrWhiteSpace(commandText) ? "<not available>" : commandText) + Environment.NewLine +
                    indent + "   Details : " + exception,
                    actionId, logs_folder, actionName);
            }
            catch
            {
                // A logging failure must never hide the original database exception.
            }
        }

        // Best-effort database update. SQL outages are logged locally and never mask the original error.
        private void TryMarkActionAsError(long actionId, long functionId, Exception exception)
        {
            const string sql = @"UPDATE [IMCA_BACKOFFICE].[dbo].[PCM_TAB_IMCA_ACTION]
SET TODO_BY=@TODO_BY, ERROR=@ERROR_NUMBER, ERROR_TEXT=@ERROR_MSG WHERE ID=@ID";
            if (actionId <= 0) return;
            try
            {
                using (SqlConnection connection = new SqlConnection(sql_con))
                using (SqlCommand command = new SqlCommand(sql, connection))
                {
                    command.CommandTimeout = 300;
                    command.Parameters.Add("@TODO_BY", SqlDbType.VarChar, 100).Value = "ERROR_" + session_name.ToUpperInvariant();
                    command.Parameters.Add("@ID", SqlDbType.BigInt).Value = actionId;
                    command.Parameters.Add("@ERROR_NUMBER", SqlDbType.Int).Value = exception is SqlException sqlException ? sqlException.Number : 0;
                    command.Parameters.Add("@ERROR_MSG", SqlDbType.NVarChar, 1024).Value = Strings.Left("FUNCTION_ID=" + functionId + " - ERROR MSG=" + exception.Message, 1024);
                    connection.Open();
                    command.ExecuteNonQuery();
                }
            }
            catch (Exception logException)
            {
                LogDatabaseError(nameof(TryMarkActionAsError), sql, logException, actionId);
            }
        }

        public void WriteToFile(string message, long id = 0, string logs_folder = "logs", string action_name = "")
        {
            string targetFolder = service_path + "\\" + logs_folder;

            lock (logLock)
            {
                if (!Directory.Exists(targetFolder))
                {
                    Directory.CreateDirectory(targetFolder);
                }

                string filePath;
                if (id == 0)
                {
                    filePath = targetFolder + "\\IMCA_" + session_name.ToUpper() + "_" +
                               DateTime.Now.Date.ToString("dd_MM_yyyy") + ".txt";
                }
                else
                {
                    filePath = targetFolder + "\\IMCA_" + session_name.ToUpper() + "_" +
                               DateTime.Now.Date.ToString("dd_MM_yyyy") + "_XX_" +
                               action_name.ToUpper() + "_ACTION_ID_" + id + ".txt";
                }

                // AppendAllText creates the file when it does not already exist.
                File.AppendAllText(filePath, message + Environment.NewLine);
            }
        }
    }

}
