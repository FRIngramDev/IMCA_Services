using Microsoft.Graph;
using Microsoft.Graph.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;

namespace Modele_multi_sharemailboxes
{
    public class Modele_multi_sharemailboxes
    {
        public Modele_multi_sharemailboxes()
        {
        }

        // ----- Current country / processing parameters -----
        private string country = "";                       // Country code being processed
        private string sk_valid = "";                      // SK_VALID parameter
        private string name = "";                          // Readable country / instance name
        private string active = "";                        // Indicates whether processing is active
        private string debug = "";                         // Enables detailed logs when TRUE
        private string start_date_scan = "";               // Start date used to search emails
        private string number_of_mails = "10";             // Maximum number of emails to retrieve

        // ----- Current shared mailbox parameters -----
        private string sharedmailbox_name = "";            // Current shared mailbox address
        private string sharedmailbox_folder_in = "";       // Source folder
        private string sharedmailbox_folder_out = "";      // Archive folder
        private string table_name = "";                    // Business destination table
        private string type_datas = "";                    // Expected data type: IMEI or SN

        // ----- Technical alert / recipient parameters -----
        private string email_in_case_of_technical_issue_parameter_global = "";
        private string email_in_case_of_technical_issue = "";
        private string email_destinataire = "";

        // ----- Files / logs / session -----
        private string logs_folder = "";
        private string temp_folder = "";
        private string global_session_name = "";

        // ----- Application name -----
        private string global_application_name = "IMEI_MANAGEMENT_FR";

        // ----- SQL connection parameters -----
        private string sql_connexion = "";
        private string sql_connexion_parameter_global = "";

        // ----- Microsoft Graph client -----
        private GraphServiceClient graphService = null;

        /// <summary>
        /// Represents the JSON configuration stored in the IMCA global parameters table.
        /// </summary>
        public class JSON_file
        {
            public string logs_folder { get; set; } = "";
            public List<Country> countries { get; set; } = new List<Country>();
        }

        public class Country
        {
            public string country { get; set; } = "";
            public string sk_valid { get; set; } = "";
            public string name { get; set; } = "";
            public string active { get; set; } = "";
            public string debug { get; set; } = "";
            public string start_date_scan { get; set; } = "";
            public string number_of_mails { get; set; } = "10";

            public List<SharedMailbox> sharedmailboxes { get; set; } = new List<SharedMailbox>();
            public string sharedmailbox_folder_in { get; set; } = "INBOX";
            public string sharedmailbox_folder_out { get; set; } = "Archives";

            public string email_in_case_of_technical_issue_parameter_global { get; set; } = "";
            public string email_destinataire { get; set; } = "";
            public string sql_connexion_parameter_global { get; set; } = "";
        }

        public class SharedMailbox
        {
            public string sharedmailbox_name { get; set; } = "";
            public string table_name { get; set; } = "";
            public string type_datas { get; set; } = "";
        }

        /// <summary>
        /// Creates and returns a Microsoft Graph client.
        /// </summary>
        private GraphServiceClient Connexion_Microsoft_Graph()
        {
            System.Net.ServicePointManager.SecurityProtocol =
                System.Net.SecurityProtocolType.Tls12;

            GraphServiceClient GraphService;

            class_dev_tools.Ews_Modern_Auth Ews_Modern_Auth = new class_dev_tools.Ews_Modern_Auth();

            GraphService = Ews_Modern_Auth.Get_Graph_Service();

            return GraphService;
        }

        /// <summary>
        /// Main entry point used by IMCA to read emails from every configured shared mailbox.
        /// The business processing must be added in ProcessEmailBusinessData.
        /// </summary>
        public void Read_Email_with_Graph(
            string sql_con,
            string logs,
            string tmp_folder,
            string session_name)
        {
            string global_parameters = "";
            string service_path = Path.GetDirectoryName(
                System.Reflection.Assembly.GetEntryAssembly().Location);

            setlogs_folder(service_path + "\\" + logs);
            settemp_folder(service_path + "\\" + tmp_folder);
            global_session_name = session_name;

            try
            {
                // Get the global parameters of the IMCA action.
                global_parameters = get_IMCA_paramters(
                    sql_con,
                    global_application_name);

                if (string.IsNullOrWhiteSpace(global_parameters))
                {
                    WriteToFile(
                        "No parameters found for " + global_application_name);
                    return;
                }

                JSON_file param =
                    JsonConvert.DeserializeObject<JSON_file>(global_parameters);

                if (param == null ||
                    param.countries == null ||
                    param.countries.Count == 0)
                {
                    WriteToFile(
                        global_application_name +
                        " parameters are empty or invalid");
                    return;
                }

                foreach (Country p in param.countries)
                {
                    setParamCountry(p.country);
                    setParamSK_Valid(p.sk_valid);
                    setParamName(p.name);
                    setParamActive(p.active);
                    setParamDebug(p.debug);
                    setStartDateScan(p.start_date_scan);
                    setNumber_of_mails(p.number_of_mails);

                    setsharedmailbox_folder_in(p.sharedmailbox_folder_in);
                    setsharedmailbox_folder_out(p.sharedmailbox_folder_out);

                    setEmailDestinataire(p.email_destinataire);
                    setSqlConnexionParam(p.sql_connexion_parameter_global);
                    setEmailInCaseOfTechnicalIssueParam(p.email_in_case_of_technical_issue_parameter_global);

                    email_in_case_of_technical_issue = get_IMCA_paramters(sql_con, email_in_case_of_technical_issue_parameter_global);

                    sql_connexion = get_IMCA_paramters(sql_con, sql_connexion_parameter_global);

                    if (active.ToUpper().Trim() != "TRUE")
                    {
                        continue;
                    }

                    try
                    {
                        WriteToFile(
                            name.ToUpper() + "(" + country.ToUpper() + ")" +
                            " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));

                        WriteToFile(
                            "   Debug Parameter is set to " + debug.ToUpper() +
                            " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));

                        ValidateRequiredCountryParameters(p);

                        graphService = Connexion_Microsoft_Graph();

                        if (!Directory.Exists(temp_folder))
                        {
                            Directory.CreateDirectory(temp_folder);
                        }

                        DeleteFiles(temp_folder, global_application_name);

                        // Process each mailbox independently so one mailbox failure
                        // does not prevent the other mailboxes from being scanned.
                        foreach (SharedMailbox mailbox in p.sharedmailboxes)
                        {
                            setSharedMailboxName(mailbox.sharedmailbox_name);
                            setTableName(mailbox.table_name);
                            setTypeDatas(mailbox.type_datas);

                            try
                            {
                                ReadCurrentSharedMailbox();
                            }
                            catch (Exception mailboxException)
                            {
                                WriteToFile(
                                    "   Error reading mailbox " +
                                    sharedmailbox_name + " : " +
                                    mailboxException.Message);
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        WriteToFile("   Error get emails : " + e.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteToFile(
                    "Global error Read_Email_with_Graph : " + ex.Message);
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        /// <summary>
        /// Reads the emails for the currently selected shared mailbox.
        /// </summary>
        private void ReadCurrentSharedMailbox()
        {
            int nb_mail = 0;

            ValidateRequiredMailboxParameters();

            if (debug.ToUpper() == "TRUE")
            {
                WriteToFile(
                    "   Connexion to " + sharedmailbox_name +
                    " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                WriteToFile("   Sharedmailbox_folder_in  : " + sharedmailbox_folder_in);
                WriteToFile("   Sharedmailbox_folder_out : " + sharedmailbox_folder_out);
                WriteToFile("   Table name               : " + table_name);
                WriteToFile("   Data type                : " + type_datas);
                WriteToFile("   Temp folder              : " + temp_folder);
                WriteToFile("   Extracting the " + number_of_mails + " oldest messages");
            }

            MailFolder inputFolder = GetInputFolder(graphService);
            MessageCollectionResponse messages =
                GetMessagesToProcess(graphService, inputFolder.Id);

            if (messages != null &&
                messages.Value != null &&
                messages.Value.Count > 0)
            {
                if (debug.ToUpper() == "TRUE")
                {
                    WriteToFile(
                        "   " + messages.Value.Count +
                        " Email(s) found at " +
                        DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                }

                foreach (Message email in messages.Value)
                {
                    try
                    {
                        if (debug.ToUpper() == "TRUE")
                        {
                            WriteToFile(
                                "       Subject : " + email.Subject +
                                " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                        }

                        // This is the only section intended to contain mailbox-specific
                        // and email-specific business processing.
                        ProcessEmailBusinessData(email);

                        nb_mail++;
                        System.Threading.Thread.Sleep(500);
                    }
                    catch (Exception ex)
                    {
                        WriteToFile(
                            "   Error processing email : " + ex.Message +
                            " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                    }
                }

                if (debug.ToUpper() == "TRUE")
                {
                    WriteToFile(
                        "   " + nb_mail +
                        " Email(s) have been processed at " +
                        DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                }
            }
            else if (debug.ToUpper() == "TRUE")
            {
                WriteToFile("   No emails found");
            }
        }

        /// <summary>
        /// Business processing placeholder.
        /// Add the logic that depends on the current mailbox and message here.
        /// </summary>
        private void ProcessEmailBusinessData(Message email)
        {
            // Current mailbox context:
            // sharedmailbox_name -> mailbox address
            // table_name         -> destination business table
            // type_datas         -> IMEI or SN
            // email              -> current Microsoft Graph message
            // sql_connexion      -> resolved business SQL connection

            // TODO: Add IMEI / serial-number business processing here.
        }

        private MessageCollectionResponse GetMessagesToProcess(
            GraphServiceClient graphService,
            string folderId)
        {
            int topEmails = 10;

            if (!int.TryParse(number_of_mails, out topEmails))
            {
                topEmails = 10;
            }

            return graphService.Users[sharedmailbox_name]
                .MailFolders[folderId]
                .Messages
                .GetAsync(config =>
                {
                    config.QueryParameters.Top = topEmails;
                    config.QueryParameters.Orderby =
                        new string[] { "receivedDateTime asc" };
                    config.QueryParameters.Select = new string[]
                    {
                        "id",
                        "subject",
                        "from",
                        "hasAttachments",
                        "receivedDateTime",
                        "isRead"
                    };

                    string filter = BuildMessageFilter();

                    if (!string.IsNullOrWhiteSpace(filter))
                    {
                        config.QueryParameters.Filter = filter;
                    }
                })
                .GetAwaiter()
                .GetResult();
        }

        private string BuildMessageFilter()
        {
            if (string.IsNullOrWhiteSpace(start_date_scan))
            {
                return "";
            }

            DateTime searchDate;

            if (DateTime.TryParse(start_date_scan, out searchDate))
            {
                return "receivedDateTime ge " +
                    searchDate.ToUniversalTime()
                        .ToString("yyyy-MM-ddTHH:mm:ssZ");
            }

            if (start_date_scan.Length == 8 &&
                DateTime.TryParseExact(
                    start_date_scan,
                    "yyyyMMdd",
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out searchDate))
            {
                return "receivedDateTime ge " +
                    searchDate.ToUniversalTime()
                        .ToString("yyyy-MM-ddTHH:mm:ssZ");
            }

            WriteToFile(
                "   Invalid start_date_scan value : " + start_date_scan);
            return "";
        }

        private MailFolder GetInputFolder(GraphServiceClient graphService)
        {
            if (sharedmailbox_folder_in.ToUpper().Trim() == "INBOX")
            {
                return graphService.Users[sharedmailbox_name]
                    .MailFolders["inbox"]
                    .GetAsync()
                    .GetAwaiter()
                    .GetResult();
            }

            return GetChildFolderByName(
                graphService,
                sharedmailbox_name,
                sharedmailbox_folder_in);
        }

        private MailFolder GetChildFolderByName(
            GraphServiceClient graphService,
            string mailbox,
            string folderName)
        {
            string safeFolderName = EscapeODataString(folderName);

            MailFolderCollectionResponse folders = graphService.Users[mailbox]
                .MailFolders["inbox"]
                .ChildFolders
                .GetAsync(config =>
                {
                    config.QueryParameters.Filter =
                        $"displayName eq '{safeFolderName}'";
                })
                .GetAwaiter()
                .GetResult();

            if (folders == null ||
                folders.Value == null ||
                folders.Value.Count == 0)
            {
                throw new Exception("Folder not found : " + folderName);
            }

            return folders.Value.First();
        }

        private void ValidateRequiredCountryParameters(Country p)
        {
            if (p.sharedmailboxes == null || p.sharedmailboxes.Count == 0)
            {
                throw new Exception("sharedmailboxes is empty");
            }

            if (string.IsNullOrWhiteSpace(temp_folder))
            {
                throw new Exception("temp_folder is empty");
            }

            if (string.IsNullOrWhiteSpace(sql_connexion))
            {
                throw new Exception(
                    "sql_connexion is empty. Parameter used : " +
                    sql_connexion_parameter_global);
            }

            if (string.IsNullOrWhiteSpace(email_in_case_of_technical_issue))
            {
                throw new Exception(
                    "email_in_case_of_technical_issue is empty. Parameter used : " +
                    email_in_case_of_technical_issue_parameter_global);
            }
        }

        private void ValidateRequiredMailboxParameters()
        {
            if (string.IsNullOrWhiteSpace(number_of_mails))
            {
                number_of_mails = "10";
            }

            if (string.IsNullOrWhiteSpace(sharedmailbox_folder_in))
            {
                sharedmailbox_folder_in = "INBOX";
            }

            if (string.IsNullOrWhiteSpace(sharedmailbox_folder_out))
            {
                sharedmailbox_folder_out = "Archives";
            }

            if (string.IsNullOrWhiteSpace(sharedmailbox_name))
            {
                throw new Exception("sharedmailbox_name is empty");
            }

            if (string.IsNullOrWhiteSpace(table_name))
            {
                throw new Exception(
                    "table_name is empty for " + sharedmailbox_name);
            }

            if (type_datas.ToUpper().Trim() != "IMEI" &&
                type_datas.ToUpper().Trim() != "SN")
            {
                throw new Exception(
                    "type_datas must be IMEI or SN for " +
                    sharedmailbox_name);
            }
        }

        private string get_IMCA_paramters(
            string sql_con,
            string param_name)
        {
            string ret = "";

            using (SqlConnection con = new SqlConnection(sql_con))
            {
                con.Open();

                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.Connection = con;
                    cmd.CommandTimeout = 300;
                    cmd.CommandText = @"
                        SELECT ISNULL(VALUE, '') AS VALUE
                        FROM [PCM_TAB_IMCA_PARAMETER_GLOBAL]
                        WHERE SK_VALID = 0
                        AND PARAMETER = @PARAMETER";

                    cmd.Parameters.AddWithValue(
                        "@PARAMETER",
                        param_name);

                    object result = cmd.ExecuteScalar();

                    if (result != null)
                    {
                        ret = result.ToString();
                    }
                }
            }

            return ret;
        }

        private void WriteToFile(string message)
        {
            if (string.IsNullOrWhiteSpace(logs_folder))
            {
                logs_folder = AppDomain.CurrentDomain.BaseDirectory;
            }

            if (!Directory.Exists(logs_folder))
            {
                Directory.CreateDirectory(logs_folder);
            }

            string filePath = Path.Combine(
                logs_folder,
                "IMCA_" + global_session_name + "_" +
                DateTime.Now.ToString("dd_MM_yyyy") + "_" +
                country + "_" + global_application_name + ".txt");

            File.AppendAllText(
                filePath,
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") +
                " - " + message + Environment.NewLine);
        }

        private string EscapeODataString(string value)
        {
            return value == null ? "" : value.Replace("'", "''");
        }

        private void DeleteFiles(string folder, string prefixfile = "")
        {
            if (!Directory.Exists(folder))
            {
                return;
            }

            foreach (string file in Directory.GetFiles(folder))
            {
                try
                {
                    if (Path.GetFileName(file)
                        .ToLower()
                        .StartsWith(prefixfile.ToLower()))
                    {
                        File.Delete(file);
                    }
                }
                catch
                {
                    // Ignore cleanup errors so they do not block mailbox processing.
                }
            }
        }

        public void setParamCountry(string value) { country = value ?? ""; }
        public void setParamSK_Valid(string value) { sk_valid = value ?? ""; }
        public void setParamName(string value) { name = value ?? ""; }
        public void setParamActive(string value) { active = value ?? ""; }
        public void setParamDebug(string value) { debug = value ?? ""; }
        public void setStartDateScan(string value) { start_date_scan = value ?? ""; }

        public void setNumber_of_mails(string value)
        {
            number_of_mails = string.IsNullOrWhiteSpace(value) ? "10" : value;
        }

        public void setSharedMailboxName(string value)
        {
            sharedmailbox_name = value ?? "";
        }

        public void setsharedmailbox_folder_in(string value)
        {
            sharedmailbox_folder_in =
                string.IsNullOrWhiteSpace(value) ? "INBOX" : value;
        }

        public void setsharedmailbox_folder_out(string value)
        {
            sharedmailbox_folder_out =
                string.IsNullOrWhiteSpace(value) ? "Archives" : value;
        }

        public void setTableName(string value) { table_name = value ?? ""; }
        public void setTypeDatas(string value) { type_datas = value ?? ""; }
        public void setEmailDestinataire(string value) { email_destinataire = value ?? ""; }
        public void setlogs_folder(string value) { logs_folder = value ?? ""; }
        public void settemp_folder(string value) { temp_folder = value ?? ""; }

        public void setSqlConnexionParam(string value)
        {
            sql_connexion_parameter_global = value ?? "";
        }

        public void setEmailInCaseOfTechnicalIssueParam(string value)
        {
            email_in_case_of_technical_issue_parameter_global = value ?? "";
        }
    }
}
