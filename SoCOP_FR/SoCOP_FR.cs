using Microsoft.Graph;
using Microsoft.Graph.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SoCOP_FR
{
    public class SoCOP_FR
    {
        private string country = "";
        private string sk_valid = "";
        private string name = "";
        private string active = "";
        private string debug = "";
        private string start_date_scan = "";
        private string number_of_mails = "10";

        private string sharedmailbox_name = "";
        private string sharedmailbox_folder_in = "";
        private string sharedmailbox_folder_out = "";

        private string email_in_case_of_technical_issue_parameter_global = "";
        private string email_in_case_of_technical_issue = "";

        private string logs_folder = "";
        private string temp_folder = "";
        private string global_session_name = "";

        private string global_application_name = "SOCOP_FR";

        private string sql_connexion = "";

        private string sql_connexion_parameter_global = "";

        private GraphServiceClient current_graph_service = null;

        public SoCOP_FR()
        {
        }

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

            public string sharedmailbox_name { get; set; } = "";
            public string sharedmailbox_folder_in { get; set; } = "INBOX";
            public string sharedmailbox_folder_out { get; set; } = "Archives";

            public string email_in_case_of_technical_issue_parameter_global { get; set; } = "";

            public string sql_connexion_parameter_global { get; set; } = "";

        }
        private GraphServiceClient Connexion_Microsoft_Graph()
        {
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;

            // On utilise EWS API MANAGED REFERENCE 2.1
            GraphServiceClient GraphService;

            class_dev_tools.Ews_Modern_Auth Ews_Modern_Auth = new class_dev_tools.Ews_Modern_Auth();

            GraphService = Ews_Modern_Auth.Get_Graph_Service();

            return GraphService;
        }

        public void Read_Email_with_Graph(string sql_con, string logs, string tmp_folder, string session_name)
        {
            string global_parameters = "";
            string service_path = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
            int nb_mail = 0;


            setlogs_folder(service_path + "\\" + logs);
            settemp_folder(service_path + "\\" + tmp_folder);
            global_session_name = session_name;

            try
            {
                // Get the global parameters of the ACTION
                global_parameters = get_IMCA_paramters(sql_con, global_application_name);

                if (string.IsNullOrWhiteSpace(global_parameters))
                {
                    WriteToFile("No parameters found for " + global_application_name);
                    return;
                }

                JSON_file param = JsonConvert.DeserializeObject<JSON_file>(global_parameters);

                if (param == null || param.countries == null || param.countries.Count == 0)
                {
                    WriteToFile(global_application_name + " parameters are empty or invalid");
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

                    setSharedMailboxName(p.sharedmailbox_name);
                    setsharedmailbox_folder_in(p.sharedmailbox_folder_in);
                    setsharedmailbox_folder_out(p.sharedmailbox_folder_out);

                    setNumber_of_mails(p.number_of_mails);


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
                        WriteToFile(name.ToUpper() + "(" + country.ToUpper() + ")" + " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                        WriteToFile("   Debug Parameter is set to " + debug.ToUpper() + " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));

                        if (debug.ToUpper() == "TRUE")
                        {
                            WriteToFile("   Connexion to " + sharedmailbox_name + " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                        }

                        GraphServiceClient graphService = Connexion_Microsoft_Graph();

                  
                        ValidateRequiredParameters();

                        if (!Directory.Exists(temp_folder))
                        {
                            Directory.CreateDirectory(temp_folder);
                        }

                       // reste à coder

                        if (debug.ToUpper() == "TRUE")
                        {
                            WriteToFile("   Sharedmailbox_folder_in  : " + sharedmailbox_folder_in);
                            WriteToFile("   Sharedmailbox_folder_out : " + sharedmailbox_folder_out);
                            WriteToFile("   Temp folder              : " + temp_folder);
                            WriteToFile("   Extracting the " + number_of_mails + " oldest messages");
                        }

                        MailFolder inputFolder = GetInputFolder(graphService);
                        MailFolder archiveFolder = GetChildFolderByName(graphService, sharedmailbox_name, sharedmailbox_folder_out);
                        MailFolder errorFolder = GetChildFolderByName(graphService, sharedmailbox_name, "Erreur");

                        MessageCollectionResponse messages = GetMessagesToProcess(graphService, inputFolder.Id);

                        nb_mail = 0;

                        if (messages != null && messages.Value != null && messages.Value.Count > 0)
                        {
                            if (debug.ToUpper() == "TRUE")
                            {
                                WriteToFile("   " + messages.Value.Count + " Email(s) found at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                            }

                            foreach (Message email in messages.Value)
                            {
                                try
                                {
                                    if (debug.ToUpper() == "TRUE")
                                    {
                                        WriteToFile("       Subject : " + email.Subject + " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                                    }

                                    if (email.HasAttachments == true)
                                    {
                                        List<string> attachments = DownloadExcelAttachmentsWithGraph(
                                            graphService,
                                            sharedmailbox_name,
                                            email.Id
                                        );

                                        if (attachments.Count == 0)
                                        {
                                            WriteToFile("       No Excel attachment found for email : " + email.Subject);
                                        }

                                        foreach (string strAttachment in attachments)
                                        {
                                            if (debug.ToUpper() == "TRUE")
                                            {
                                                WriteToFile("       Attachment : " + strAttachment);
                                            }

                                            ImportationFichierExcelEnBase(strAttachment);
                                        }
                                    }
                                    else
                                    {
                                        WriteToFile("       No attachment found for email : " + email.Subject);
                                    }

                                    MarkEmailAsRead(graphService, sharedmailbox_name, email.Id);
                                    MoveEmail(graphService, sharedmailbox_name, email.Id, archiveFolder.Id);

                                    nb_mail++;

                                    System.Threading.Thread.Sleep(500);
                                }
                                catch (Exception ex)
                                {
                                    WriteToFile("   Error processing email : " + ex.Message + " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));

                                    try
                                    {
                                        string recipient = !string.IsNullOrWhiteSpace(email_in_case_of_technical_issue)
                                            ? email_in_case_of_technical_issue
                                            : email_to;

                                        EnvoiEmail_with_Graph(
                                            graphService,
                                            "Importation des cotations Brother",
                                            "import_cotation_brother - Une erreur (" + ex.Message + ") est survenue lors du traitement d'un fichier. Le mail a été déplacé dans le dossier Erreur.",
                                            recipient
                                        );
                                    }
                                    catch (Exception mailEx)
                                    {
                                        WriteToFile("   Error sending alert email : " + mailEx.Message);
                                    }

                                    try
                                    {
                                        MoveEmail(graphService, sharedmailbox_name, email.Id, errorFolder.Id);
                                    }
                                    catch (Exception moveEx)
                                    {
                                        WriteToFile("   Error moving email to Erreur folder : " + moveEx.Message);
                                    }
                                }
                            }

                            if (debug.ToUpper() == "TRUE")
                            {
                                WriteToFile("   " + nb_mail + " Email(s) have been processed at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                            }
                        }
                        else
                        {
                            if (debug.ToUpper() == "TRUE")
                            {
                                WriteToFile("   No emails found");
                            }
                        }

                        messages = null;
                    }
                    catch (Exception e)
                    {
                        WriteToFile("   Error get emails : " + e.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteToFile("Global error Read_Email_with_Graph : " + ex.Message);
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void ValidateRequiredParameters()
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

            if (string.IsNullOrWhiteSpace(temp_folder))
            {
                throw new Exception("temp_folder is empty");
            }

            if (string.IsNullOrWhiteSpace(sql_connexion))
            {
                throw new Exception("sql_connexion is empty. Parameter used : " + sql_connexion_parameter_global);
            }

            if (string.IsNullOrWhiteSpace(email_in_case_of_technical_issue))
            {
                throw new Exception("email_in_case_of_technical_issue is empty. Parameter used : " + email_in_case_of_technical_issue_parameter_global);
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
                    cmd.CommandText = @"
                        SELECT ISNULL(VALUE, '') AS VALUE
                        FROM [PCM_TAB_IMCA_PARAMETER_GLOBAL]
                        WHERE SK_VALID = 0
                        AND PARAMETER = @PARAMETER";

                    cmd.Parameters.AddWithValue("@PARAMETER", param_name);

                    object result = cmd.ExecuteScalar();

                    if (result != null)
                    {
                        ret = result.ToString();
                    }
                }

                con.Close();
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
                "IMCA_" + global_session_name + "_" + DateTime.Now.ToString("dd_MM_yyyy") + "_" + country + "_" + global_application_name + ".txt"
            );

            File.AppendAllText(
                filePath,
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - " + message + Environment.NewLine
            );
        }

        public void setParamCountry(string country)
        {
            this.country = country ?? "";
        }

        public void setParamSK_Valid(string sk_valid)
        {
            this.sk_valid = sk_valid ?? "";
        }

        public void setParamName(string name)
        {
            this.name = name ?? "";
        }

        public void setParamActive(string active)
        {
            this.active = active ?? "";
        }

        public void setParamDebug(string debug)
        {
            this.debug = debug ?? "";
        }

        public void setStartDateScan(string start_date_scan)
        {
            this.start_date_scan = start_date_scan ?? "";
        }

        public void setNumber_of_mails(string number_of_mails)
        {
            this.number_of_mails = string.IsNullOrWhiteSpace(number_of_mails) ? "10" : number_of_mails;
        }

        public void setSharedMailboxName(string sharedmailbox_name)
        {
            this.sharedmailbox_name = sharedmailbox_name ?? "";
        }

        public void setsharedmailbox_folder_in(string sharedmailbox_folder_in)
        {
            this.sharedmailbox_folder_in = string.IsNullOrWhiteSpace(sharedmailbox_folder_in) ? "INBOX" : sharedmailbox_folder_in;
        }

        public void setsharedmailbox_folder_out(string sharedmailbox_folder_out)
        {
            this.sharedmailbox_folder_out = string.IsNullOrWhiteSpace(sharedmailbox_folder_out) ? "Archives" : sharedmailbox_folder_out;
        }

        public void setEmailInCaseOfTechnicalIssue(string email_in_case_of_technical_issue)
        {
            this.email_in_case_of_technical_issue = email_in_case_of_technical_issue ?? "";
        }


        public void setlogs_folder(string logs_folder)
        {
            this.logs_folder = logs_folder ?? "";
        }

        public void settemp_folder(string temp_folder)
        {
            this.temp_folder = temp_folder ?? "";
        }

        public void setSqlConnexion(string sql_connexion)
        {
            this.sql_connexion = sql_connexion ?? "";
        }


        public void setSqlConnexionParam(string sql_connexion_parameter_global)
        {
            this.sql_connexion_parameter_global = sql_connexion_parameter_global ?? "";
        }

        public void setEmailInCaseOfTechnicalIssueParam(string email_in_case_of_technical_issue_parameter_global)
        {
            this.email_in_case_of_technical_issue_parameter_global = email_in_case_of_technical_issue_parameter_global ?? "";
        }


    }
}
