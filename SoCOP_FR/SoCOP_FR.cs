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

        private GraphServiceClient graphService = null;

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

                        graphService = Connexion_Microsoft_Graph();

                  
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

                                    if (!string.IsNullOrWhiteSpace(email.Subject) && email.Subject.ToUpper().Contains("CISCO")==false && email.Subject.ToUpper().Contains("RE:") == false)
                                    {
                                        if (!string.IsNullOrWhiteSpace(email.From.EmailAddress.Address.ToString()) && email.From.EmailAddress.Address.ToString().ToLower().Trim() != "automail@ingrammicro.com")
                                        {
                                            //On archive pour l'instant les mails 

                                            //Flague le mail comme lu
                                            MarkEmailAsRead(graphService, sharedmailbox_name, email.Id);

                                            // Déplace le mail dans le dossier Erreur
                                            MoveEmail(graphService, sharedmailbox_name, email.Id, errorFolder.Id);
                                        }
                                        else
                                        {
                                            if (!string.IsNullOrWhiteSpace(email.Subject) && email.Subject.ToUpper().Contains("SO COP COMPLETED") == false)
                                            {
                                                //Flague le mail comme lu
                                                MarkEmailAsRead(graphService, sharedmailbox_name, email.Id);

                                                // Déplace le mail dans le dossier Erreur
                                                MoveEmail(graphService, sharedmailbox_name, email.Id, errorFolder.Id);

                                            }
                                            else
                                            {
                                                //Numero du BL dans le bon format + les personnes en copie pour T_deblocage_sales
                                                string num_bl = email.Subject.Substring(25, 11).Replace("-", "");

                                                // On recupére l'adresse du /(des) destinataire(s)
                                                string list_recipients = email.ToRecipients != null ? string.Join(";", email.ToRecipients.Select(r => r.EmailAddress.Address)) : "";

                                                // On insert dans T_deblocage_sales
                                                UpdateT_deblocage_sales(sql_connexion, num_bl, list_recipients);

                                                //Flague le mail comme lu
                                                MarkEmailAsRead(graphService, sharedmailbox_name, email.Id);

                                                // Déplace le mail dans le dossier Archive
                                                MoveEmail(graphService, sharedmailbox_name, email.Id, archiveFolder.Id);

                                            }

                                        }


                                    }
                                    else
                                    {
                                        //Flague le mail comme lu
                                        MarkEmailAsRead(graphService, sharedmailbox_name, email.Id);

                                        // Déplace le mail dans le dossier Archive
                                        MoveEmail(graphService, sharedmailbox_name, email.Id, archiveFolder.Id);

                                    }

                                    nb_mail++;

                                    System.Threading.Thread.Sleep(500);
                                }
                                catch (Exception ex)
                                {
                                    WriteToFile("   Error processing email : " + ex.Message + " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));


                                    try
                                    {
                                        EnvoiEmail_with_Graph(
                                            graphService,
                                            global_application_name,
                                            global_application_name + " - Une erreur (" + ex.Message + ") est survenue lors du traitement du mail. Le mail a été déplacé dans le dossier Erreur.",
                                            email_in_case_of_technical_issue
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


        private void UpdateT_deblocage_sales(string con, string num_bl, string Destinataires)
        {
            string sql = @"INSERT INTO T_deblocage_sales(nom_user, date_demande, num_bl, type_demande, top_traite, email_demandeur) VALUES ('macro_socop', getdate(), @num_bl, 'AUTO', 'N', @email_demandeur)" ;

            try
            {
                using (SqlConnection conn = new SqlConnection(con))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.CommandTimeout = 300;
                        cmd.Parameters.AddWithValue("@num_bl", num_bl);
                        cmd.Parameters.AddWithValue("@email_demandeur", Destinataires);
                        cmd.ExecuteNonQuery();
                    }
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                SendSqlTechnicalIssueMail("UpdateT_deblocage_sales", sql, ex);
                throw;
            }
        }

        private void SendSqlTechnicalIssueMail(string methodName, string sql, Exception ex)
        {
            try
            {
                WriteToFile("SQL error in " + methodName + " : " + ex.Message);



                string subject = global_application_name + " - Erreur SQL dans " + methodName;

                string body =
                    "Une erreur SQL est survenue dans " + global_application_name + ".<br/><br/>" +
                    "<b>Méthode :</b> " + methodName + "<br/>" +
                    "<b>Timeout configuré :</b> " + 300 + " secondes<br/>" +
                    "<b>Message :</b> " + ex.Message + "<br/><br/>" +
                    "<b>Requête SQL :</b><br/>" +
                    "<pre>" + sql + "</pre>";

                EnvoiEmail_with_Graph(
                    graphService,
                    subject,
                    body,
                    email_in_case_of_technical_issue
                );
            }
            catch (Exception mailEx)
            {
                WriteToFile("Error sending SQL technical issue email : " + mailEx.Message);
            }
        }


        private void EnvoiEmail_with_Graph(GraphServiceClient graphService, string subject, string body, string recipient)
        {
            if (string.IsNullOrWhiteSpace(recipient) || !recipient.Contains("@"))
            {
                return;
            }

            List<Recipient> toRecipients = BuildRecipients(recipient);
           // List<Recipient> ccRecipients = BuildRecipients(email_cc);

            Message message = new Message
            {
                Subject = subject,
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = body.Replace(Environment.NewLine, "<br/>")
                },
                ToRecipients = toRecipients
            };

            /*
            if (ccRecipients.Count > 0)
            {
                message.CcRecipients = ccRecipients;
            }
            */
            var requestBody = new Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody
            {
                Message = message,
                SaveToSentItems = true
            };

            graphService.Users[sharedmailbox_name]
                .SendMail
                .PostAsync(requestBody)
                .GetAwaiter()
                .GetResult();
        }


        private List<Recipient> BuildRecipients(string emails)
        {
            List<Recipient> recipients = new List<Recipient>();

            if (string.IsNullOrWhiteSpace(emails))
            {
                return recipients;
            }

            string[] splitEmails = emails.Split(new char[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (string email in splitEmails)
            {
                string address = email.Trim();

                if (address.Contains("@"))
                {
                    recipients.Add(new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = address
                        }
                    });
                }
            }

            return recipients;
        }

        private void MarkEmailAsRead(GraphServiceClient graphService, string mailbox, string messageId)
        {
            Message messageUpdate = new Message
            {
                IsRead = true
            };

            graphService.Users[mailbox]
                .Messages[messageId]
                .PatchAsync(messageUpdate)
                .GetAwaiter()
                .GetResult();
        }

        private void MoveEmail(GraphServiceClient graphService, string mailbox, string messageId, string destinationFolderId)
        {
            var requestBody = new Microsoft.Graph.Users.Item.Messages.Item.Move.MovePostRequestBody
            {
                DestinationId = destinationFolderId
            };

            graphService.Users[mailbox]
                .Messages[messageId]
                .Move
                .PostAsync(requestBody)
                .GetAwaiter()
                .GetResult();
        }

        private MessageCollectionResponse GetMessagesToProcess(GraphServiceClient graphService, string folderId)
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
                    config.QueryParameters.Orderby = new string[] { "receivedDateTime asc" };
                    config.QueryParameters.Select = new string[]
                    {
                        "id",
                        "subject",
                        "from",
                        "hasAttachments",
                        "receivedDateTime",
                        "isRead",
                        "ToRecipients"
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
                return "receivedDateTime ge " + searchDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ");
            }

            if (start_date_scan.Length == 8)
            {
                searchDate = new DateTime(
                    int.Parse(start_date_scan.Substring(0, 4)),
                    int.Parse(start_date_scan.Substring(4, 2)),
                    int.Parse(start_date_scan.Substring(6, 2))
                );

                return "receivedDateTime ge " + searchDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ");
            }

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

            return GetChildFolderByName(graphService, sharedmailbox_name, sharedmailbox_folder_in);
        }

        private MailFolder GetChildFolderByName(GraphServiceClient graphService, string mailbox, string folderName)
        {
            string safeFolderName = EscapeODataString(folderName);

            MailFolderCollectionResponse folders = graphService.Users[mailbox]
                .MailFolders["inbox"]
                .ChildFolders
                .GetAsync(config =>
                {
                    config.QueryParameters.Filter = $"displayName eq '{safeFolderName}'";
                })
                .GetAwaiter()
                .GetResult();

            if (folders == null || folders.Value == null || folders.Value.Count == 0)
            {
                throw new Exception("Folder not found : " + folderName);
            }

            return folders.Value.First();
        }

        private string EscapeODataString(string value)
        {
            if (value == null)
            {
                return "";
            }

            return value.Replace("'", "''");
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
                    cmd.CommandTimeout = 300;
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
