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
using System.Text;
using ExcelDataReader;
using System.Threading.Tasks;

namespace Backorder_HP_FR
{
    public class Backorder_HP_FR
    {
        private string country = "";
        private string sk_valid = "";
        private string name = "";
        private string active = "";
        private string debug = "";
        private string start_date_scan = "";
        private string number_of_mails = "10";

        private string sharedmailbox_name = "";
        private string sharedmailbox_folder_in = "INBOX";
        private string sharedmailbox_folder_out = "Archives";

        private string email_in_case_of_technical_issue_parameter_global = "";
        private string email_in_case_of_technical_issue = "";


        private string logs_folder = "";
        private string temp_folder = "";
        private string global_session_name = "";

        private string global_application_name = "BACKORDER_HP_FR";

        private string sql_connexion = "";

        private string sql_connexion_parameter_global = "";

        private string sql_connexion_licences = "";

        private string sql_licences_parameter_global { get; set; } = "";

        private GraphServiceClient graphService = null;



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

            public string sql_licences_parameter_global { get; set; } = "";


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
                    setSqlConnexionLicencesParam(p.sql_licences_parameter_global);

                    setEmailInCaseOfTechnicalIssueParam(p.email_in_case_of_technical_issue_parameter_global);

                    email_in_case_of_technical_issue = get_IMCA_paramters(sql_con, email_in_case_of_technical_issue_parameter_global);

                    sql_connexion = get_IMCA_paramters(sql_con, sql_connexion_parameter_global);
                    sql_connexion_licences = get_IMCA_paramters(sql_con, sql_licences_parameter_global);



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

                        DeleteFiles(temp_folder, global_application_name);

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


                                        EnvoiEmail_with_Graph(
                                            graphService,
                                            "Importation du Backorder HP",
                                            "import_backorder_hp - Une erreur (" + ex.Message + ") est survenue lors du traitement d'un fichier. Le mail a été déplacé dans le dossier Erreur.",
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

        private List<string> DownloadExcelAttachmentsWithGraph(GraphServiceClient graphService, string mailbox, string messageId)
        {
            List<string> downloadedFiles = new List<string>();

            AttachmentCollectionResponse attachments = graphService.Users[mailbox]
                .Messages[messageId]
                .Attachments
                .GetAsync()
                .GetAwaiter()
                .GetResult();

            if (attachments == null || attachments.Value == null)
            {
                return downloadedFiles;
            }

            foreach (Attachment attachment in attachments.Value)
            {
                if (attachment is Microsoft.Graph.Models.FileAttachment fileAttachment)
                {
                    string attachmentName = global_application_name + '_' + (fileAttachment.Name ?? "");
                    string extension = Path.GetExtension(attachmentName).ToLower();

                    if (extension == ".xlsx" || extension == ".xls")
                    {
                        if (fileAttachment.ContentBytes == null)
                        {
                            WriteToFile("       Attachment content is empty : " + attachmentName);
                            continue;
                        }

                        string cleanFileName = CleanFileName(attachmentName);
                        string filePath = Path.Combine(temp_folder, cleanFileName);

                        File.WriteAllBytes(filePath, fileAttachment.ContentBytes);

                        downloadedFiles.Add(filePath);
                    }
                }
            }

            return downloadedFiles;
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

        private void EnvoiEmail_with_Graph(GraphServiceClient graphService, string subject, string body, string recipient)
        {
            if (string.IsNullOrWhiteSpace(recipient) || !recipient.Contains("@"))
            {
                return;
            }

            List<Recipient> toRecipients = BuildRecipients(recipient);
            //  List<Recipient> ccRecipients = BuildRecipients(email_cc);

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

            /* if (ccRecipients.Count > 0)
             {
                 message.CcRecipients = ccRecipients;
             }*/

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
            if (string.IsNullOrWhiteSpace(sql_connexion_licences))
            {
                throw new Exception("sql_connexion_licences is empty. Parameter used : " + sql_licences_parameter_global);
            }

            if (string.IsNullOrWhiteSpace(email_in_case_of_technical_issue))
            {
                throw new Exception("email_in_case_of_technical_issue is empty. Parameter used : " + email_in_case_of_technical_issue_parameter_global);
            }
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
                "IMCA_" + global_session_name + "_" + DateTime.Now.ToString("dd_MM_yyyy") + "_" + country + "_" + global_application_name +".txt"
            );

            File.AppendAllText(
                filePath,
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - " + message + Environment.NewLine
            );
        }

        private string CleanFileName(string fileName)
        {
            return string.Join("_", fileName.Split(Path.GetInvalidFileNameChars())).Trim();
        }

        private string EscapeODataString(string value)
        {
            if (value == null)
            {
                return "";
            }

            return value.Replace("'", "''");
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
                    if (Path.GetFileName(file).ToLower().StartsWith(prefixfile.ToLower()))
                    {
                        File.Delete(file);
                    }

                }
                catch
                {
                    // On ignore pour ne pas bloquer le service.
                }
            }
        }

        private void SendSqlTechnicalIssueMail(string methodName, string sql, Exception ex)
        {
            try
            {
                WriteToFile("SQL error in " + methodName + " : " + ex.Message);

                if (graphService == null)
                {
                    WriteToFile("Unable to send SQL technical issue email because GraphServiceClient is null");
                    return;
                }

                if (string.IsNullOrWhiteSpace(email_in_case_of_technical_issue) || !email_in_case_of_technical_issue.Contains("@"))
                {
                    WriteToFile("Unable to send SQL technical issue email because recipient is empty");
                    return;
                }

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


        public void setSqlConnexionLicencesParam(string sql_licences_parameter_global)
        {
            this.sql_licences_parameter_global = sql_licences_parameter_global ?? "";
        }

        public void setSqlLicences(string sql_connexion_licences)
        {
            this.sql_connexion_licences = sql_connexion_licences ?? "";
        }



        private void ImportationFichierExcelEnBase(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                throw new FileNotFoundException("Fichier Excel introuvable", filePath);

            string receptionDate = DateTime.Now.ToString("yyyyMMdd");
            try
            {
                bool newHpeFormat;
                DataTable sheet = ReadOrderStatusSheet(filePath, out newHpeFormat);

                using (SqlConnection con = new SqlConnection(sql_connexion))
                {
                    con.Open();
                    ImportExcelRows(con, sheet, newHpeFormat, receptionDate);
                }

                CopyPendingRowsToLicences(receptionDate);
                MarkRowsAsProcessed(receptionDate);
            }
            finally
            {
                try { if (File.Exists(filePath)) File.Delete(filePath); }
                catch (Exception ex) { WriteToFile("Unable to delete " + filePath + " : " + ex.Message); }
            }
        }

        private DataTable ReadOrderStatusSheet(string filePath, out bool newHpeFormat)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            DataSet dataSet;
            using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
            {
                dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = false }
                });
            }

            DataTable table = FindSheet(dataSet, "HP OrderStatus") ?? FindSheet(dataSet, "HPE OrderStatus");
            newHpeFormat = false;
            if (table == null)
            {
                table = FindSheet(dataSet, "Orders");
                newHpeFormat = table != null;
            }
            if (table == null)
                throw new Exception("Onglet HP OrderStatus, HPE OrderStatus ou Orders introuvable");
            return table;
        }

        private static DataTable FindSheet(DataSet dataSet, string name)
        {
            foreach (DataTable table in dataSet.Tables)
                if (string.Equals(table.TableName.Trim(), name, StringComparison.OrdinalIgnoreCase)) return table;
            return null;
        }

        private void ImportExcelRows(SqlConnection con, DataTable sheet, bool newHpeFormat, string receptionDate)
        {
            bool hasSerial = string.Equals(CellText(sheet, 0, "O"), "SERIALNUMBER", StringComparison.OrdinalIgnoreCase);
            for (int row = 1; row < sheet.Rows.Count; row++)
            {
                if (string.IsNullOrWhiteSpace(CellText(sheet, row, "A"))) break;
                try
                {
                    HpRow item = newHpeFormat
                        ? BuildNewHpeRow(sheet, row, receptionDate)
                        : BuildLegacyRow(sheet, row, receptionDate, hasSerial);
                    InsertSmartRow(con, item);
                }
                catch (SqlException ex) when (ex.Number == 2601 || ex.Number == 2627)
                {
                    WriteToFile("Duplicate ignored at Excel row " + (row + 1));
                }
                catch (Exception ex)
                {
                    throw new Exception("Error at Excel row " + (row + 1) + " : " + ex.Message, ex);
                }
            }
        }

        private static HpRow BuildLegacyRow(DataTable t, int r, string date, bool hasSerial)
        {
            return new HpRow
            {
                NumPo = C(t, r, "A"),
                HpOrderNo = C(t, r, "B"),
                RefFab = C(t, r, "C"),
                Status = C(t, r, "D"),
                ScheduleDate = Left(C(t, r, "E"), 10),
                ShipStatus = C(t, r, "F"),
                ShipmentNumber = C(t, r, "G"),
                DateReception = date,
                DatePoHp = Left(C(t, r, "H"), 10),
                PurchaseOrderDate = Left(C(t, r, "J"), 10),
                SchedShipDate = Left(C(t, r, "K"), 10),
                NetOrderPrice = DecimalValue(V(t, r, "L")),
                BundleId = C(t, r, "M"),
                DealId = C(t, r, "N"),
                SerialNumber = hasSerial ? C(t, r, "O") : null
            };
        }

        private static HpRow BuildNewHpeRow(DataTable t, int r, string date)
        {
            decimal quantity = DecimalValue(V(t, r, "AN")) ?? 0m;
            decimal total = DecimalValue(V(t, r, "EE")) ?? 0m;
            decimal unitPrice = quantity != 0 ? Math.Round(total / quantity, 2) : quantity;
            return new HpRow
            {
                NumPo = C(t, r, "B"),
                HpOrderNo = C(t, r, "A"),
                RefFab = C(t, r, "AJ"),
                Status = C(t, r, "J"),
                ScheduleDate = Left(C(t, r, "X"), 10),
                ShipStatus = C(t, r, "J"),
                ShipmentNumber = C(t, r, "BE"),
                DateReception = date,
                DatePoHp = Left(C(t, r, "V"), 10),
                PurchaseOrderDate = Left(C(t, r, "T"), 10),
                SchedShipDate = C(t, r, "X"),
                NetOrderPrice = unitPrice,
                BundleId = C(t, r, "AD"),
                DealId = C(t, r, "F"),
                SerialNumber = C(t, r, "BP")
            };
        }

        private void InsertSmartRow(SqlConnection con, HpRow r)
        {
            const string sql = @"INSERT INTO T_BO_HP_DEPUIS_SMART
(num_PO,HPOrderNO,ref_fab,status,ScheduleDate,ShipStatus,ShipmentNumber,date_reception,date_po_HP,PurchaseOrderDate,SchedShipDate,NetOrderPrice,BundleId,DealId,SerialNumber)
VALUES (@num_PO,@HPOrderNO,@ref_fab,@status,@ScheduleDate,@ShipStatus,@ShipmentNumber,@date_reception,@date_po_HP,@PurchaseOrderDate,@SchedShipDate,@NetOrderPrice,@BundleId,@DealId,@SerialNumber)";
            ExecuteNonQuery(con, sql,
                P("@num_PO", r.NumPo), P("@HPOrderNO", r.HpOrderNo), P("@ref_fab", r.RefFab), P("@status", r.Status),
                P("@ScheduleDate", r.ScheduleDate), P("@ShipStatus", r.ShipStatus), P("@ShipmentNumber", r.ShipmentNumber),
                P("@date_reception", r.DateReception), P("@date_po_HP", r.DatePoHp), P("@PurchaseOrderDate", r.PurchaseOrderDate),
                P("@SchedShipDate", r.SchedShipDate), P("@NetOrderPrice", r.NetOrderPrice), P("@BundleId", r.BundleId),
                P("@DealId", r.DealId), P("@SerialNumber", r.SerialNumber));
        }

        private void CopyPendingRowsToLicences(string receptionDate)
        {
            const string selectSql = @"SELECT num_PO,HPOrderNO,ref_fab,status,ScheduleDate,date_po_HP,date_reception,
PurchaseOrderDate,NetOrderPrice,BundleID,DealID,SerialNumber FROM T_BO_HP_DEPUIS_SMART
WHERE top_traite IS NULL AND date_reception=@date_reception";
            using (SqlConnection source = new SqlConnection(sql_connexion))
            using (SqlConnection target = new SqlConnection(sql_connexion_licences))
            {
                source.Open(); target.Open();
                using (SqlCommand cmd = new SqlCommand(selectSql, source))
                {
                    cmd.CommandTimeout = 300;
                    cmd.Parameters.Add(P("@date_reception", receptionDate));
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            string numPo = NormalizePo(Convert.ToString(dr["num_PO"]));
                            if (!string.IsNullOrEmpty(numPo)) InsertLicenceRow(target, dr, numPo);
                        }
                    }
                }
            }
        }

        private void InsertLicenceRow(SqlConnection con, SqlDataReader r, string numPo)
        {
            const string sql = @"INSERT INTO fichiers_HP_TOP_CONFIG
(date_maj,PurchaseOrderNo,PurchaseOrderDate,HPOrderNo,SchedDelvDate,SchedShipDate,Status,NetOrderPrice,ref_fab,BundleId,DealId,SerialNumber)
VALUES (@date_maj,@PurchaseOrderNo,@PurchaseOrderDate,@HPOrderNo,@SchedDelvDate,@SchedShipDate,@Status,@NetOrderPrice,@ref_fab,@BundleId,@DealId,@SerialNumber)";
            ExecuteNonQuery(con, sql,
                P("@date_maj", r["date_reception"]), P("@PurchaseOrderNo", numPo), P("@PurchaseOrderDate", r["PurchaseOrderDate"]),
                P("@HPOrderNo", r["HPOrderNO"]), P("@SchedDelvDate", r["ScheduleDate"]), P("@SchedShipDate", r["date_po_HP"]),
                P("@Status", r["status"]), P("@NetOrderPrice", r["NetOrderPrice"]), P("@ref_fab", r["ref_fab"]),
                P("@BundleId", r["BundleID"]), P("@DealId", r["DealID"]), P("@SerialNumber", r["SerialNumber"]));
        }

        private void MarkRowsAsProcessed(string receptionDate)
        {
            using (SqlConnection con = new SqlConnection(sql_connexion))
            {
                con.Open();
                ExecuteNonQuery(con, "UPDATE T_BO_HP_DEPUIS_SMART SET top_traite='O' WHERE top_traite IS NULL AND date_reception=@date",
                    P("@date", receptionDate));
            }
        }

        private void ExecuteNonQuery(SqlConnection con, string sql, params SqlParameter[] parameters)
        {
            try
            {
                using (SqlCommand cmd = new SqlCommand(sql, con))
                {
                    cmd.CommandTimeout = 300;
                    if (parameters != null) cmd.Parameters.AddRange(parameters);
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                SendSqlTechnicalIssueMail("ExecuteNonQuery", sql, ex);
                throw;
            }
        }

        private static SqlParameter P(string name, object value)
        {
            return new SqlParameter(name, value ?? DBNull.Value);
        }

        private static string NormalizePo(string value)
        {
            string p = (value ?? "").Trim();
            if (p.Length == 8 && p.EndsWith("D", StringComparison.OrdinalIgnoreCase)) return p.Substring(0, 2) + "-" + p.Substring(2, 5);
            if (p.Length >= 7 && p[6] == '/') return "21-" + p.Substring(0, 5).ToUpperInvariant();
            return p.Length == 5 ? "21-" + p : "";
        }

        private static object V(DataTable t, int row, string col)
        {
            int index = ColumnIndex(col);
            if (row < 0 || row >= t.Rows.Count || index < 0 || index >= t.Columns.Count) return null;
            object value = t.Rows[row][index];
            return value == DBNull.Value ? null : value;
        }

        private static string C(DataTable t, int row, string col) { return CellText(t, row, col); }
        private static string CellText(DataTable t, int row, string col)
        {
            object value = V(t, row, col);
            if (value == null) return "";
            if (value is DateTime) return ((DateTime)value).ToString("dd/MM/yyyy");
            return Convert.ToString(value).Trim();
        }

        private static int ColumnIndex(string col)
        {
            int result = 0;
            foreach (char c in col.ToUpperInvariant()) result = result * 26 + (c - 'A' + 1);
            return result - 1;
        }

        private static decimal? DecimalValue(object value)
        {
            if (value == null) return null;
            if (value is double) return Convert.ToDecimal(value);
            if (value is decimal) return (decimal)value;
            string text = Convert.ToString(value).Replace("EUR", "").Replace("€", "").Replace(" ", "").Trim();
            decimal result;
            if (decimal.TryParse(text, NumberStyles.Any, new CultureInfo("fr-FR"), out result)) return result;
            if (decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out result)) return result;
            return null;
        }

        private static string Left(string value, int len)
        {
            if (string.IsNullOrEmpty(value) || value.Length <= len) return value ?? "";
            return value.Substring(0, len);
        }

        private sealed class HpRow
        {
            public string NumPo, HpOrderNo, RefFab, Status, ScheduleDate, ShipStatus, ShipmentNumber;
            public string DateReception, DatePoHp, PurchaseOrderDate, SchedShipDate, BundleId, DealId, SerialNumber;
            public decimal? NetOrderPrice;
        }


        public Backorder_HP_FR()
        {





        }
    }


}
