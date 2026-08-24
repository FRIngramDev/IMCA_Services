using ClosedXML.Excel;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Brother_load_Bid
{
    public class Brother_load_Bid
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
        private string email_to = "";
        private string email_cc = "";    

        private string logs_folder = "";
        private string temp_folder = "";
        private string global_session_name = "";

        private string global_application_name = "BROTHER_LOAD_BID";

        private string sql_connexion = "";
        private string dss_con_openrowset = "";

        private string sql_connexion_parameter_global = "";
        private string dss_con_openrowset_parameter_global = "";

        private GraphServiceClient graphService = null;

        public Brother_load_Bid()
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
            public string email_to { get; set; } = "";
            public string email_cc { get; set; } = "";
            public string sql_connexion_parameter_global { get; set; } = "";
            public string dss_con_openrowset_parameter_global { get; set; } = "";

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
                    setEmailTo(p.email_to);
                    setEmailCc(p.email_cc);

                    setSqlConnexionParam(p.sql_connexion_parameter_global);
                    setDssConOpenrowsetParam(p.dss_con_openrowset_parameter_global);
                    setEmailInCaseOfTechnicalIssueParam(p.email_in_case_of_technical_issue_parameter_global);

                    email_in_case_of_technical_issue = get_IMCA_paramters(sql_con, email_in_case_of_technical_issue_parameter_global);
                    sql_connexion = get_IMCA_paramters(sql_con, sql_connexion_parameter_global);
                    dss_con_openrowset = get_IMCA_paramters(sql_con, dss_con_openrowset_parameter_global);

  

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
                            WriteToFile("   Temp folder              : " + temp_folder  );
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
                                            global_application_name,
                                            global_application_name + " - Une erreur (" + ex.Message + ") est survenue lors du traitement d'un fichier. Le mail a été déplacé dans le dossier Erreur.",
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

            if (string.IsNullOrWhiteSpace(dss_con_openrowset))
            {
                throw new Exception("dss_con_openrowset is empty. Parameter used : " + dss_con_openrowset_parameter_global);
            }
            if (string.IsNullOrWhiteSpace(email_in_case_of_technical_issue))
            {
                throw new Exception("email_in_case_of_technical_issue is empty. Parameter used : " + email_in_case_of_technical_issue_parameter_global);
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
            List<Recipient> ccRecipients = BuildRecipients(email_cc);

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

            if (ccRecipients.Count > 0)
            {
                message.CcRecipients = ccRecipients;
            }

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

        private class BrotherQuotationLine
        {
            public string MfrPartNumber { get; set; } = "";
            public string Description { get; set; } = "";
            public int QtyMin { get; set; } = 1;
            public int QtyTotal { get; set; } = 9999;
            public decimal PrixAchat { get; set; } = 0;
            public decimal PrixBase { get; set; } = 0;
        }

        private void ImportationFichierExcelEnBase(string filePath)
        {
            SqlConnection con = null;

            try
            {
                if (string.IsNullOrWhiteSpace(sql_connexion))
                {
                    throw new Exception("sql_connexion is empty");
                }

                if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                {
                    throw new Exception("Excel file not found : " + filePath);
                }

                RecalculateBrotherUsedQuantities();

                con = new SqlConnection(sql_connexion);
                con.Open();

                using (XLWorkbook wb = new XLWorkbook(filePath))
                {
                    IXLWorksheet ws = wb.Worksheet(1);

                    string nomContact = "";
                    string numCotaVendor = "";
                    string numVersion = "";
                    string nomRevendeur = "";
                    string nomEu = "";

                    DateTime creationDate;
                    DateTime dateFinValidite;

                    bool topCotationBrother = false;

                    int row = 2;

                    while (row <= 500)
                    {
                        string col1 = CellText(ws, row, 1);

                        if (col1.ToUpper().Contains("REVENDEUR"))
                        {
                            if (!ws.Cell(row, 2).IsEmpty())
                            {
                                nomRevendeur = CellText(ws, row, 2).ToUpper().Trim();
                            }

                            topCotationBrother = true;
                            break;
                        }

                        row++;
                    }

                    if (!topCotationBrother)
                    {
                        string message = global_application_name + " ignored because REVENDEUR was not found : " + filePath;
                        WriteToFile(message);
                        return;
                    }

                    row = 1;

                    while (row <= 500)
                    {
                        string col1 = CellText(ws, row, 1).ToUpper();

                        if (col1.Contains("N° COTATION") ||
                            col1.Contains("N°") ||
                            col1.Contains("REF. PROMO"))
                        {
                            numCotaVendor = CellText(ws, row, 2).ToUpper().Trim();

                            string versionCell = CellText(ws, row, 3).ToUpper().Trim();

                            if (!string.IsNullOrWhiteSpace(versionCell))
                            {
                                numVersion = versionCell.Replace("V", "").Replace(" ", "").Trim();
                            }

                            break;
                        }

                        row++;
                    }

                    if (string.IsNullOrWhiteSpace(numCotaVendor))
                    {
                        string message = global_application_name + " - Impossible de retrouver le n° de cotation dans le fichier Excel donné par Brother";
                        WriteToFile(message);
                        SendFunctionalBrotherMail(message);
                        throw new Exception(message);
                    }

                    while (row <= 500)
                    {
                        string col1 = CellText(ws, row, 1).ToUpper();

                        if (col1.StartsWith("CLIENT FINAL"))
                        {
                            nomEu = CellText(ws, row, 2).ToUpper().Trim();

                            if (string.IsNullOrWhiteSpace(nomEu))
                            {
                                nomEu = "PAS DE CLIENT FINAL";
                            }

                            break;
                        }

                        row++;
                    }

                    if (string.IsNullOrWhiteSpace(nomEu))
                    {
                        nomEu = "PAS DE CLIENT FINAL";
                    }

                    int existingIdCotation = GetExistingBrotherCotationId(con, numCotaVendor);

                    if (existingIdCotation > 0 && string.IsNullOrWhiteSpace(numVersion))
                    {
                        DateTime prolongationEndDate;

                        try
                        {
                            prolongationEndDate = FindBrotherEndDate(ws);
                        }
                        catch (Exception ex)
                        {
                            string message = global_application_name + " " + numCotaVendor + " - " + ex.Message;
                            WriteToFile(message);
                            SendFunctionalBrotherMail(message);
                            throw;
                        }

                        UpdateBrotherEndDate(con, numCotaVendor, prolongationEndDate);

                        string messageProlongation = global_application_name + " " + numCotaVendor + " - Prolongation de cette cotation";
                        WriteToFile(messageProlongation);
                        SendFunctionalBrotherMail(messageProlongation);

                        return;
                    }

                    List<BrotherQuotationLine> lines;

                    try
                    {
                        lines = ExtractBrotherQuotationLines(ws);
                    }
                    catch (Exception ex)
                    {
                        string message = global_application_name + " " + numCotaVendor + " - " + ex.Message;
                        WriteToFile(message);
                        SendFunctionalBrotherMail(message);
                        throw;
                    }

                    if (lines.Count == 0)
                    {
                        string message = global_application_name + " " + numCotaVendor + " - Impossible de retrouver les références produits";
                        WriteToFile(message);
                        SendFunctionalBrotherMail(message);
                        throw new Exception(message);
                    }

                    try
                    {
                        FindBrotherValidityDates(ws, out creationDate, out dateFinValidite);
                    }
                    catch (Exception ex)
                    {
                        string message = global_application_name + " " + numCotaVendor + " - " + ex.Message;
                        WriteToFile(message);
                        SendFunctionalBrotherMail(message);
                        throw;
                    }

                    if (string.IsNullOrWhiteSpace(numVersion))
                    {
                        numVersion = "1";
                    }
                    else
                    {
                        DeletePreviousBrotherVersion(con, numCotaVendor);
                    }

                    int idCotation = InsertBrotherHeader(
                        con,
                        numCotaVendor,
                        numVersion,
                        nomEu,
                        nomContact,
                        creationDate,
                        dateFinValidite
                    );

                    InsertBrotherLines(con, idCotation, numCotaVendor, lines);

                    UpdateBrotherSku(con, idCotation);

                    UpdateBrotherUsedQuantities(con, idCotation, numCotaVendor, creationDate, dateFinValidite);

                    UpdateBrotherRemainingQuantities(con, idCotation);

                    if (!string.IsNullOrWhiteSpace(nomRevendeur))
                    {
                        InsertAndMatchBrotherCustomer(con, idCotation, nomRevendeur);
                    }

                    bool hasMatchedCustomer = HasMatchedBrotherCustomer(con, idCotation);

                    ClearMatchedBrotherCustomerComment(con, idCotation);

                    if (!hasMatchedCustomer)
                    {
                        string message;

                        if (!string.IsNullOrWhiteSpace(nomRevendeur))
                        {
                            message = global_application_name + " " + numCotaVendor + " - non matchée automatiquement pour " + nomRevendeur;
                        }
                        else
                        {
                            message = "Cotation BROTHER " + numCotaVendor + " - non matchée automatiquement car pas de revendeur dans la cotation - Client Final " + nomEu;
                        }

                        WriteToFile(message);
                        SendFunctionalBrotherMail(message);
                    }
                    else
                    {
                        string message = "Cotation BROTHER " + numCotaVendor + " - chargé pour le client " + nomRevendeur;

                        WriteToFile(message);
                        SendFunctionalBrotherMail(message);
                    }
                }
            }
            finally
            {
                if (con != null)
                {
                    con.Close();
                    con.Dispose();
                }

                try
                {
                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath);
                    }
                }
                catch (Exception ex)
                {
                    WriteToFile("Unable to delete Excel file : " + filePath + " - " + ex.Message);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void RecalculateBrotherUsedQuantities()
        {
            using (SqlConnection con = new SqlConnection(sql_connexion))
            {
                con.Open();

                string selectSql = @"
            select idCotation,
                   numCotaVendor,
                   dateSelloutStart,
                   dateSelloutEnd
            from Cotation_header
            where groupName = 'BROTHER'
            and dateEntry is not null
            and dateSelloutEnd >= getdate()
            order by idCotation";

                using (SqlCommand cmd = new SqlCommand(selectSql, con))
                {
                    cmd.CommandTimeout = 300;

                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        List<Tuple<int, string, DateTime, DateTime>> quotations = new List<Tuple<int, string, DateTime, DateTime>>();

                        while (dr.Read())
                        {
                            quotations.Add(new Tuple<int, string, DateTime, DateTime>(
                                Convert.ToInt32(dr["idCotation"]),
                                dr["numCotaVendor"].ToString(),
                                Convert.ToDateTime(dr["dateSelloutStart"]),
                                Convert.ToDateTime(dr["dateSelloutEnd"])
                            ));
                        }

                        dr.Close();

                        foreach (Tuple<int, string, DateTime, DateTime> quotation in quotations)
                        {
                            UpdateBrotherUsedQuantities(con, quotation.Item1, quotation.Item2, quotation.Item3, quotation.Item4);
                            UpdateBrotherRemainingQuantities(con, quotation.Item1);
                        }
                    }
                }
            }
        }

        private int GetExistingBrotherCotationId(SqlConnection con, string numCotaVendor)
        {
            string sql = @"
        select isnull(max(idCotation), 0)
        from Cotation_header
        where numCotaVendor = @numCotaVendor
        and groupName = 'BROTHER'";

            using (SqlCommand cmd = new SqlCommand(sql, con))
            {
                cmd.CommandTimeout = 300;
                cmd.Parameters.AddWithValue("@numCotaVendor", numCotaVendor.ToUpper().Trim());

                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        private void UpdateBrotherEndDate(SqlConnection con, string numCotaVendor, DateTime dateFinValidite)
        {
            string sql = @"
        update Cotation_header
        set dateSelloutEnd = @dateSelloutEnd
        where numCotaVendor = @numCotaVendor
        and groupName = 'BROTHER'";

            using (SqlCommand cmd = new SqlCommand(sql, con))
            {
                cmd.CommandTimeout = 300;
                cmd.Parameters.AddWithValue("@dateSelloutEnd", dateFinValidite);
                cmd.Parameters.AddWithValue("@numCotaVendor", numCotaVendor.ToUpper().Trim());
                cmd.ExecuteNonQuery();
            }
        }

        private List<BrotherQuotationLine> ExtractBrotherQuotationLines(IXLWorksheet ws)
        {
            List<BrotherQuotationLine> lines = new List<BrotherQuotationLine>();

            int row = 10;
            bool headerFound = false;

            // Recherche de l'entête "Référence SAP"
            while (row <= 500)
            {
                string col3 = CellText(ws, row, 3).ToUpper().Trim();

                if (col3.Contains("RÉFÉRENCE SAP") || col3.Contains("REFERENCE SAP"))
                {
                    headerFound = true;
                    row++;
                    break;
                }

                row++;
            }

            if (!headerFound)
            {
                throw new Exception("Format incorrect. Entête de colonne REFERENCE SAP non trouvée");
            }

            // Lecture des lignes produits
            while (row <= 500)
            {
                string col1 = CellText(ws, row, 1).ToUpper().Trim();

                if (col1.Contains("OFFRES SOUMISES") ||
                    col1.Contains("VALIDITÉ DES OFFRES") ||
                    col1.Contains("VALIDITE DES OFFRES") ||
                    col1.Contains("COMMENTAIRES"))
                {
                    break;
                }

                string referenceSap = CellText(ws, row, 3).ToUpper().Trim();

                if (!string.IsNullOrWhiteSpace(referenceSap))
                {
                    decimal prixBase = 0;

                    // Cas standard : prix en colonne 4
                    if (!TryGetDecimal(ws, row, 4, out prixBase))
                    {
                        if (debug.ToUpper() == "TRUE")
                        {
                            WriteToFile(
                                "       Ligne " + row +
                                " : prix de vente invalide/vide en colonne 4. Valeur=[" + CellText(ws, row, 4) +
                                "], tentative colonne 5. Référence=[" + referenceSap + "]"
                            );
                        }

                        // Cas spécifique : modèle Brother avec colonne supplémentaire, prix en colonne 5
                        if (!TryGetDecimal(ws, row, 5, out prixBase))
                        {
                            string col4Value = CellText(ws, row, 4);
                            string col5Value = CellText(ws, row, 5);

                            // Même comportement que l'ancien programme : si prix vide, on force à 0
                            if (string.IsNullOrWhiteSpace(col4Value) && string.IsNullOrWhiteSpace(col5Value))
                            {
                                prixBase = 0;

                                if (debug.ToUpper() == "TRUE")
                                {
                                    WriteToFile(
                                        "       Ligne " + row +
                                        " : prix de vente vide en colonnes 4 et 5. PrixBase forcé à 0. Référence=[" + referenceSap + "]"
                                    );
                                }
                            }
                            else
                            {
                                if (debug.ToUpper() == "TRUE")
                                {
                                    WriteToFile(
                                        "       Ligne ignorée " + row +
                                        " : prix de vente invalide en colonnes 4 et 5. Col4=[" + col4Value +
                                        "], Col5=[" + col5Value +
                                        "], Référence=[" + referenceSap + "]"
                                    );
                                }

                                row++;
                                continue;
                            }
                        }
                    }

                    decimal prixAchat = 0;

                    if (!TryGetDecimal(ws, row, 6, out prixAchat))
                    {
                        if (debug.ToUpper() == "TRUE")
                        {
                            WriteToFile(
                                "       Ligne " + row +
                                " : prix achat invalide en colonne 6. Valeur=[" + CellText(ws, row, 6) +
                                "]. PrixAchat forcé à 0. Référence=[" + referenceSap + "]"
                            );
                        }
                    }

                    int qtyTotal = 9999;

                    string qtyText = CellText(ws, row, 1)
                        .ToUpper()
                        .Replace("PCE", "")
                        .Replace("PCS", "")
                        .Trim();

                    int qtyCell;
                    if (int.TryParse(qtyText, out qtyCell))
                    {
                        qtyTotal = qtyCell;
                    }

                    BrotherQuotationLine line = new BrotherQuotationLine
                    {
                        MfrPartNumber = Left(referenceSap.Replace("'", "''"), 50),
                        Description = "",
                        QtyMin = 1,
                        QtyTotal = qtyTotal,
                        PrixAchat = Math.Round(prixAchat, 2),
                        PrixBase = Math.Round(prixBase, 4)
                    };

                    lines.Add(line);

                    if (debug.ToUpper() == "TRUE")
                    {
                        WriteToFile(
                            "       Ligne ajoutée " + row +
                            " : ref=[" + line.MfrPartNumber +
                            "], qty=[" + line.QtyTotal +
                            "], prixAchat=[" + line.PrixAchat +
                            "], prixBase=[" + line.PrixBase + "]"
                        );
                    }
                }

                row++;
            }

            return lines;
        }

        private void FindBrotherValidityDates(IXLWorksheet ws, out DateTime creationDate, out DateTime dateFinValidite)
        {
            int row = 17;

            while (row <= 5000)
            {
                string col1 = CellText(ws, row, 1).ToUpper();

                if (col1.Contains("DATE DE DÉBUT :") ||
                    col1.Contains("DATE DE DÉBUT:") ||
                    col1.Contains("DATE DE DEBUT :") ||
                    col1.Contains("DATE DE DEBUT:"))
                {
                    if (!TryGetDate(ws, row, 2, out creationDate))
                    {
                        throw new Exception("La date de début de validité est inconnue ou invalide");
                    }

                    if (!TryGetDate(ws, row + 1, 2, out dateFinValidite))
                    {
                        throw new Exception("La date de fin de validité est inconnue ou invalide");
                    }

                    return;
                }

                row++;
            }

            throw new Exception("Impossible de retrouver les dates de validité");
        }

        private DateTime FindBrotherEndDate(IXLWorksheet ws)
        {
            DateTime creationDate;
            DateTime dateFinValidite;

            FindBrotherValidityDates(ws, out creationDate, out dateFinValidite);

            return dateFinValidite;
        }

        private void DeletePreviousBrotherVersion(SqlConnection con, string numCotaVendor)
        {
            ExecuteNonQuery(con, @"
        delete from cotation_customer
        where idCotation in (
            select idCotation
            from Cotation_header
            where numCotaVendor = @numCotaVendor
            and groupName = 'BROTHER'
        )",
                new SqlParameter("@numCotaVendor", numCotaVendor));

            ExecuteNonQuery(con, @"
        delete from cotation_line
        where idCotation in (
            select idCotation
            from Cotation_header
            where numCotaVendor = @numCotaVendor
            and groupName = 'BROTHER'
        )",
                new SqlParameter("@numCotaVendor", numCotaVendor));

            ExecuteNonQuery(con, @"
        delete from Cotation_header
        where numCotaVendor = @numCotaVendor
        and groupName = 'BROTHER'",
                new SqlParameter("@numCotaVendor", numCotaVendor));
        }

        private int InsertBrotherHeader(
            SqlConnection con,
            string numCotaVendor,
            string numVersion,
            string nomEu,
            string nomContact,
            DateTime creationDate,
            DateTime dateFinValidite)
        {
            string insertSql = @"
        insert into Cotation_header
        (
            numCotaVendor,
            groupName,
            version,
            eu,
            [type],
            dateSellOutStart,
            DateSellOutEnd,
            countryCode,
            contact,
            dateEntry,
            versionTXT
        )
        values
        (
            @numCotaVendor,
            'BROTHER',
            @version,
            @eu,
            'rebate',
            @dateSellOutStart,
            @dateSellOutEnd,
            'FR',
            @contact,
            getdate(),
            @versionTXT
        )";

            using (SqlCommand cmd = new SqlCommand(insertSql, con))
            {
                cmd.CommandTimeout = 300 ;
                cmd.Parameters.AddWithValue("@numCotaVendor", numCotaVendor);
                cmd.Parameters.AddWithValue("@version", Convert.ToInt32(numVersion));
                cmd.Parameters.AddWithValue("@eu", nomEu ?? "");
                cmd.Parameters.AddWithValue("@dateSellOutStart", creationDate);
                cmd.Parameters.AddWithValue("@dateSellOutEnd", dateFinValidite);
                cmd.Parameters.AddWithValue("@contact", nomContact ?? "");
                cmd.Parameters.AddWithValue("@versionTXT", numVersion);
                cmd.ExecuteNonQuery();
            }

            string selectSql = @"
        select isnull(max(idCotation), 0)
        from Cotation_header
        where numCotaVendor = @numCotaVendor
        and groupName = 'BROTHER'";

            using (SqlCommand cmd = new SqlCommand(selectSql, con))
            {
                cmd.CommandTimeout = 300;
                cmd.Parameters.AddWithValue("@numCotaVendor", numCotaVendor);

                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        private void InsertBrotherLines(SqlConnection con, int idCotation, string numCotaVendor, List<BrotherQuotationLine> lines)
        {
            int lineNumber = 1;

            foreach (BrotherQuotationLine line in lines)
            {
                int existingLine = Convert.ToInt32(ExecuteScalar(con, @"
            select count(*)
            from Cotation_line
            where idCotation = @idCotation
            and mfr_prt_nbr = @mfr_prt_nbr",
                    new SqlParameter("@idCotation", idCotation),
                    new SqlParameter("@mfr_prt_nbr", line.MfrPartNumber)));

                if (existingLine == 0)
                {
                    string sql = @"
                insert into Cotation_line
                (
                    numCotaVendor,
                    idCotation,
                    ligne,
                    [type],
                    mfr_prt_nbr,
                    description,
                    qtyMin,
                    qtyTotal,
                    prixAchat,
                    currency,
                    prixBase,
                    RemainingQtyIMFR
                )
                values
                (
                    @numCotaVendor,
                    @idCotation,
                    @ligne,
                    'PN',
                    @mfr_prt_nbr,
                    @description,
                    @qtyMin,
                    @qtyTotal,
                    @prixAchat,
                    'EUR',
                    @prixBase,
                    @remainingQty
                )";

                    ExecuteNonQuery(con, sql,
                        new SqlParameter("@numCotaVendor", numCotaVendor),
                        new SqlParameter("@idCotation", idCotation),
                        new SqlParameter("@ligne", lineNumber),
                        new SqlParameter("@mfr_prt_nbr", line.MfrPartNumber),
                        new SqlParameter("@description", line.Description),
                        new SqlParameter("@qtyMin", line.QtyMin),
                        new SqlParameter("@qtyTotal", line.QtyTotal),
                        new SqlParameter("@prixAchat", line.PrixAchat),
                        new SqlParameter("@prixBase", line.PrixBase),
                        new SqlParameter("@remainingQty", line.QtyTotal));
                }

                lineNumber++;
            }
        }

        private void UpdateBrotherSku(SqlConnection con, int idCotation)
        {
            string sql = @"
        update Cotation_line
        set sku = (
            select sku
            from OPENROWSET('SQLOLEDB', " + dss_con_openrowset + @", 
            'SELECT product.mfr_part_nbr, product.sku
              FROM product
              INNER JOIN (
                    SELECT product_1.mfr_part_nbr
                    FROM product AS product_1
                    INNER JOIN vendor ON product_1.vendor_nbr = vendor.vendor_nbr
                    WHERE vendor.GroupNam LIKE ''%BROTHER%''
                    AND (product_class + product_status + product_type <> ''D*I'')
                    AND (CRC_CODE = ''STD'' OR CRC_CODE = '''')
                    GROUP BY product_1.mfr_part_nbr
                    HAVING COUNT(*) = 1
              ) AS product_bis ON product.mfr_part_nbr = product_bis.mfr_part_nbr
              INNER JOIN vendor AS vendor_1 ON vendor_1.vendor_nbr = product.vendor_nbr
              WHERE vendor_1.GroupNam LIKE ''%BROTHER%''
              AND (product_class + product_status + product_type <> ''D*I'')
              AND (CRC_CODE = ''STD'' OR CRC_CODE = '''')
              ORDER BY product.mfr_part_nbr') dss
            where dss.mfr_part_nbr = Cotation_line.mfr_prt_nbr
        )
        where idCotation = @idCotation";

            ExecuteNonQuery(con, sql, new SqlParameter("@idCotation", idCotation));
        }

        private void UpdateBrotherUsedQuantities(SqlConnection con, int idCotation, string numCotaVendor, DateTime startDate, DateTime endDate)
        {
            string sql = @"
        update Cotation_line
        set QtyUtilisee = isnull((
            select sum(qty) as nb
            from dbo.meet_comp_header
            inner join dbo.Meet_comp_line
                on dbo.meet_comp_header.Number = dbo.Meet_comp_line.[Meet comp number]
            where dbo.meet_comp_header.[Cotation number] = @numCotaVendor
            and dbo.Meet_comp_line.VPN = Cotation_line.mfr_prt_nbr
            and CANCEL = 'Non'
            and dbo.meet_comp_header.Date between @startDate and @endDate
        ), 0)
        where idCotation = @idCotation";

            ExecuteNonQuery(con, sql,
                new SqlParameter("@numCotaVendor", numCotaVendor),
                new SqlParameter("@startDate", startDate),
                new SqlParameter("@endDate", endDate),
                new SqlParameter("@idCotation", idCotation));
        }

        private void UpdateBrotherRemainingQuantities(SqlConnection con, int idCotation)
        {
            ExecuteNonQuery(con,
                "update Cotation_line set RemainingQtyIMFR = qtyTotal - QtyUtilisee where idCotation = @idCotation",
                new SqlParameter("@idCotation", idCotation));
        }

        private void InsertAndMatchBrotherCustomer(SqlConnection con, int idCotation, string nomRevendeur)
        {
            ExecuteNonQuery(con, @"
        insert into Cotation_customer
        (
            idCotation,
            customerName,
            Comment1
        )
        values
        (
            @idCotation,
            @customerName,
            'Client non trouve'
        )",
                new SqlParameter("@idCotation", idCotation),
                new SqlParameter("@customerName", nomRevendeur));

            string escapedCustomer = nomRevendeur.Replace("'", "''''").ToUpper().Trim();

            string sqlDss = @"
        update Cotation_customer
        set customerNumber = (
            select branch_customer_nbr
            from OPENROWSET('SQLOLEDB', " + dss_con_openrowset + @",
            'select TOP 1 branch_customer_nbr
              from customer
              where branch_nbr in (''21'', ''24'', ''15'')
              and statusCustFlg <> ''D''
              and cust_name = ''" + escapedCustomer + @"''
              order by credit_limit desc')
        )
        where idCotation = @idCotation";

            ExecuteNonQuery(con, sqlDss, new SqlParameter("@idCotation", idCotation));

            ExecuteNonQuery(con, @"
        update Cotation_customer
        set customerNumber = (
            select customer_nbr
            from Matchage_Client_BROTHER
            where customer_nbr is not null
            and Cotation_customer.customerName = Matchage_Client_BROTHER.nom_client_BROTHER
        )
        where idCotation = @idCotation
        and customerNumber is null",
                new SqlParameter("@idCotation", idCotation));

            ExecuteNonQuery(con, @"
        insert into Matchage_Client_BROTHER
        select resel, null
        from (
            select distinct rtrim(upper(customerName)) as resel
            from Cotation_customer
            where customerNumber is null
            and idCotation = @idCotation
        ) as the_bid
        left join (
            select distinct rtrim(upper(nom_client_BROTHER)) as res_brother
            from Matchage_Client_BROTHER
        ) as the_brother
            on resel = res_brother
        where res_brother is null",
                new SqlParameter("@idCotation", idCotation));
        }

        private bool HasMatchedBrotherCustomer(SqlConnection con, int idCotation)
        {
            object result = ExecuteScalar(con, @"
        select count(*)
        from Cotation_customer
        where idCotation = @idCotation
        and customerNumber is not null",
                new SqlParameter("@idCotation", idCotation));

            return Convert.ToInt32(result) > 0;
        }

        private void ClearMatchedBrotherCustomerComment(SqlConnection con, int idCotation)
        {
            ExecuteNonQuery(con, @"
        update Cotation_customer
        set Comment1 = ''
        where customerNumber is not null
        and idCotation = @idCotation",
                new SqlParameter("@idCotation", idCotation));
        }

        private string CellText(IXLWorksheet ws, int row, int col)
        {
            try
            {
                return ws.Cell(row, col).GetFormattedString().Trim();
            }
            catch
            {
                return "";
            }
        }

        private bool TryGetDecimal(IXLWorksheet ws, int row, int col, out decimal value)
        {
            value = 0;

            try
            {
                IXLCell cell = ws.Cell(row, col);

                try
                {
                    double doubleValue = cell.GetDouble();
                    value = Convert.ToDecimal(doubleValue);
                    return true;
                }
                catch
                {
                    // Pas un nombre natif Excel.
                }

                string text = CellText(ws, row, col);

                if (string.IsNullOrWhiteSpace(text))
                {
                    return false;
                }

                text = text
                    .Replace("€", "")
                    .Replace("EUR", "")
                    .Replace("HT", "")
                    .Replace("\u00A0", "")
                    .Replace(" ", "")
                    .Trim();

                if (string.IsNullOrWhiteSpace(text))
                {
                    return false;
                }

                CultureInfo fr = new CultureInfo("fr-FR");

                if (decimal.TryParse(text, NumberStyles.Any, fr, out value))
                {
                    return true;
                }

                if (decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out value))
                {
                    return true;
                }

                string invariantText = text.Replace(",", ".");

                if (decimal.TryParse(invariantText, NumberStyles.Any, CultureInfo.InvariantCulture, out value))
                {
                    return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        private bool TryGetDate(IXLWorksheet ws, int row, int col, out DateTime value)
        {
            value = DateTime.MinValue;

            try
            {
                IXLCell cell = ws.Cell(row, col);

                try
                {
                    value = cell.GetDateTime();
                    return true;
                }
                catch
                {
                    // La cellule n'est pas directement lisible comme DateTime.
                }

                string text = CellText(ws, row, col).Trim();

                if (string.IsNullOrWhiteSpace(text))
                {
                    return false;
                }

                if (text.Length > 10)
                {
                    text = text.Substring(0, 10);
                }

                text = text.Replace(".", "/");

                CultureInfo fr = new CultureInfo("fr-FR");

                if (DateTime.TryParse(text, fr, DateTimeStyles.None, out value))
                {
                    return true;
                }

                if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out value))
                {
                    return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        private string Left(string value, int length)
        {
            if (string.IsNullOrEmpty(value))
            {
                return "";
            }

            if (value.Length <= length)
            {
                return value;
            }

            return value.Substring(0, length);
        }

        private object ExecuteScalar(SqlConnection con, string sql, params SqlParameter[] parameters)
        {
            try
            {
                using (SqlCommand cmd = new SqlCommand(sql, con))
                {
                    cmd.CommandTimeout = 300;

                    if (parameters != null)
                    {
                        cmd.Parameters.AddRange(parameters);
                    }

                    return cmd.ExecuteScalar();
                }
            }
            catch (Exception ex)
            {
                SendSqlTechnicalIssueMail("ExecuteScalar", sql, ex);
                throw;
            }
        }

        private void ExecuteNonQuery(SqlConnection con, string sql, params SqlParameter[] parameters)
        {
            try
            {
                using (SqlCommand cmd = new SqlCommand(sql, con))
                {
                    cmd.CommandTimeout = 300;

                    if (parameters != null)
                    {
                        cmd.Parameters.AddRange(parameters);
                    }

                    cmd.ExecuteNonQuery();
                }
            }
             catch (Exception ex)
            {
                SendSqlTechnicalIssueMail("ExecuteScalar", sql, ex);
                throw;
            }
        }

        private void SendFunctionalBrotherMail(string message)
        {
            try
            {
                string recipient = !string.IsNullOrWhiteSpace(email_to)
                    ? email_to
                    : email_in_case_of_technical_issue;

                if (string.IsNullOrWhiteSpace(recipient) || !recipient.Contains("@"))
                {
                    WriteToFile("Functional Brother email not sent because recipient is empty. Message : " + message);
                    return;
                }

                if (graphService == null)
                {
                    WriteToFile("Functional Brother email not sent because graphService is null. Message : " + message);
                    return;
                }

                EnvoiEmail_with_Graph(
                    graphService,
                    "Importation des cotations Brother",
                    message,
                    recipient
                );
            }
            catch (Exception ex)
            {
                WriteToFile("Error sending functional Brother email : " + ex.Message);
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

                string recipient = !string.IsNullOrWhiteSpace(email_in_case_of_technical_issue)
                    ? email_in_case_of_technical_issue
                    : email_to;

                if (string.IsNullOrWhiteSpace(recipient) || !recipient.Contains("@"))
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
                    "<pre>" + SanitizeSqlForMail(sql) + "</pre>";

                EnvoiEmail_with_Graph(
                    graphService,
                    subject,
                    body,
                    recipient
                );
            }
            catch (Exception mailEx)
            {
                WriteToFile("Error sending SQL technical issue email : " + mailEx.Message);
            }
        }

        private string SanitizeSqlForMail(string sql)
        {
            if (string.IsNullOrWhiteSpace(sql))
            {
                return "";
            }

            string sanitizedSql = sql;

            if (!string.IsNullOrWhiteSpace(dss_con_openrowset))
            {
                sanitizedSql = sanitizedSql.Replace(dss_con_openrowset, "***DSS_CON_OPENROWSET***");
            }

            if (sanitizedSql.Length > 3000)
            {
                sanitizedSql = sanitizedSql.Substring(0, 3000) + "...";
            }

            return sanitizedSql;
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

        public void setEmailTo(string email_to)
        {
            this.email_to = email_to ?? "";
        }

        public void setEmailCc(string email_cc)
        {
            this.email_cc = email_cc ?? "";
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

        public void setDssConOpenrowset(string dss_con_openrowset)
        {
            this.dss_con_openrowset = dss_con_openrowset ?? "";
        }

        public void setSqlConnexionParam(string sql_connexion_parameter_global)
        {
            this.sql_connexion_parameter_global = sql_connexion_parameter_global ?? "";
        }

        public void setDssConOpenrowsetParam(string dss_con_openrowset_parameter_global)
        {
            this.dss_con_openrowset_parameter_global = dss_con_openrowset_parameter_global ?? "";
        }

        public void setEmailInCaseOfTechnicalIssueParam(string email_in_case_of_technical_issue_parameter_global)
        {
            this.email_in_case_of_technical_issue_parameter_global = email_in_case_of_technical_issue_parameter_global ?? "";
        }

    }
}
