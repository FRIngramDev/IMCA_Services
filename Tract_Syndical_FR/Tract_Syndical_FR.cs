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
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using System.DirectoryServices.AccountManagement;

namespace TRACT_SYNDICAL_FR
{
    public class TRACT_SYNDICAL_FR
    {
        public TRACT_SYNDICAL_FR()
        {
        }

        // ----- Current country / processing parameters -----
        private string country = "";
        private string sk_valid = "";
        private string name = "";
        private string active = "";
        private string debug = "";
        private string start_date_scan = "";
        private string number_of_mails = "10";

        // ----- Current shared mailbox parameters -----
        private string sharedmailbox_name = "";
        private string nom_mailboxe = "";
        private string mailboxe_reponse = "";
        private string expediteur_autorise = "";
        private string sharedmailbox_folder_in = "";
        private string sharedmailbox_folder_out = "";
        private int id_mailboxe = 0;

        // ----- Technical alert / recipient parameters -----
        private string email_in_case_of_technical_issue_parameter_global = "";
        private string email_in_case_of_technical_issue = "";
        private string email_destinataire = "";

        // ----- Files / logs / session -----
        private string logs_folder = "";
        private string temp_folder = "";
        private string global_session_name = "";

        // ----- Application name -----
        private string global_application_name = "TRACT_SYNDICAL_FR";

        // ----- SQL connection parameters -----
        private string sql_connexion = "";
        private string sql_connexion_parameter_global = "";

        // ----- Microsoft Graph client -----
        private GraphServiceClient graphService = null;

        // ----- Microsoft Graph attachment limits -----
        private const int simple_attachment_limit_bytes = 3 * 1024 * 1024;
        private const long maximum_attachment_size_bytes = 150L * 1024L * 1024L;
        private const int upload_slice_size_bytes = 10 * 320 * 1024;

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
            public string sharedmailbox_folder_in { get; set; } = "INBOX";
            public string sharedmailbox_folder_out { get; set; } = "Archives";
            public string email_in_case_of_technical_issue_parameter_global { get; set; } = "";
            public string email_destinataire { get; set; } = "";
            public string sql_connexion_parameter_global { get; set; } = "";
        }

        private class TractAttachment
        {
            public string Name { get; set; } = "";
            public string ContentType { get; set; } = "application/octet-stream";
            public string ContentId { get; set; } = "";
            public bool IsInline { get; set; } = false;
            public byte[] ContentBytes { get; set; } = new byte[0];
        }

        private sealed class SharedMailboxConfiguration
        {
            public int IdMailboxe { get; set; }
            public string NomMailboxe { get; set; } = "";
            public string SharedMailboxName { get; set; } = "";
            public string MailboxeReponse { get; set; } = "";
            public string ExpediteurAutorise { get; set; } = "";
            public int Ordre { get; set; }
        }


        /// <summary>
        /// Indicates that the technical alert associated with an exception
        /// has already been sent, preventing duplicate alert emails.
        /// </summary>
        private sealed class TechnicalAlertAlreadySentException : Exception
        {
            public TechnicalAlertAlreadySentException(
                string message,
                Exception innerException)
                : base(message, innerException)
            {
            }
        }

        /// <summary>
        /// Creates and returns a Microsoft Graph client.
        /// </summary>
        private GraphServiceClient Connexion_Microsoft_Graph()
        {
            System.Net.ServicePointManager.SecurityProtocol =
                System.Net.SecurityProtocolType.Tls12;

            GraphServiceClient GraphService;
            class_dev_tools.Ews_Modern_Auth Ews_Modern_Auth =
                new class_dev_tools.Ews_Modern_Auth();

            GraphService = Ews_Modern_Auth.Get_Graph_Service();
            return GraphService;
        }

        /// <summary>
        /// Main entry point used by IMCA to read emails from every configured shared mailbox.
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
                    setEmailInCaseOfTechnicalIssueParam(
                        p.email_in_case_of_technical_issue_parameter_global);

                    email_in_case_of_technical_issue = get_IMCA_paramters(
                        sql_con,
                        email_in_case_of_technical_issue_parameter_global);

                    sql_connexion = get_IMCA_paramters(
                        sql_con,
                        sql_connexion_parameter_global);

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

                        List<SharedMailboxConfiguration> sharedMailboxes =
                            GetActiveSharedMailboxes();

                        if (sharedMailboxes.Count == 0)
                        {
                            WriteToFile(
                                "   No active shared mailbox found in " +
                                "T_TRACT_SYNDICAL_Mailboxes");
                            continue;
                        }

                        List<Exception> mailboxErrors = new List<Exception>();

                        foreach (SharedMailboxConfiguration mailbox in sharedMailboxes)
                        {
                            setSharedMailboxName(mailbox.SharedMailboxName);
                            setNomMailboxe(mailbox.NomMailboxe);
                            setMailboxeReponse(mailbox.MailboxeReponse);
                            setExpediteurAutorise(mailbox.ExpediteurAutorise);
                            id_mailboxe = mailbox.IdMailboxe;

                            try
                            {
                                if (debug.ToUpper().Trim() == "TRUE")
                                {
                                    WriteToFile(
                                        "   Processing mailbox " +
                                        nom_mailboxe + " : " +
                                        sharedmailbox_name +
                                        " - Order : " +
                                        mailbox.Ordre);
                                }

                                ReadCurrentSharedMailbox();
                            }
                            catch (Exception mailboxException)
                            {
                                WriteToFile(
                                    "   Error reading mailbox " +
                                    sharedmailbox_name + " : " +
                                    mailboxException.Message);

                                mailboxErrors.Add(
                                    new Exception(
                                        "Mailbox " +
                                        sharedmailbox_name +
                                        " : " +
                                        mailboxException.Message,
                                        mailboxException));

                                if (!(mailboxException is
                                    TechnicalAlertAlreadySentException))
                                {
                                    SendTechnicalIssueMail(
                                        nameof(Read_Email_with_Graph),
                                        "",
                                        sharedmailbox_name + " - " +
                                        mailboxException.Message,
                                        "MAILBOX PROCESSING");
                                }
                            }
                        }

                        if (mailboxErrors.Count > 0)
                        {
                            throw new AggregateException(
                                mailboxErrors.Count +
                                " shared mailbox(es) could not be processed.",
                                mailboxErrors);
                        }
                    }
                    catch (Exception e)
                    {
                        WriteToFile("   Error get emails : " + e.Message);
                        throw;
                    }
                }
            }
            catch (Exception ex)
            {
                WriteToFile(
                    "Global error Read_Email_with_Graph : " + ex.Message);
                throw;
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }


        /// <summary>
        /// Synchronise les utilisateurs de l'Active Directory France avec les abonnements
        /// aux tracts syndicaux. Cette methode est independante et n'est jamais appelee
        /// par Read_Email_with_Graph. Elle est destinee a etre appelee directement par IMCA.
        /// </summary>
        public void Recherche_AD_FR_Pour_Tracts_Syndicaux(string sql_con,string logs,string tmp_folder,string session_name)
        {
            const string domainController = "DEFRIZWADC1001";
            const string organizationalUnit =
                "OU=FR,OU=EMEA,OU=associates,OU=usersAndGroups," +
                "DC=corporate,DC=ingrammicro,DC=com";

            const string insertSql = @"
                                    INSERT INTO T_TRACT_SYNDICAL_Abonnes
                                        (id_mailboxe, id_user, date_abonnement, actif)
                                    SELECT
                                        m.id_mailboxe,
                                        u.id_user,
                                        GETDATE(),
                                        'O'
                                    FROM T_TRACT_SYNDICAL_Mailboxes m
                                    INNER JOIN annuaire.dbo.tbl_users u ON 1 = 1
                                    LEFT JOIN T_TRACT_SYNDICAL_Abonnes a
                                        ON a.id_user = u.id_user
                                       AND a.id_mailboxe = m.id_mailboxe
                                    WHERE m.actif = 1
                                      AND u.matPeopleSoft = @EMPLOYEE_ID
                                      AND u.user_email IS NOT NULL
                                      AND LTRIM(RTRIM(u.user_email)) <> ''
                                      AND a.id_user IS NULL;";

            string service_path = Path.GetDirectoryName(
                System.Reflection.Assembly.GetEntryAssembly().Location);

            setlogs_folder(service_path + "\\" + logs);
            settemp_folder(service_path + "\\" + tmp_folder);
            global_session_name = session_name;

            int usersRead = 0;
            int validUsers = 0;
            int subscriptionsInserted = 0;

            try
            {
                string global_parameters = get_IMCA_paramters(
                    sql_con,
                    global_application_name);

                if (string.IsNullOrWhiteSpace(global_parameters))
                {
                    throw new InvalidOperationException(
                        "No parameters found for " + global_application_name);
                }

                JSON_file param = JsonConvert.DeserializeObject<JSON_file>(
                    global_parameters);

                Country france = param?.countries?
                    .FirstOrDefault(item =>
                        string.Equals(
                            item.country?.Trim(),
                            "FR",
                            StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(
                            item.active?.Trim(),
                            "TRUE",
                            StringComparison.OrdinalIgnoreCase));

                if (france == null)
                {
                    throw new InvalidOperationException(
                        "No active FR configuration found for " +
                        global_application_name);
                }

                setParamCountry(france.country);
                setParamSK_Valid(france.sk_valid);
                setParamName(france.name);
                setParamActive(france.active);
                setParamDebug(france.debug);
                setSqlConnexionParam(france.sql_connexion_parameter_global);
                setEmailInCaseOfTechnicalIssueParam(
                    france.email_in_case_of_technical_issue_parameter_global);
                setEmailDestinataire(france.email_destinataire);

                sql_connexion = get_IMCA_paramters(
                    sql_con,
                    sql_connexion_parameter_global);

                email_in_case_of_technical_issue = get_IMCA_paramters(
                    sql_con,
                    email_in_case_of_technical_issue_parameter_global);

                if (string.IsNullOrWhiteSpace(sql_connexion))
                {
                    throw new InvalidOperationException(
                        "sql_connexion is empty. Parameter used : " +
                        sql_connexion_parameter_global);
                }

                // Initialise Graph uniquement pour pouvoir envoyer une alerte technique.
                graphService = Connexion_Microsoft_Graph();

                WriteToFile(
                    "Starting French AD tract subscription synchronization");

                using (PrincipalContext context = new PrincipalContext(
                    ContextType.Domain,
                    domainController,
                    organizationalUnit))
                using (UserPrincipal searchTemplate = new UserPrincipal(context))
                using (PrincipalSearcher searcher = new PrincipalSearcher(
                    searchTemplate))
                using (PrincipalSearchResult<Principal> results = searcher.FindAll())
                using (SqlConnection connection = new SqlConnection(sql_connexion))
                using (SqlCommand command = new SqlCommand(insertSql, connection))
                {
                    command.CommandTimeout = 300;

                    SqlParameter employeeIdParameter = command.Parameters.Add(
                        "@EMPLOYEE_ID",
                        SqlDbType.NVarChar,
                        6);

                    connection.Open();

                    foreach (Principal principal in results)
                    {
                        usersRead++;

                        UserPrincipal user = principal as UserPrincipal;
                        if (user == null)
                        {
                            continue;
                        }

                        string employeeId = (user.EmployeeId ?? "").Trim();

                        if (employeeId.Length != 6 ||
                            employeeId == "000000")
                        {
                            continue;
                        }

                        validUsers++;
                        employeeIdParameter.Value = employeeId;
                        subscriptionsInserted += command.ExecuteNonQuery();
                    }
                }

                WriteToFile(
                    "French AD tract subscription synchronization completed");

                WriteToFile(
                    "   Users read               : " + usersRead);

                WriteToFile(
                    "   Valid users              : " + validUsers);

                WriteToFile(
                    "   Subscriptions inserted   : " +
                    subscriptionsInserted);
            }
            catch (Exception ex)
            {
                WriteToFile(
                    "Error " +
                    nameof(Recherche_AD_FR_Pour_Tracts_Syndicaux) +
                    " : " +
                    ex.Message);

                if (graphService != null)
                {
                    SendTechnicalIssueMail(
                        nameof(Recherche_AD_FR_Pour_Tracts_Syndicaux),
                        insertSql,
                        ex.Message,
                        ex is SqlException
                            ? "SQL"
                            : "AD SYNCHRONIZATION");
                }

                throw;
            }
            finally
            {
                graphService = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }


        /// <summary>
        /// Reads and distributes emails for the currently selected shared mailbox.
        /// </summary>
        private void ReadCurrentSharedMailbox()
        {
            try
            {
                int nb_mail = 0;
                ValidateRequiredMailboxParameters();

                if (id_mailboxe <= 0)
                {
                    throw new InvalidOperationException(
                        "id_mailboxe is invalid for " +
                        sharedmailbox_name);
                }

                if (debug.ToUpper() == "TRUE")
                {
                    WriteToFile(
                        "   Connexion to " + sharedmailbox_name +
                        " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                    WriteToFile(
                        "   Sharedmailbox_folder_in  : " +
                        sharedmailbox_folder_in);
                    WriteToFile(
                        "   Sharedmailbox_folder_out : " +
                        sharedmailbox_folder_out);
                    WriteToFile(
                        "   Response mailbox          : " +
                        mailboxe_reponse);
                    WriteToFile(
                        "   Extracting the " + number_of_mails +
                        " oldest messages");
                }

                MailFolder inputFolder = GetInputFolder(graphService);
                MailFolder archiveFolder = GetChildFolderByName(graphService, sharedmailbox_name, sharedmailbox_folder_out);
                MailFolder errorFolder = GetChildFolderByName(graphService, sharedmailbox_name, "Erreur");

                MessageCollectionResponse messages = GetMessagesToProcess(graphService, inputFolder.Id);

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

                    foreach (Message emailSummary in messages.Value)
                    {
                        try
                        {
                            Message email = GetCompleteMessage(graphService, sharedmailbox_name, emailSummary.Id);

                            if (debug.ToUpper() == "TRUE")
                            {
                                WriteToFile(
                                    "       Subject : " + email.Subject +
                                    " at " +
                                    DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                            }

                            ProcessTractEmail(graphService, email);

                            MarkEmailAsRead(
                                graphService,
                                sharedmailbox_name,
                                email.Id);

                            MoveEmail(
                                graphService,
                                sharedmailbox_name,
                                email.Id,
                                archiveFolder.Id);

                            nb_mail++;
                            System.Threading.Thread.Sleep(500);
                        }
                        catch (Exception ex)
                        {
                            WriteToFile(
                                "       Error processing email : " + ex.Message +
                                " at " +
                                DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));

                            if (!(ex is TechnicalAlertAlreadySentException))
                            {
                                SendTechnicalIssueMail(
                                    nameof(ReadCurrentSharedMailbox),
                                    "",
                                    sharedmailbox_name + " - Mail : " +
                                    (emailSummary.Subject ?? "<no subject>") +
                                    " - " + ex.Message,
                                    "EMAIL PROCESSING");
                            }

                            try
                            {
                                MarkEmailAsRead(
                                    graphService,
                                    sharedmailbox_name,
                                    emailSummary.Id);

                                MoveEmail(
                                    graphService,
                                    sharedmailbox_name,
                                    emailSummary.Id,
                                    errorFolder.Id);
                            }
                            catch (Exception moveException)
                            {
                                WriteToFile(
                                    "       Error moving email to Erreur folder : " +
                                    moveException.Message);
                            }
                        }
                    }

                    if (debug.ToUpper() == "TRUE")
                    {
                        WriteToFile(
                            "       " + nb_mail +
                            " Email(s) have been processed at " +
                            DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                    }
                }
                else if (debug.ToUpper() == "TRUE")
                {
                    WriteToFile("       No emails found");
                }
            }
            catch (Exception ex)
            {
                WriteToFile(
                    "Global error ReadCurrentSharedMailbox : " + ex.Message);
                throw;
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private Message GetCompleteMessage(
            GraphServiceClient graphService,
            string mailbox,
            string messageId)
        {
            return graphService.Users[mailbox]
                .Messages[messageId]
                .GetAsync(config =>
                {
                    config.QueryParameters.Select = new string[]
                    {
                        "id",
                        "subject",
                        "body",
                        "from",
                        "sender",
                        "hasAttachments",
                        "receivedDateTime"
                    };
                })
                .GetAwaiter()
                .GetResult();
        }

        private void ProcessTractEmail(
            GraphServiceClient graphService,
            Message email)
        {
            string senderAddress = "";

            if (email.From != null &&
                email.From.EmailAddress != null)
            {
                senderAddress = email.From.EmailAddress.Address ?? "";
            }
            else if (email.Sender != null &&
                     email.Sender.EmailAddress != null)
            {
                senderAddress = email.Sender.EmailAddress.Address ?? "";
            }

            if (!IsAuthorizedSender(senderAddress))
            {
                throw new UnauthorizedAccessException(
                    "Unauthorized sender : " + senderAddress);
            }

            List<string> subscribers = GetSubscribers(id_mailboxe);

            if (subscribers.Count == 0)
            {
                WriteToFile(
                    "       No active subscriber for mailbox " +
                    sharedmailbox_name);
                return;
            }

            List<TractAttachment> attachments =
                GetAttachmentsWithGraph(
                    graphService,
                    sharedmailbox_name,
                    email.Id);

            string messageBody = BuildTractBody(email);
            BodyType messageBodyType = BodyType.Html;

            SendTractWithGraph(
                graphService,
                email.Subject ?? "",
                messageBody,
                messageBodyType,
                subscribers,
                attachments);
        }

        private string BuildTractBody(Message email)
        {
            string body = email.Body == null
                ? ""
                : email.Body.Content ?? "";

            if (email.Body != null &&
                email.Body.ContentType == BodyType.Text)
            {
                body = WebUtility.HtmlEncode(body)
                    .Replace(Environment.NewLine, "<br/>");
            }

            body = body +
                "<br/><br/><i>" +
                "Ce courriel est un tract syndical qui ne doit pas faire " +
                "l'objet d'une réponse généralisée ou d'un transfert " +
                "généralisé. Merci de répondre uniquement à l'expéditeur " +
                "via les boîtes mails dédiées.<br/>" +
                "Le non-respect de cette règle donnera lieu à des sanctions " +
                "disciplinaires.<br/>" +
                "À tout moment, vous avez le choix de recevoir ou non des " +
                "communications syndicales." +
                "</i>";

            return body;
        }

        private void SendTractWithGraph(
            GraphServiceClient graphService,
            string subject,
            string body,
            BodyType bodyType,
            List<string> subscribers,
            List<TractAttachment> attachments)
        {
            const int recipientsPerBatch = 250;

            for (int offset = 0;
                 offset < subscribers.Count;
                 offset += recipientsPerBatch)
            {
                List<Recipient> bccRecipients = subscribers
                    .Skip(offset)
                    .Take(recipientsPerBatch)
                    .Select(address => new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = address
                        }
                    })
                    .ToList();

                SendDraftWithAttachments(
                    graphService,
                    subject,
                    body,
                    bodyType,
                    bccRecipients,
                    attachments);
            }
        }

        private void SendDraftWithAttachments(
            GraphServiceClient graphService,
            string subject,
            string body,
            BodyType bodyType,
            List<Recipient> bccRecipients,
            List<TractAttachment> attachments)
        {
            Message draft = null;

            try
            {
                Message newMessage = new Message
                {
                    Subject = subject,
                    Body = new ItemBody
                    {
                        ContentType = bodyType,
                        Content = body
                    },
                    BccRecipients = bccRecipients
                };

                draft = graphService.Users[mailboxe_reponse]
                    .Messages
                    .PostAsync(newMessage)
                    .GetAwaiter()
                    .GetResult();

                if (draft == null ||
                    string.IsNullOrWhiteSpace(draft.Id))
                {
                    throw new Exception(
                        "Unable to create the Graph draft message.");
                }

                foreach (TractAttachment attachment in attachments)
                {
                    if (attachment.ContentBytes.LongLength <
                        simple_attachment_limit_bytes)
                    {
                        AddSmallAttachmentToDraft(
                            graphService,
                            draft.Id,
                            attachment);
                    }
                    else
                    {
                        AddLargeAttachmentToDraft(
                            graphService,
                            draft.Id,
                            attachment);
                    }
                }

                graphService.Users[mailboxe_reponse]
                    .Messages[draft.Id]
                    .Send
                    .PostAsync()
                    .GetAwaiter()
                    .GetResult();
            }
            catch
            {
                if (draft != null &&
                    !string.IsNullOrWhiteSpace(draft.Id))
                {
                    try
                    {
                        graphService.Users[mailboxe_reponse]
                            .Messages[draft.Id]
                            .DeleteAsync()
                            .GetAwaiter()
                            .GetResult();
                    }
                    catch
                    {
                    }
                }

                throw;
            }
        }

        private void AddSmallAttachmentToDraft(
            GraphServiceClient graphService,
            string draftId,
            TractAttachment attachment)
        {
            FileAttachment fileAttachment = new FileAttachment
            {
                OdataType = "#microsoft.graph.fileAttachment",
                Name = attachment.Name,
                ContentType = attachment.ContentType,
                ContentId = attachment.ContentId,
                IsInline = attachment.IsInline,
                ContentBytes = attachment.ContentBytes
            };

            graphService.Users[mailboxe_reponse]
                .Messages[draftId]
                .Attachments
                .PostAsync(fileAttachment)
                .GetAwaiter()
                .GetResult();
        }

        private void AddLargeAttachmentToDraft(
            GraphServiceClient graphService,
            string draftId,
            TractAttachment attachment)
        {
            if (attachment.ContentBytes.LongLength >
                maximum_attachment_size_bytes)
            {
                throw new InvalidOperationException(
                    "Attachment larger than 150 MB : " + attachment.Name);
            }

            AttachmentItem attachmentItem = new AttachmentItem
            {
                AttachmentType = AttachmentType.File,
                Name = attachment.Name,
                Size = attachment.ContentBytes.LongLength,
                ContentType = attachment.ContentType,
                IsInline = attachment.IsInline
            };

            var requestBody =
                new Microsoft.Graph.Users.Item.Messages.Item.Attachments
                    .CreateUploadSession.CreateUploadSessionPostRequestBody
                {
                    AttachmentItem = attachmentItem
                };

            UploadSession uploadSession =
                graphService.Users[mailboxe_reponse]
                    .Messages[draftId]
                    .Attachments
                    .CreateUploadSession
                    .PostAsync(requestBody)
                    .GetAwaiter()
                    .GetResult();

            if (uploadSession == null ||
                string.IsNullOrWhiteSpace(uploadSession.UploadUrl))
            {
                throw new Exception(
                    "Unable to create upload session for " +
                    attachment.Name);
            }

            UploadAttachmentRanges(
                uploadSession.UploadUrl,
                attachment.ContentBytes,
                attachment.Name);
        }

        private void UploadAttachmentRanges(
            string uploadUrl,
            byte[] contentBytes,
            string attachmentName)
        {
            using (HttpClient httpClient = new HttpClient())
            {
                httpClient.Timeout = TimeSpan.FromMinutes(15);
                long offset = 0;

                while (offset < contentBytes.LongLength)
                {
                    int currentSliceSize = (int)Math.Min(
                        upload_slice_size_bytes,
                        contentBytes.LongLength - offset);

                    using (ByteArrayContent content = new ByteArrayContent(
                        contentBytes,
                        (int)offset,
                        currentSliceSize))
                    {
                        content.Headers.ContentLength = currentSliceSize;
                        content.Headers.ContentRange =
                            new ContentRangeHeaderValue(
                                offset,
                                offset + currentSliceSize - 1,
                                contentBytes.LongLength);

                        HttpResponseMessage response = httpClient
                            .PutAsync(uploadUrl, content)
                            .GetAwaiter()
                            .GetResult();

                        string responseContent = response.Content
                            .ReadAsStringAsync()
                            .GetAwaiter()
                            .GetResult();

                        if (!response.IsSuccessStatusCode)
                        {
                            throw new HttpRequestException(
                                "Upload failed for " + attachmentName +
                                " at byte " + offset +
                                " : HTTP " +
                                (int)response.StatusCode +
                                " - " + responseContent);
                        }
                    }

                    offset += currentSliceSize;

                    if (debug.ToUpper() == "TRUE")
                    {
                        WriteToFile(
                            "       Upload " + attachmentName + " : " +
                            offset + "/" +
                            contentBytes.LongLength + " bytes");
                    }
                }
            }
        }

        private List<TractAttachment> GetAttachmentsWithGraph(
            GraphServiceClient graphService,
            string mailbox,
            string messageId)
        {
            List<TractAttachment> result =
                new List<TractAttachment>();

            AttachmentCollectionResponse attachments =
                graphService.Users[mailbox]
                    .Messages[messageId]
                    .Attachments
                    .GetAsync()
                    .GetAwaiter()
                    .GetResult();

            if (attachments == null ||
                attachments.Value == null)
            {
                return result;
            }

            foreach (Attachment attachment in attachments.Value)
            {
                if (!(attachment is FileAttachment))
                {
                    throw new NotSupportedException(
                        "Unsupported attachment type : " +
                        attachment.Name);
                }

                FileAttachment fileAttachment =
                    (FileAttachment)attachment;

                if (fileAttachment.ContentBytes == null &&
                    !string.IsNullOrWhiteSpace(fileAttachment.Id))
                {
                    fileAttachment =
                        graphService.Users[mailbox]
                            .Messages[messageId]
                            .Attachments[fileAttachment.Id]
                            .GetAsync()
                            .GetAwaiter()
                            .GetResult() as FileAttachment;
                }

                if (fileAttachment == null ||
                    fileAttachment.ContentBytes == null)
                {
                    throw new Exception(
                        "Attachment content is empty : " +
                        attachment.Name);
                }

                result.Add(new TractAttachment
                {
                    Name = fileAttachment.Name ?? "attachment",
                    ContentType = string.IsNullOrWhiteSpace(
                        fileAttachment.ContentType)
                        ? "application/octet-stream"
                        : fileAttachment.ContentType,
                    ContentId = fileAttachment.ContentId ?? "",
                    IsInline = fileAttachment.IsInline == true,
                    ContentBytes = fileAttachment.ContentBytes
                });
            }

            return result;
        }

        private List<string> GetSubscribers(int mailboxId)
        {
            const string sql = @"
        SELECT DISTINCT annuaire.dbo.tbl_users.user_email
        FROM T_TRACT_SYNDICAL_Abonnes
        INNER JOIN annuaire.dbo.tbl_users
            ON T_TRACT_SYNDICAL_Abonnes.id_user =
               annuaire.dbo.tbl_users.id_user
        WHERE T_TRACT_SYNDICAL_Abonnes.id_mailboxe = @ID_MAILBOXE
          AND T_TRACT_SYNDICAL_Abonnes.actif = 'O'
          AND annuaire.dbo.tbl_users.actif = 1
          AND annuaire.dbo.tbl_users.user_email IS NOT NULL
          AND annuaire.dbo.tbl_users.user_email <> ''
        ORDER BY annuaire.dbo.tbl_users.user_email";

            try
            {
                List<string> subscribers = new List<string>();

                using (SqlConnection connection =
                    new SqlConnection(sql_connexion))
                using (SqlCommand command =
                    new SqlCommand(sql, connection))
                {
                    command.CommandTimeout = 300;

                    command.Parameters.Add(
                        "@ID_MAILBOXE",
                        SqlDbType.Int).Value = mailboxId;

                    connection.Open();

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string address = Convert.ToString(
                                reader["user_email"]).Trim();

                            if (IsEmailAddress(address))
                            {
                                subscribers.Add(address);
                            }
                        }
                    }
                }

                return subscribers
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }
            catch (Exception ex)
            {
                SendTechnicalIssueMail(
                    nameof(GetSubscribers),
                    sql,
                    sharedmailbox_name + " - " + ex.Message,
                    "SQL");

                throw new TechnicalAlertAlreadySentException(
                    "SQL error in GetSubscribers for " +
                    sharedmailbox_name + " : " + ex.Message,
                    ex);
            }
        }

        private List<SharedMailboxConfiguration> GetActiveSharedMailboxes()
        {
            const string sql = @"
                SELECT
                    id_mailboxe,
                    nom_mailboxe,
                    mailboxe,
                    mailboxe_reponse,
                    expediteur_autorise,
                    ordre
                FROM T_TRACT_SYNDICAL_Mailboxes
                WHERE actif = 1
                ORDER BY ordre";

            try
            {
                List<SharedMailboxConfiguration> mailboxes =
                    new List<SharedMailboxConfiguration>();

                using (SqlConnection connection =
                    new SqlConnection(sql_connexion))
                using (SqlCommand command =
                    new SqlCommand(sql, connection))
                {
                    command.CommandTimeout = 300;
                    connection.Open();

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            mailboxes.Add(new SharedMailboxConfiguration
                            {
                                IdMailboxe = Convert.ToInt32(
                                    reader["id_mailboxe"]),
                                NomMailboxe = Convert.ToString(
                                    reader["nom_mailboxe"]).Trim(),
                                SharedMailboxName = Convert.ToString(
                                    reader["mailboxe"]).Trim(),
                                MailboxeReponse = Convert.ToString(
                                    reader["mailboxe_reponse"]).Trim(),
                                ExpediteurAutorise = Convert.ToString(
                                    reader["expediteur_autorise"]).Trim(),
                                Ordre = reader["ordre"] == DBNull.Value
                                    ? 0
                                    : Convert.ToInt32(reader["ordre"])
                            });
                        }
                    }
                }

                return mailboxes;
            }
            catch (Exception ex)
            {
                SendTechnicalIssueMail(
                    nameof(GetActiveSharedMailboxes),
                    sql,
                    ex.Message,
                    "SQL");

                throw new TechnicalAlertAlreadySentException(
                    "SQL error in GetActiveSharedMailboxes : " +
                    ex.Message,
                    ex);
            }
        }

        private bool IsAuthorizedSender(string senderAddress)
        {
            if (!IsEmailAddress(senderAddress))
            {
                return false;
            }

            string[] authorizedAddresses =
                (expediteur_autorise ?? "").Split(
                    new char[] { ';', ',' },
                    StringSplitOptions.RemoveEmptyEntries);

            return authorizedAddresses.Any(address =>
                string.Equals(
                    address.Trim(),
                    senderAddress.Trim(),
                    StringComparison.OrdinalIgnoreCase));
        }

        private void SendTechnicalIssueMail(
            string methodName,
            string sql,
            string message,
            string type_error)
        {
            try
            {
                WriteToFile(
                    "       " + type_error + " error in " +
                    methodName + " : " + message);

                if (graphService == null)
                {
                    WriteToFile(
                        "Unable to send technical issue email because " +
                        "GraphServiceClient is null");
                    return;
                }

                if (string.IsNullOrWhiteSpace(
                        email_in_case_of_technical_issue) ||
                    !email_in_case_of_technical_issue.Contains("@"))
                {
                    WriteToFile(
                        "Unable to send technical issue email because " +
                        "recipient is empty");
                    return;
                }

                string subject = global_application_name +
                    " - Erreur " + type_error +
                    " dans " + methodName;

                string body =
                    "Une erreur " + type_error +
                    " est survenue dans " +
                    global_application_name + ".<br/><br/>" +
                    "<b>Méthode :</b> " + methodName + "<br/>" +
                    "<b>Message :</b> " +
                    WebUtility.HtmlEncode(message) + "<br/><br/>";

                if (!string.IsNullOrWhiteSpace(sql))
                {
                    body = body +
                        "<b>Requête SQL :</b><br/><pre>" +
                        WebUtility.HtmlEncode(sql) +
                        "</pre>";
                }

                EnvoiEmail_with_Graph(
                    graphService,
                    subject,
                    body,
                    email_in_case_of_technical_issue,
                    "",
                    type_error);
            }
            catch (Exception mailEx)
            {
                WriteToFile(
                    "Error sending technical issue email : " +
                    mailEx.Message);
            }
        }

        private void EnvoiEmail_with_Graph(
            GraphServiceClient graphService,
            string subject,
            string body,
            string recipient,
            string attachmentPath = "",
            string type_error = "")
        {
            if (string.IsNullOrWhiteSpace(recipient) ||
                !recipient.Contains("@"))
            {
                return;
            }

            List<Recipient> toRecipients =
                BuildRecipients(recipient);

            Message message = new Message
            {
                Subject = subject,
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = body.Replace(
                        Environment.NewLine,
                        "<br/>")
                },
                ToRecipients = toRecipients
            };

            if (!string.IsNullOrWhiteSpace(attachmentPath) &&
                File.Exists(attachmentPath))
            {
                message.Attachments = new List<Attachment>
                {
                    new FileAttachment
                    {
                        OdataType = "#microsoft.graph.fileAttachment",
                        Name = Path.GetFileName(attachmentPath),
                        ContentType = "application/octet-stream",
                        ContentBytes = File.ReadAllBytes(attachmentPath)
                    }
                };
            }

            if (type_error.ToUpper() != "SQL")
            {
                List<Recipient> ccRecipients =
                    BuildRecipients(email_destinataire);

                if (ccRecipients.Count > 0)
                {
                    message.CcRecipients = ccRecipients;
                }
            }

            var requestBody =
                new Microsoft.Graph.Users.Item.SendMail
                    .SendMailPostRequestBody
                {
                    Message = message,
                    SaveToSentItems = true
                };

            string senderMailbox =
                !string.IsNullOrWhiteSpace(mailboxe_reponse)
                    ? mailboxe_reponse
                    : sharedmailbox_name;

            graphService.Users[senderMailbox]
                .SendMail
                .PostAsync(requestBody)
                .GetAwaiter()
                .GetResult();
        }

        private List<Recipient> BuildRecipients(string emails)
        {
            List<Recipient> recipients =
                new List<Recipient>();

            if (string.IsNullOrWhiteSpace(emails))
            {
                return recipients;
            }

            string[] splitEmails = emails.Split(
                new char[] { ';', ',' },
                StringSplitOptions.RemoveEmptyEntries);

            foreach (string email in splitEmails)
            {
                string address = email.Trim();

                if (IsEmailAddress(address))
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

        private void MarkEmailAsRead(
            GraphServiceClient graphService,
            string mailbox,
            string messageId)
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

        private void MoveEmail(
            GraphServiceClient graphService,
            string mailbox,
            string messageId,
            string destinationFolderId)
        {
            var requestBody =
                new Microsoft.Graph.Users.Item.Messages.Item.Move
                    .MovePostRequestBody
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
                "   Invalid start_date_scan value : " +
                start_date_scan);
            return "";
        }

        private MailFolder GetInputFolder(
            GraphServiceClient graphService)
        {
            if (sharedmailbox_folder_in.ToUpper().Trim() ==
                "INBOX")
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
            string safeFolderName =
                EscapeODataString(folderName);

            MailFolderCollectionResponse folders =
                graphService.Users[mailbox]
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
                throw new Exception(
                    "Folder not found : " + folderName);
            }

            return folders.Value.First();
        }

        private void ValidateRequiredCountryParameters(Country p)
        {

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

            if (string.IsNullOrWhiteSpace(
                email_in_case_of_technical_issue))
            {
                throw new Exception(
                    "email_in_case_of_technical_issue is empty. " +
                    "Parameter used : " +
                    email_in_case_of_technical_issue_parameter_global);
            }
        }

        private void ValidateRequiredMailboxParameters()
        {
            if (string.IsNullOrWhiteSpace(number_of_mails))
            {
                number_of_mails = "10";
            }

            if (string.IsNullOrWhiteSpace(
                sharedmailbox_folder_in))
            {
                sharedmailbox_folder_in = "INBOX";
            }

            if (string.IsNullOrWhiteSpace(
                sharedmailbox_folder_out))
            {
                sharedmailbox_folder_out = "Archives";
            }

            if (string.IsNullOrWhiteSpace(sharedmailbox_name))
            {
                throw new Exception(
                    "sharedmailbox_name is empty");
            }

            if (string.IsNullOrWhiteSpace(mailboxe_reponse))
            {
                throw new Exception(
                    "mailboxe_reponse is empty for " +
                    sharedmailbox_name);
            }

            if (string.IsNullOrWhiteSpace(
                expediteur_autorise))
            {
                throw new Exception(
                    "expediteur_autorise is empty for " +
                    sharedmailbox_name);
            }
        }

        private string get_IMCA_paramters(
            string sql_con,
            string param_name)
        {
            string ret = "";

            using (SqlConnection con =
                new SqlConnection(sql_con))
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
                logs_folder =
                    AppDomain.CurrentDomain.BaseDirectory;
            }

            if (!Directory.Exists(logs_folder))
            {
                Directory.CreateDirectory(logs_folder);
            }

            string filePath = Path.Combine(
                logs_folder,
                "IMCA_" + global_session_name + "_" +
                DateTime.Now.ToString("dd_MM_yyyy") + "_" +
                country + "_" +
                global_application_name + ".txt");

            File.AppendAllText(
                filePath,
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") +
                " - " + message + Environment.NewLine);
        }

        private string EscapeODataString(string value)
        {
            return value == null
                ? ""
                : value.Replace("'", "''");
        }

        private void DeleteFiles(
            string folder,
            string prefixfile = "")
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
                }
            }
        }

        private bool IsEmailAddress(string value)
        {
            return !string.IsNullOrWhiteSpace(value) &&
                Regex.IsMatch(
                    value.Trim(),
                    @"^[^\s@]+@[^\s@]+\.[^\s@]+$");
        }

        public void setParamCountry(string value)
        {
            country = value ?? "";
        }

        public void setParamSK_Valid(string value)
        {
            sk_valid = value ?? "";
        }

        public void setParamName(string value)
        {
            name = value ?? "";
        }

        public void setParamActive(string value)
        {
            active = value ?? "";
        }

        public void setParamDebug(string value)
        {
            debug = value ?? "";
        }

        public void setStartDateScan(string value)
        {
            start_date_scan = value ?? "";
        }

        public void setNumber_of_mails(string value)
        {
            number_of_mails =
                string.IsNullOrWhiteSpace(value)
                    ? "10"
                    : value;
        }

        public void setSharedMailboxName(string value)
        {
            sharedmailbox_name = value ?? "";
        }


        public void setNomMailboxe(string value)
        {
            nom_mailboxe = value ?? "";
        }

        public void setMailboxeReponse(string value)
        {
            mailboxe_reponse = value ?? "";
        }

        public void setExpediteurAutorise(string value)
        {
            expediteur_autorise = value ?? "";
        }

        public void setsharedmailbox_folder_in(string value)
        {
            sharedmailbox_folder_in =
                string.IsNullOrWhiteSpace(value)
                    ? "INBOX"
                    : value;
        }

        public void setsharedmailbox_folder_out(string value)
        {
            sharedmailbox_folder_out =
                string.IsNullOrWhiteSpace(value)
                    ? "Archives"
                    : value;
        }

        public void setEmailDestinataire(string value)
        {
            email_destinataire = value ?? "";
        }

        public void setlogs_folder(string value)
        {
            logs_folder = value ?? "";
        }

        public void settemp_folder(string value)
        {
            temp_folder = value ?? "";
        }

        public void setSqlConnexionParam(string value)
        {
            sql_connexion_parameter_global = value ?? "";
        }

        public void setEmailInCaseOfTechnicalIssueParam(
            string value)
        {
            email_in_case_of_technical_issue_parameter_global =
                value ?? "";
        }
    }
}
