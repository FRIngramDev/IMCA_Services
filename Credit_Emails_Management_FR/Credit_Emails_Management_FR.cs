using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using Newtonsoft.Json;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
namespace CREDIT_EMAILS_MANAGEMENT_FR
{
    public class CREDIT_EMAILS_MANAGEMENT_FR
    {
        private const string global_application_name = "CREDIT_EMAILS_MANAGEMENT_FR";
        private const string immutable_id_preference = "IdType=\"ImmutableId\"";
        private const string processed_property_id = "String {8BF48C6E-2C72-46F0-965D-919A7C2E54A9} Name IMCACreditProcessed";
        private const string credit_card_mailbox_name = "Cartes Bleues";
        private const string score_fraud_mailbox_name = "Surveillance ScoreFraud";
        private const string opening_account_mailbox_name = "Ouverture Comptes";
        private const string automatic_mail_archive_folder_name = "archives mails automatiques";
        private const string opening_processed_property_id = "String {8BF48C6E-2C72-46F0-965D-919A7C2E54A9} Name IMCAOpeningProcessed";
        private const int retry_delay_hours = 2;
        private string country = "", name = "", active = "", debug = "", start_date_scan = "", number_Of_Mails = "10";
        private string sharedmailbox_folder_in = "Inbox", sharedmailbox_folder_out = "Archives";
        private string sql_connexion_parameter_global = "", sql_connexion = "";
        private string sql_gestion_cdes_parameter_global = "", sql_gestion_cdes = "";
        private string sql_dss_copie_parameter_global = "", sql_dss_copie = "";
        private string sql_creation_compte_parameter_global = "", sql_creation_compte = "";
        private string email_in_case_of_technical_issue_parameter_global = "", email_in_case_of_technical_issue = "";
        private string fr_graph_send_as_parameter_global = "", fr_graph_send_as = "";
        private string credit_managers_contentieux = "", contentieux_recipients = "";
        private string path_archives = "";
        private string dss_con_openrowset_parameter_global = "";
        private string dss_con_openrowset = "";
        private string uri_webservice = "";
        private HashSet<string> credit_managers_contentieux_list = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private string logsFolder = "", tempFolder = "", sessionName = "";
        private string maquettes_rapports_pj_credit = "", fr_ouverture_graph_send_as = "", fr_credit_administration_graph_send_as = "";
        private string email_rapport_compteur = "", email_rapport_stat_ouverture = "", admin_ventes = "";
        private string templatesReportsFolder = "";
        private static readonly object logSyncRoot = new object();
        private GraphServiceClient graphService;
        private int mailboxId, refreshMinutes;
        private string mailboxName = "", mailboxAddress = "", inputFolderName = "", outputFolderName = "";
        private DateTime mailboxFilterDate = new DateTime(1900, 1, 1), lastRefresh = new DateTime(1900, 1, 1);
        private HashSet<string> allowedSenders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public sealed class JsonFile
        {
            public List<Country> countries
            {
                get;
                set;
            }
            = new List<Country>();
        }
        public sealed class Country
        {
            public string country { get; set; } = "";
            public string sk_valid { get; set; } = "";
            public string name { get; set; } = "";
            public string active { get; set; } = "TRUE";
            public string debug { get; set; } = "FALSE";
            public string start_date_scan { get; set; } = "";
            public string number_of_mails { get; set; } = "10";
            public string sharedmailbox_folder_in { get; set; } = "Inbox";
            public string sharedmailbox_folder_out { get; set; } = "Archives";

            public string sql_connexion_parameter_global { get; set; } = "";

            public string sql_gestion_cdes_parameter_global { get; set; } = "";

            public string sql_dss_copie_parameter_global { get; set; } = "";
            public string sql_creation_compte_parameter_global { get; set; } = "";

            public string email_in_case_of_technical_issue_parameter_global { get; set; } = "";

            public string fr_graph_send_as_parameter_global { get; set; } = "";

            public string credit_managers_contentieux { get; set; } = "";

            public string contentieux_recipients { get; set; } = "";
            public string path_archives { get; set; } = "";
            public string dss_con_openrowset_parameter_global { get; set; } = "";
            public string uri_webservice { get; set; } = "";
            public string maquettes_rapports_pj_credit { get; set; } = "";
            public string fr_ouverture_graph_send_as { get; set; } = "";
            public string fr_credit_administration_graph_send_as { get; set; } = "";
            public string email_rapport_compteur { get; set; } = "";
            public string email_rapport_stat_ouverture { get; set; } = "";
            public string admin_ventes { get; set; } = "";
        }

        private sealed class MailboxConfiguration
        {
            public int Id { get; set; }

            public string Name { get; set; } = "";

            public string Address { get; set; } = "";

            public DateTime FilterDate { get; set; }

            public int RefreshMinutes { get; set; }

            public DateTime LastRefresh { get; set; }

            public string InputFolder { get; set; } = "";

            public string OutputFolder { get; set; } = "";

            public string SenderAllowed { get; set; } = "";
        }

        private sealed class CreditMailData
        {
            public string CustomerName { get; set; } = "";

            public string CustomerCode { get; set; } = "";

            public string CustomerBranch { get; set; } = "";

            public string CustomerEmail { get; set; } = "";

            public DateTime RequestDate { get; set; }

            public double Amount { get; set; }

            public string TransactionStatus { get; set; } = "";

            public string DeliveryOrInvoice { get; set; } = "";

            public string AuthorizationNumber { get; set; } = "";

            public string Comments { get; set; } = "";
        }
        private enum ProcessingOutcome
        {
            Completed, Deferred, Ignored
        }

        public void Read_Email_with_Graph(string sql_con, string logs, string tmp_folder, string session_name)
        {
            string root = GetServicePath();
            logsFolder = Path.Combine(root, logs ?? "");
            tempFolder = Path.Combine(root, tmp_folder ?? "");
            sessionName = session_name ?? "";
            try
            {
                JsonFile cfg = JsonConvert.DeserializeObject<JsonFile>(GetImcaParameter(sql_con, global_application_name) ?? "");
                if (cfg?.countries == null || cfg.countries.Count == 0) throw new InvalidOperationException(global_application_name + " parameters are empty or invalid");
                foreach (Country item in cfg.countries)
                {
                    ApplyCountryConfiguration(item, sql_con);
                    if (!IsTrue(active)) continue;

                    WriteLog(
                        name.ToUpperInvariant() +
                        "(" + country.ToUpperInvariant() + ") at " +
                        DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));

                    WriteLog(
                        "   Debug Parameter is set to " +
                        debug.ToUpperInvariant());

                    ValidateCountryConfiguration();
                    Directory.CreateDirectory(tempFolder);

                    WriteLog(
                        "   Working folder uses IMCA temp folder : " +
                        tempFolder);

                    WriteLog("   Connecting to Microsoft Graph");
                    graphService = ConnectGraph();
                    WriteLog("   Microsoft Graph connection established");

                    List<MailboxConfiguration> mailboxes = GetActiveMailboxes();
                    if (mailboxes.Count == 0)
                    {
                        WriteLog(
                            "   No active shared mailbox found in T_SharedMailboxes");
                        continue;
                    }

                    WriteLog(
                        "   Active shared mailboxes loaded from T_SharedMailboxes : " +
                        mailboxes.Count);

                    var errors = new List<Exception>();
                    try
                    {
                        foreach (MailboxConfiguration box in mailboxes)
                        {
                            SetCurrentMailbox(box);

                            WriteLog(
                                "   Mailbox loaded from T_SharedMailboxes : ID " +
                                mailboxId + " - " + mailboxName + " - " + mailboxAddress);

                            WriteLog(
                                "       Input folder : " + inputFolderName +
                                " - Output folder : " + outputFolderName +
                                " - Refresh : " + refreshMinutes + " minute(s)");

                            if (IsTrue(debug))
                            {
                                WriteLog(
                                    "       Filter date : " +
                                    mailboxFilterDate.ToString("dd/MM/yyyy HH:mm:ss") +
                                    " - Last refresh : " +
                                    lastRefresh.ToString("dd/MM/yyyy HH:mm:ss"));

                                WriteLog(
                                    "       Allowed sender(s) : " +
                                    string.Join(";", allowedSenders));
                            }

                            if (!IsRefreshDue(lastRefresh, refreshMinutes))
                            {
                                if (IsTrue(debug))
                                {
                                    WriteLog(
                                        "   Mailbox skipped, refresh is not due : " +
                                        mailboxAddress);
                                }
                                continue;
                            }

                            try
                            {
                                WriteLog(
                                    "   Processing mailbox " +
                                    mailboxName + " : " + mailboxAddress);

                                DateTime? latest = ReadCurrentMailbox();
                                UpdateMailboxRefreshInformation(mailboxId, latest);

                                WriteLog(
                                    "   Mailbox processing completed : " +
                                    mailboxAddress);
                            }
                            catch (Exception ex)
                            {
                                UpsertMailboxError(ex);
                                errors.Add(ex);
                            }
                        }
                    }
                    finally
                    {
                        graphService = null;
                    }
                    if (errors.Count > 0)
                    {
                        WriteLog(
                            "   Country processing completed with " +
                            errors.Count +
                            " mailbox technical error(s). " +
                            "Errors were logged and alerted without stopping " +
                            "the application.");
                    }
                    else
                    {
                        WriteLog(
                            "   Country processing completed : " +
                            country.ToUpperInvariant());
                    }
                }
            }
            finally
            {
                graphService = null;
            }
        }

        private GraphServiceClient ConnectGraph()
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            return new class_dev_tools.Ews_Modern_Auth().Get_Graph_Service();
        }

        private void ApplyCountryConfiguration(Country i, string imca)
        {
            country = i.country ?? "";
            name = i.name ?? "";
            active = i.active ?? "";
            debug = i.debug ?? "";
            start_date_scan = i.start_date_scan ?? "";
            number_Of_Mails = string.IsNullOrWhiteSpace(i.number_of_mails) ? "10" : i.number_of_mails;
            sharedmailbox_folder_in = string.IsNullOrWhiteSpace(i.sharedmailbox_folder_in) ? "Inbox" : i.sharedmailbox_folder_in;
            sharedmailbox_folder_out = string.IsNullOrWhiteSpace(i.sharedmailbox_folder_out) ? "Archives" : i.sharedmailbox_folder_out;
            sql_connexion_parameter_global = i.sql_connexion_parameter_global ?? "";
            sql_gestion_cdes_parameter_global = i.sql_gestion_cdes_parameter_global ?? "";
            sql_dss_copie_parameter_global = i.sql_dss_copie_parameter_global ?? "";
            sql_creation_compte_parameter_global = i.sql_creation_compte_parameter_global ?? "";
            email_in_case_of_technical_issue_parameter_global = i.email_in_case_of_technical_issue_parameter_global ?? "";
            fr_graph_send_as_parameter_global = i.fr_graph_send_as_parameter_global ?? "";
            credit_managers_contentieux = i.credit_managers_contentieux ?? "";
            contentieux_recipients = i.contentieux_recipients ?? "";
            path_archives = i.path_archives ?? "";
            dss_con_openrowset_parameter_global = i.dss_con_openrowset_parameter_global ?? "";
            uri_webservice = i.uri_webservice ?? "";
            maquettes_rapports_pj_credit = i.maquettes_rapports_pj_credit ?? "";
            fr_ouverture_graph_send_as = i.fr_ouverture_graph_send_as ?? "";
            fr_credit_administration_graph_send_as = i.fr_credit_administration_graph_send_as ?? "";
            email_rapport_compteur = i.email_rapport_compteur ?? "";
            email_rapport_stat_ouverture = i.email_rapport_stat_ouverture ?? "";
            admin_ventes = i.admin_ventes ?? "";
            templatesReportsFolder = Path.Combine(GetServicePath(), maquettes_rapports_pj_credit);
            sql_connexion = GetImcaParameter(imca, sql_connexion_parameter_global);
            sql_gestion_cdes = GetImcaParameter(imca, sql_gestion_cdes_parameter_global);
            sql_dss_copie = GetImcaParameter(imca, sql_dss_copie_parameter_global);
            sql_creation_compte = GetImcaParameter(imca, sql_creation_compte_parameter_global);
            email_in_case_of_technical_issue = GetImcaParameter(imca, email_in_case_of_technical_issue_parameter_global);
            fr_graph_send_as = GetImcaParameter(imca, fr_graph_send_as_parameter_global);
            dss_con_openrowset = GetImcaParameter(imca, dss_con_openrowset_parameter_global);
            credit_managers_contentieux_list = SplitQuotedValues(credit_managers_contentieux).ToHashSet(StringComparer.OrdinalIgnoreCase);
        }

        private List<MailboxConfiguration> GetActiveMailboxes()
        {
            const string sql = @"SELECT id_mailboxe,ISNULL(nom_mailboxe,'') nom_mailboxe,ISNULL(mailboxe,'') mailboxe,ISNULL(dt_heure_filtre,'19000101') dt_heure_filtre,ISNULL(raffraichissement_min,0) raffraichissement_min,ISNULL(date_dernier_raf,'19000101') date_dernier_raf,ISNULL(folder_in,'') folder_in,ISNULL(folder_out,'') folder_out,ISNULL(sender_allowed,'') sender_allowed FROM dbo.T_SharedMailboxes WHERE actif=1 ORDER BY ISNULL(ordre,0),id_mailboxe;";
            var list = new List<MailboxConfiguration>();
            using (var c = new SqlConnection(sql_connexion)) using (var cmd = new SqlCommand(sql, c))
            {
                c.Open();
                using (var r = cmd.ExecuteReader()) while (r.Read()) list.Add(new MailboxConfiguration
                {
                    Id = Convert.ToInt32(r["id_mailboxe"]),
                    Name = Convert.ToString(r["nom_mailboxe"]).Trim(),
                    Address = Convert.ToString(r["mailboxe"]).Trim(),
                    FilterDate = Convert.ToDateTime(r["dt_heure_filtre"]),
                    RefreshMinutes = Convert.ToInt32(r["raffraichissement_min"]),
                    LastRefresh = Convert.ToDateTime(r["date_dernier_raf"]),
                    InputFolder = Convert.ToString(r["folder_in"]).Trim(),
                    OutputFolder = Convert.ToString(r["folder_out"]).Trim(),
                    SenderAllowed = Convert.ToString(r["sender_allowed"])
                }
                );
            }
            return list;
        }

        private void SetCurrentMailbox(MailboxConfiguration b)
        {
            mailboxId = b.Id;
            mailboxName = b.Name;
            mailboxAddress = b.Address;
            mailboxFilterDate = b.FilterDate;
            refreshMinutes = b.RefreshMinutes;
            lastRefresh = b.LastRefresh;
            inputFolderName = string.IsNullOrWhiteSpace(b.InputFolder) ? sharedmailbox_folder_in : b.InputFolder;
            outputFolderName = string.IsNullOrWhiteSpace(b.OutputFolder) ? sharedmailbox_folder_out : b.OutputFolder;
            allowedSenders = SplitValues(b.SenderAllowed).ToHashSet(StringComparer.OrdinalIgnoreCase);
        }

        private DateTime? ReadCurrentMailbox()
        {
            ValidateMailboxConfiguration();

            if (mailboxName.Equals(
                    score_fraud_mailbox_name,
                    StringComparison.OrdinalIgnoreCase))
            {
                return ReadScoreFraudMailbox();
            }

            if (mailboxName.Equals(
                    opening_account_mailbox_name,
                    StringComparison.OrdinalIgnoreCase))
            {
                DateTime? latestOpeningMessage =
                    ReadOpeningAccountMailbox();

                UploadOpeningMailsToDatabase();

                return latestOpeningMessage;
            }

            if (!mailboxName.Equals(
                    credit_card_mailbox_name,
                    StringComparison.OrdinalIgnoreCase))
            {
                WriteLog(
                    "       No processor implemented yet for : " +
                    mailboxName);
                return null;
            }

            MailFolder input = GetRequiredFolder(inputFolderName);
            MailFolder output = GetRequiredFolder(outputFolderName);
            List<Message> messages = GetCandidateMessages(input.Id);

            WriteLog(
                "       Email(s) found for processing : " +
                messages.Count);

            if (messages.Count == 0)
            {
                if (IsTrue(debug)) WriteLog("       No emails found");
                return null;
            }

            DateTime? latest = null;
            int completed = 0;
            int deferred = 0;
            int ignored = 0;
            int alreadyProcessed = 0;
            int errors = 0;

            foreach (Message summary in messages)
            {
                if (summary.ReceivedDateTime.HasValue)
                {
                    DateTime receivedDate =
                        summary.ReceivedDateTime.Value.LocalDateTime;

                    if (!latest.HasValue || receivedDate > latest.Value)
                        latest = receivedDate;
                }

                try
                {
                    if (IsMessageAlreadyProcessed(summary))
                    {
                        alreadyProcessed++;
                        continue;
                    }

                    Message email = GetCompleteMessage(summary.Id);
                    WriteLog(
                        "       Subject : " +
                        (email.Subject ?? "<no subject>"));

                    ProcessingOutcome outcome =
                        ProcessCreditMessage(email, output.Id);

                    switch (outcome)
                    {
                        case ProcessingOutcome.Completed:
                            completed++;
                            break;
                        case ProcessingOutcome.Deferred:
                            deferred++;
                            break;
                        case ProcessingOutcome.Ignored:
                            ignored++;
                            break;
                    }
                }
                catch (Exception ex)
                {
                    errors++;
                    UpsertProcessingState(summary, "E", null, ex.Message);
                    WriteLog("       Error processing email : " + ex.Message);
                    SendTechnicalAlert(
                        nameof(ReadCurrentMailbox),
                        mailboxAddress + " - Mail : " +
                        (summary.Subject ?? "<no subject>") + " - " +
                        ex.Message,
                        "CREDIT EMAIL PROCESSING");
                }
            }

            WriteLog(
                "       Email processing summary - Found : " + messages.Count +
                " - Completed : " + completed +
                " - Deferred : " + deferred +
                " - Ignored : " + ignored +
                " - Already processed : " + alreadyProcessed +
                " - Errors : " + errors);

            return latest;
        }

        private DateTime? ReadOpeningAccountMailbox()
        {
            if (string.IsNullOrWhiteSpace(sql_creation_compte)) throw new InvalidOperationException("sql_creation_compte is empty");
            MailFolder inbox = GetRequiredFolder(inputFolderName);
            MailFolder archive = GetOrCreateOpeningFolder(inbox.Id, automatic_mail_archive_folder_name);
            List<Message> messages = GetOpeningAccountMessages(inbox.Id);
            DateTime? latestTagged = null;
            int moved = 0, ignored = 0, alreadyProcessed = 0, errors = 0;
            foreach (Message summary in messages)
            {
                try
                {
                    if (IsOpeningMessageAlreadyProcessed(summary))
                    {
                        DateTime alreadyProcessedReceivedDate =
                            summary.ReceivedDateTime?.LocalDateTime ??
                            DateTime.Now;

                        if (!latestTagged.HasValue ||
                            alreadyProcessedReceivedDate > latestTagged.Value)
                        {
                            latestTagged = alreadyProcessedReceivedDate;
                        }

                        alreadyProcessed++;
                        WriteLog(
                            "       Opening account email already tagged : " +
                            (summary.Subject ?? "<no subject>"));
                        continue;
                    }
                    Message email = GetOpeningAccountMessage(summary.Id);
                    string sender = GetSender(email).Trim();
                    string subject = (email.Subject ?? "").Trim();
                    string dossierId = "";
                    bool matched = false;
                    if (sender.Equals("ouverture@ingrammicro.fr", StringComparison.OrdinalIgnoreCase) && subject.Equals("ouverture de compte client ingram micro", StringComparison.OrdinalIgnoreCase))
                    { matched = true; dossierId = GetOpeningDossierIdFromPdfAttachment(email.Id); }
                    else if (sender.Equals("analystes.credit@ingrammicro.fr", StringComparison.OrdinalIgnoreCase))
                    { matched = true; dossierId = GetOpeningDossierIdFromCustomerNumber(subject); }

                    DateTime received = email.ReceivedDateTime?.LocalDateTime ?? summary.ReceivedDateTime?.LocalDateTime ?? DateTime.Now;
                    if (!matched)
                    {
                        TryMarkOpeningMessageAsProcessed(email.Id, false);

                        UpsertProcessingState(
                            email,
                            "S",
                            null,
                            "Opening account email outside robot criteria - tagged and left unread");

                        if (!latestTagged.HasValue || received > latestTagged.Value)
                            latestTagged = received;

                        ignored++;
                        WriteLog(
                            "       Opening account email tagged, left unread " +
                            "and in Inbox : " + subject);
                        continue;
                    }
                    if (string.IsNullOrWhiteSpace(dossierId))
                    {
                        // Le message correspond à une règle métier du robot.
                        // L'absence de dossier est donc une anomalie technique :
                        // aucun tag, aucune lecture et aucun déplacement. Le mail
                        // sera repris au prochain passage après correction.
                        throw new InvalidOperationException(
                            "Opening account business email matched but no " +
                            "dossier ID could be resolved. Sender : " +
                            sender + " - Subject : " + subject);
                    }
                    MailFolder dossier = GetOrCreateOpeningFolder(archive.Id, dossierId);
                    MoveOpeningMessage(email.Id, dossier.Id);
                    TryMarkOpeningMessageAsProcessed(email.Id, true);

                    UpsertProcessingState(
                        email,
                        "S",
                        null,
                        "Opening account email processed and archived in dossier " +
                        dossierId);

                    if (!latestTagged.HasValue || received > latestTagged.Value)
                        latestTagged = received;
                    moved++;
                    WriteLog(
                        "       Opening account email tagged and moved to " +
                        automatic_mail_archive_folder_name +
                        "\\" + dossierId + " : " + (email.Subject ?? "<no subject>"));
                }
                catch (Exception ex)
                {
                    errors++;

                    UpsertProcessingState(
                        summary,
                        "E",
                        null,
                        ex.Message);

                    WriteLog(
                        "       Opening account email error stored in " +
                        "T_CreditMailProcessing : " + ex.Message);

                    SendTechnicalAlert(
                        nameof(ReadOpeningAccountMailbox),
                        mailboxAddress + " - Mail : " +
                        (summary.Subject ?? "<no subject>") + " - " +
                        ex.Message,
                        "OPENING ACCOUNT EMAIL PROCESSING");

                    // L'erreur est mémorisée avec le GraphMessageId. Le traitement
                    // des messages suivants continue et le mail en erreur sera
                    // rechargé explicitement au prochain passage.
                    continue;
                }
            }
            WriteLog("       Opening account summary - Candidates : " + messages.Count + " - Moved : " + moved + " - Ignored tagged unread : " + ignored + " - Already processed : " + alreadyProcessed + " - Errors : " + errors);
            return latestTagged;
        }

        private List<Message> GetOpeningAccountMessages(string folderId)
        {
            int top;
            if (!int.TryParse(number_Of_Mails, out top) || top <= 0)
                top = 10;

            DateTime date =
                mailboxFilterDate > new DateTime(1900, 1, 1)
                    ? mailboxFilterDate
                    : ParseStartDate();

            MessageCollectionResponse response = ExecuteGraphWithRetry(
                () => graphService.Users[mailboxAddress]
                    .MailFolders[folderId]
                    .Messages
                    .GetAsync(q =>
                    {
                        AddImmutableHeader(q.Headers);
                        q.QueryParameters.Top = top;
                        q.QueryParameters.Orderby =
                            new[] { "receivedDateTime asc" };
                        q.QueryParameters.Filter =
                            "receivedDateTime gt " +
                            date.ToUniversalTime().ToString(
                                "yyyy-MM-ddTHH:mm:ss.fffZ",
                                CultureInfo.InvariantCulture);
                        q.QueryParameters.Select = new[]
                        {
                            "id",
                            "subject",
                            "receivedDateTime",
                            "from",
                            "hasAttachments",
                            "internetMessageId"
                        };
                        q.QueryParameters.Expand = new[]
                        {
                            "singleValueExtendedProperties($filter=id eq '" +
                            EscapeODataString(opening_processed_property_id) +
                            "')"
                        };
                    })
                    .GetAwaiter()
                    .GetResult(),
                "List opening account messages");

            var byId =
                (response?.Value ?? new List<Message>())
                .Where(message => !string.IsNullOrWhiteSpace(message.Id))
                .ToDictionary(
                    message => message.Id,
                    message => message,
                    StringComparer.Ordinal);

            foreach (string graphMessageId in GetOpeningErrorMessageIds())
            {
                if (byId.ContainsKey(graphMessageId))
                    continue;

                try
                {
                    Message errorMessage =
                        graphService.Users[mailboxAddress]
                            .Messages[graphMessageId]
                            .GetAsync(q =>
                            {
                                AddImmutableHeader(q.Headers);
                                q.QueryParameters.Select = new[]
                                {
                                    "id",
                                    "subject",
                                    "receivedDateTime",
                                    "from",
                                    "hasAttachments",
                                    "internetMessageId"
                                };
                                q.QueryParameters.Expand = new[]
                                {
                                    "singleValueExtendedProperties($filter=id eq '" +
                                    EscapeODataString(opening_processed_property_id) +
                                    "')"
                                };
                            })
                            .GetAwaiter()
                            .GetResult();

                    if (errorMessage != null &&
                        !string.IsNullOrWhiteSpace(errorMessage.Id))
                    {
                        byId[errorMessage.Id] = errorMessage;
                    }
                }
                catch (Exception ex) when (IsGraphObjectNotFound(ex))
                {
                    MarkOpeningErrorMessageUnavailable(
                        graphMessageId,
                        ex.Message);

                    WriteLog(
                        "       Opening account message stored in error is no " +
                        "longer available in Graph : " + graphMessageId);
                }
            }

            return byId.Values
                .OrderBy(message => message.ReceivedDateTime)
                .ToList();
        }

        private List<string> GetOpeningErrorMessageIds()
        {
            const string sql = @"
SELECT GraphMessageId
FROM dbo.T_CreditMailProcessing
WHERE MailboxId = @MAILBOX_ID
  AND ProcessingStatus = 'E'
  AND NULLIF(LTRIM(RTRIM(GraphMessageId)), '') IS NOT NULL
ORDER BY Id;";

            var result = new List<string>();

            using (var connection = new SqlConnection(sql_connexion))
            using (var command = new SqlCommand(sql, connection))
            {
                command.CommandTimeout = 300;
                command.Parameters.Add(
                    "@MAILBOX_ID",
                    SqlDbType.Int).Value = mailboxId;

                connection.Open();

                using (SqlDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string graphMessageId =
                            Convert.ToString(reader["GraphMessageId"]).Trim();

                        if (!string.IsNullOrWhiteSpace(graphMessageId))
                            result.Add(graphMessageId);
                    }
                }
            }

            return result;
        }

        private void MarkOpeningErrorMessageUnavailable(
            string graphMessageId,
            string error)
        {
            const string sql = @"
UPDATE dbo.T_CreditMailProcessing
SET ProcessingStatus = 'I',
    ErrorMessage = @ERROR,
    ProcessedAtUtc = SYSUTCDATETIME()
WHERE MailboxId = @MAILBOX_ID
  AND GraphMessageId = @GRAPH_MESSAGE_ID;";

            ExecuteNonQuery(
                sql_connexion,
                sql,
                new SqlParameter(
                    "@ERROR",
                    SqlDbType.NVarChar,
                    2000)
                {
                    Value = Truncate(
                        "Graph message unavailable during retry - " + error,
                        2000)
                },
                new SqlParameter(
                    "@MAILBOX_ID",
                    SqlDbType.Int)
                {
                    Value = mailboxId
                },
                new SqlParameter(
                    "@GRAPH_MESSAGE_ID",
                    SqlDbType.NVarChar,
                    500)
                {
                    Value = graphMessageId
                });
        }

        private Message GetOpeningAccountMessage(string id)
        {
            return ExecuteGraphWithRetry(() => graphService.Users[mailboxAddress].Messages[id].GetAsync(q => { AddImmutableHeader(q.Headers); q.QueryParameters.Select = new[] { "id", "subject", "receivedDateTime", "from", "sender", "hasAttachments", "internetMessageId" }; }).GetAwaiter().GetResult(), "Get opening account message");
        }

        private static bool IsOpeningMessageAlreadyProcessed(Message m)
        {
            return m?.SingleValueExtendedProperties?.Any(x => string.Equals(x.Id, opening_processed_property_id, StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(x.Value)) == true;
        }

        private void TryMarkOpeningMessageAsProcessed(string id, bool markAsRead)
        {
            ExecuteGraphWithRetry(() => { graphService.Users[mailboxAddress].Messages[id].PatchAsync(new Message { IsRead = markAsRead ? (bool?)true : null, SingleValueExtendedProperties = new List<SingleValueLegacyExtendedProperty> { new SingleValueLegacyExtendedProperty { Id = opening_processed_property_id, Value = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ", CultureInfo.InvariantCulture) } } }, q => AddImmutableHeader(q.Headers)).GetAwaiter().GetResult(); return true; }, "Tag opening account message");
        }

        private void MoveOpeningMessage(string id, string destinationId)
        {
            ExecuteGraphWithRetry(() => { graphService.Users[mailboxAddress].Messages[id].Move.PostAsync(new Microsoft.Graph.Users.Item.Messages.Item.Move.MovePostRequestBody { DestinationId = destinationId }, q => AddImmutableHeader(q.Headers)).GetAwaiter().GetResult(); return true; }, "Move opening account message");
        }

        private string GetOpeningDossierIdFromPdfAttachment(string id)
        {
            AttachmentCollectionResponse r = ExecuteGraphWithRetry(() => graphService.Users[mailboxAddress].Messages[id].Attachments.GetAsync(q => { AddImmutableHeader(q.Headers); q.QueryParameters.Top = 999; }).GetAwaiter().GetResult(), "Read opening attachments");
            foreach (Microsoft.Graph.Models.Attachment a in r?.Value ?? new List<Microsoft.Graph.Models.Attachment>())
            { string n = a?.Name ?? ""; if (n.IndexOf(".PDF", StringComparison.OrdinalIgnoreCase) < 0) continue; string v = n.ToUpperInvariant().Replace("RECAPITULATIF_", "").Replace(".PDF", "").Trim(); if (v.All(char.IsDigit)) return v; }
            return "";
        }

        private string GetOpeningDossierIdFromCustomerNumber(string subject)
        {
            if (string.IsNullOrWhiteSpace(subject) || subject.Length < 8) return "";
            return ExecuteScalarString(sql_creation_compte, "SELECT TOP (1) CAST(id AS varchar(20)) FROM dbo.T_customer WHERE ImpCustNbr=@C ORDER BY id DESC;", new SqlParameter("@C", SqlDbType.VarChar, 50) { Value = subject.Substring(2, 6) });
        }

        private MailFolder GetOrCreateOpeningFolder(string parentId, string name)
        {
            MailFolder f = FindOpeningFolder(parentId, name); if (f != null) return f;
            return ExecuteGraphWithRetry(() => graphService.Users[mailboxAddress].MailFolders[parentId].ChildFolders.PostAsync(new MailFolder { DisplayName = name, IsHidden = false }).GetAwaiter().GetResult(), "Create opening folder " + name);
        }

        private MailFolder FindOpeningFolder(string parentId, string name)
        {
            MailFolderCollectionResponse r = ExecuteGraphWithRetry(() => graphService.Users[mailboxAddress].MailFolders[parentId].ChildFolders.GetAsync(q => { q.QueryParameters.Top = 999; q.QueryParameters.Filter = "displayName eq '" + EscapeODataString(name) + "'"; }).GetAwaiter().GetResult(), "Find opening folder " + name);
            return r?.Value?.FirstOrDefault(x => string.Equals(x.DisplayName, name, StringComparison.OrdinalIgnoreCase));
        }

        private void UploadOpeningMailsToDatabase()
        {
            if (string.IsNullOrWhiteSpace(sql_creation_compte))
            {
                throw new InvalidOperationException(
                    "sql_creation_compte is empty");
            }

            MailFolder inbox = GetRequiredFolder(inputFolderName);
            MailFolder automaticArchive = FindOpeningFolder(
                inbox.Id,
                automatic_mail_archive_folder_name);

            if (automaticArchive == null)
            {
                if (IsTrue(debug))
                {
                    WriteLog(
                        "       Opening account upload skipped: folder " +
                        automatic_mail_archive_folder_name +
                        " was not found");
                }

                return;
            }

            List<MailFolder> dossierFolders =
                GetOpeningDossierFolders(automaticArchive.Id);

            WriteLog(
                "       Opening account dossier folder(s) found for database upload : " +
                dossierFolders.Count);

            int importedMessages = 0;
            int deletedFolders = 0;
            int errors = 0;

            foreach (MailFolder dossierFolder in dossierFolders)
            {
                string dossierName =
                    (dossierFolder.DisplayName ?? "").Trim();

                if (!int.TryParse(
                        dossierName,
                        NumberStyles.None,
                        CultureInfo.InvariantCulture,
                        out int dossierId))
                {
                    if (IsTrue(debug))
                    {
                        WriteLog(
                            "       Opening account upload ignored non-numeric " +
                            "folder : " + dossierName);
                    }

                    continue;
                }

                try
                {
                    List<Message> archivedMessages =
                        GetOpeningArchivedMessages(dossierFolder.Id);

                    WriteLog(
                        "       Opening account mail(s) found for database upload" +
                        " - Dossier : " + dossierId +
                        " - Found : " + archivedMessages.Count);

                    foreach (Message summary in archivedMessages)
                    {
                        Message email =
                            GetOpeningArchivedMessage(summary.Id);

                        byte[] mime = GetMimeContent(email.Id);

                        InsertOpeningMailInDatabase(
                            dossierId,
                            email,
                            mime);

                        PermanentlyDeleteOpeningMessage(email.Id);
                        importedMessages++;

                        WriteLog(
                            "       Opening account mail imported in database" +
                        " - Dossier : " + dossierId + " - Message permanently deleted from mailbox" +
                        " : " + (email.Subject ?? "<no subject>"));
                    }

                    // Comme l'ancien upload_mail_en_bdd_general, le dossier
                    // numérique est supprimé après l'import de tous ses mails.
                    PermanentlyDeleteOpeningFolder(dossierFolder.Id);
                    UpdateOpeningUploadStatus(dossierId);
                    deletedFolders++;

                    WriteLog(
                        "       Opening account dossier folder permanently deleted" +
                        " - Dossier : " + dossierId);
                }
                catch (Exception ex)
                {
                    errors++;

                    WriteLog(
                        "       Opening account database upload error" +
                        " - Dossier : " + dossierId +
                        " - Error : " + ex.Message);

                    SendTechnicalAlert(
                        nameof(UploadOpeningMailsToDatabase),
                        mailboxAddress +
                        " - Dossier : " +
                        dossierId +
                        " - " +
                        ex.Message,
                        "OPENING ACCOUNT DATABASE UPLOAD");
                }
            }

            WriteLog(
                "       Opening account database upload summary" +
                " - Dossier folders : " + dossierFolders.Count +
                " - Imported messages : " + importedMessages +
                " - Deleted folders : " + deletedFolders +
                " - Errors : " + errors);
        }

        private List<MailFolder> GetOpeningDossierFolders(
            string automaticArchiveFolderId)
        {
            MailFolderCollectionResponse response =
                ExecuteGraphWithRetry(
                    () => graphService.Users[mailboxAddress]
                        .MailFolders[automaticArchiveFolderId]
                        .ChildFolders
                        .GetAsync(q =>
                        {
                            q.QueryParameters.Top = 999;
                            q.QueryParameters.Select = new[]
                            {
                                "id",
                                "displayName",
                                "totalItemCount",
                                "childFolderCount"
                            };
                        })
                        .GetAwaiter()
                        .GetResult(),
                    "List opening account dossier folders");

            return response?.Value ?? new List<MailFolder>();
        }

        private List<Message> GetOpeningArchivedMessages(
            string folderId)
        {
            MessageCollectionResponse response =
                ExecuteGraphWithRetry(
                    () => graphService.Users[mailboxAddress]
                        .MailFolders[folderId]
                        .Messages
                        .GetAsync(q =>
                        {
                            AddImmutableHeader(q.Headers);
                            q.QueryParameters.Top = 999;
                            q.QueryParameters.Orderby =
                                new[] { "receivedDateTime asc" };
                            q.QueryParameters.Select = new[]
                            {
                                "id",
                                "subject",
                                "receivedDateTime",
                                "internetMessageId"
                            };
                        })
                        .GetAwaiter()
                        .GetResult(),
                    "List opening account archived messages");

            return response?.Value ?? new List<Message>();
        }

        private Message GetOpeningArchivedMessage(string messageId)
        {
            return ExecuteGraphWithRetry(
                () => graphService.Users[mailboxAddress]
                    .Messages[messageId]
                    .GetAsync(q =>
                    {
                        q.Headers.Add(
                            "Prefer",
                            "outlook.body-content-type=\"text\", " +
                            immutable_id_preference);

                        q.QueryParameters.Select = new[]
                        {
                            "id",
                            "subject",
                            "receivedDateTime",
                            "createdDateTime",
                            "toRecipients",
                            "internetMessageId"
                        };
                    })
                    .GetAwaiter()
                    .GetResult(),
                "Get opening account archived message");
        }

        private void InsertOpeningMailInDatabase(
            int dossierId,
            Message email,
            byte[] mime)
        {
            if (dossierId <= 0)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(dossierId));
            }

            if (mime == null || mime.Length == 0)
            {
                throw new InvalidDataException(
                    "Opening account MIME content is empty");
            }

            DateTime receivedDate =
                email.ReceivedDateTime?.LocalDateTime ??
                email.CreatedDateTime?.LocalDateTime ??
                DateTime.Now;

            string recipients =
                string.Join(
                    ";",
                    (email.ToRecipients ?? new List<Recipient>())
                        .Select(recipient =>
                            recipient?.EmailAddress?.Address)
                        .Where(address =>
                            !string.IsNullOrWhiteSpace(address)));

            string fileName =
                dossierId +
                "_EMAIL_" +
                receivedDate.ToString(
                    "yyyyMMdd_HHmmss",
                    CultureInfo.InvariantCulture);

            using (var connection =
                new SqlConnection(sql_creation_compte))
            using (var command =
                new SqlCommand(
                    "dbo.USP_Insert_mail_boite_ouverture",
                    connection))
            {
                command.CommandType = CommandType.StoredProcedure;
                command.CommandTimeout = 300;

                command.Parameters.Add(
                    "@id_dossier",
                    SqlDbType.Int).Value = dossierId;

                command.Parameters.Add(
                    "@fichier",
                    SqlDbType.Image).Value = mime;

                command.Parameters.Add(
                    "@nom_fichier",
                    SqlDbType.VarChar,
                    500).Value = Truncate(fileName, 500);

                command.Parameters.Add(
                    "@extension",
                    SqlDbType.NChar,
                    10).Value = "eml";

                command.Parameters.Add(
                    "@date_mail",
                    SqlDbType.Date).Value = receivedDate.Date;

                command.Parameters.Add(
                    "@sujet_mail",
                    SqlDbType.VarChar,
                    -1).Value = email.Subject ?? "";

                command.Parameters.Add(
                    "@destinataire_mail",
                    SqlDbType.VarChar,
                    -1).Value = recipients;

                command.Parameters.Add(
                    "@date_importation",
                    SqlDbType.Date).Value = DateTime.Now.Date;

                connection.Open();
                command.ExecuteNonQuery();
            }
        }

        private void PermanentlyDeleteOpeningMessage(string messageId)
        {
            ExecuteGraphWithRetry(
                () =>
                {
                    graphService.Users[mailboxAddress]
                        .Messages[messageId]
                        .PermanentDelete
                        .PostAsync(q => AddImmutableHeader(q.Headers))
                        .GetAwaiter()
                        .GetResult();

                    return true;
                },
                "Permanently delete imported opening account message");
        }

        private void PermanentlyDeleteOpeningFolder(string folderId)
        {
            ExecuteGraphWithRetry(
                () =>
                {
                    graphService.Users[mailboxAddress]
                        .MailFolders[folderId]
                        .PermanentDelete
                        .PostAsync()
                        .GetAwaiter()
                        .GetResult();

                    return true;
                },
                "Permanently delete imported opening account folder");
        }

        private void UpdateOpeningUploadStatus(int dossierId)
        {
            ExecuteNonQuery(
                sql_creation_compte,
                "UPDATE dbo.envoi_mail " +
                "SET upload_en_base='F' " +
                "WHERE id_dossier=@DOSSIER_ID;",
                new SqlParameter(
                    "@DOSSIER_ID",
                    SqlDbType.Int)
                {
                    Value = dossierId
                });
        }

        private DateTime? ReadScoreFraudMailbox()
        {
            if (string.IsNullOrWhiteSpace(path_archives))
            {
                throw new InvalidOperationException(
                    "path_archives is empty for Surveillance ScoreFraud");
            }

            Directory.CreateDirectory(path_archives);

            MailFolder input = GetRequiredFolder(inputFolderName);
            MailFolder output = GetRequiredFolder(outputFolderName);
            List<Message> messages = GetCandidateMessages(input.Id);

            WriteLog(
                "       ScoreFraud email(s) found for processing : " +
                messages.Count);

            DateTime? latest = null;
            int archived = 0;
            int deleted = 0;
            int xmlFiles = 0;
            int errors = 0;

            foreach (Message summary in messages)
            {
                DateTime receptionDate =
                    summary.ReceivedDateTime?.LocalDateTime ?? DateTime.Now;

                if (!latest.HasValue || receptionDate > latest.Value)
                {
                    latest = receptionDate;
                }

                try
                {
                    Message email = GetScoreFraudMessage(summary.Id);

                    if (email.ReceivedDateTime.HasValue)
                    {
                        receptionDate =
                            email.ReceivedDateTime.Value.LocalDateTime;
                    }

                    WriteLog(
                        "       Subject : " +
                        (email.Subject ?? "<no subject>"));

                    AttachmentCollectionResponse response =
                        GetScoreFraudAttachments(email.Id);

                    bool containsCsv = false;
                    int messageXmlFiles = 0;

                    foreach (Microsoft.Graph.Models.Attachment attachment
                        in response?.Value
                        ?? new List<Microsoft.Graph.Models.Attachment>())
                    {
                        string attachmentName = attachment?.Name ?? "";

                        // Comportement historique : recherche de la chaîne
                        // .CSV ou .XML dans le nom, sans exiger que ce soit
                        // l'extension finale.
                        if (attachmentName.IndexOf(
                                ".CSV",
                                StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            containsCsv = true;
                        }

                        if (attachmentName.IndexOf(
                                ".XML",
                                StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            FileAttachment file =
                                GetFileAttachmentContent(
                                    email.Id,
                                    attachment);

                            if (file?.ContentBytes == null)
                            {
                                throw new InvalidDataException(
                                    "Attachment content is empty : " +
                                    attachmentName);
                            }

                            string safeAttachmentName =
                                CleanFileName(attachmentName);

                            string workingXmlPath =
                                Path.Combine(
                                    tempFolder,
                                    safeAttachmentName);

                            Directory.CreateDirectory(tempFolder);
                            File.WriteAllBytes(
                                workingXmlPath,
                                file.ContentBytes);

                            WriteLog(
                                "       ScoreFraud XML attachment downloaded : " +
                                workingXmlPath);

                            string archivedXmlPath =
                                TraitementFichiersXmlNew(
                                    workingXmlPath,
                                    safeAttachmentName,
                                    receptionDate);

                            messageXmlFiles++;
                            xmlFiles++;

                            WriteLog(
                                "       ScoreFraud XML processing completed : " +
                                archivedXmlPath);
                        }
                    }

                    if (containsCsv)
                    {
                        // Équivalent Graph du DeleteMode.HardDelete historique.
                        PermanentlyDeleteMessage(email.Id);
                        deleted++;

                        WriteLog(
                            "       Message permanently deleted because " +
                            "a CSV attachment was found");
                    }
                    else
                    {
                        // Comportement historique : lecture puis déplacement
                        // dans le dossier Archives.
                        MarkMessageAsReadAndMove(
                            email.Id,
                            output.Id);

                        archived++;

                        WriteLog(
                            "       Message marked as read and moved to " +
                            outputFolderName +
                            " - XML attachment(s) : " +
                            messageXmlFiles);
                    }

                }
                catch (Exception ex)
                {
                    errors++;

                    WriteLog(
                        "       ScoreFraud email processing error : " +
                        ex.Message);

                    SendTechnicalAlert(
                        nameof(ReadScoreFraudMailbox),
                        mailboxAddress +
                        " - Mail : " +
                        (summary.Subject ?? "<no subject>") +
                        " - " +
                        ex.Message,
                        "SCOREFRAUD EMAIL PROCESSING");
                }
            }

            WriteLog(
                "       ScoreFraud email processing summary - Found : " +
                messages.Count +
                " - Archived : " + archived +
                " - Permanently deleted : " + deleted +
                " - XML files saved : " + xmlFiles +
                " - Errors : " + errors);

            return latest;
        }

        private string TraitementFichiersXmlNew(string fullPath, string fileName, DateTime graphReceptionDate)
        {
            if (string.IsNullOrWhiteSpace(fullPath) || !File.Exists(fullPath)) throw new FileNotFoundException("ScoreFraud XML file not found", fullPath);
            if (string.IsNullOrWhiteSpace(uri_webservice)) throw new InvalidOperationException("uri_webservice is empty");
            if (string.IsNullOrWhiteSpace(dss_con_openrowset)) throw new InvalidOperationException("dss_con_openrowset is empty");

            string period = ExtractHistoricalScoreFraudPeriod(fileName);
            DateTime archiveDate = DateTime.Now;
            var document = new XmlDocument();
            document.Load(fullPath);
            XmlNodeList reports = document.DocumentElement?.SelectNodes("//reports/report");
            if (reports == null || reports.Count == 0) throw new InvalidDataException("No //reports/report node found in " + fileName);

            using (var connection = new SqlConnection(sql_connexion))
            {
                connection.Open();
                ExecuteHistoricalSql(connection, "DELETE FROM dbo.T_messages_ORT WHERE date_reception IS NULL;");
                foreach (XmlNode report in reports)
                {
                    InsertHistoricalScoreFraudReport(connection, "dbo.T_messages_ORT", report);
                    InsertHistoricalScoreFraudReport(connection, "dbo.T_messages_ORT_TRANSFERT_CAPGEMINI", report);
                }
                ExecuteHistoricalSql(connection, "UPDATE dbo.T_messages_ORT SET top_selection=NULL,note=0 WHERE date_reception IS NULL; UPDATE dbo.T_messages_ORT_TRANSFERT_CAPGEMINI SET top_selection=NULL,note=0 WHERE date_reception IS NULL;");
                EnrichHistoricalScoreFraudRows(connection);
                string[] historicalQueries = new[]
                {
                @"UPDATE T_messages_ORT_TRANSFERT_CAPGEMINI SET id_rapport=(select id_rapport from T_messages_ORT where  T_messages_ORT.date_reception is null and  T_messages_ORT.num_siren=T_messages_ORT_TRANSFERT_CAPGEMINI.num_siren), valeur_score=(select valeur_score from T_messages_ORT where  T_messages_ORT.date_reception is null and  T_messages_ORT.num_siren=T_messages_ORT_TRANSFERT_CAPGEMINI.num_siren) WHERE  date_reception is null ",
                @"UPDATE T_messages_ORT SET top_selection='O'  WHERE(date_reception Is null) AND EXISTS(Select 1 FROM T_GENERAL as gen WHERE (NB_PRIVILEGE >0 OR NB_PRIVILEGE_TRESOR>0) AND T_messages_ORT.id_rapport = gen.id_rapport)",
                @"UPDATE T_messages_ORT_TRANSFERT_CAPGEMINI SET top_selection='O'  WHERE(date_reception Is null) AND EXISTS(Select 1 FROM T_GENERAL as gen WHERE (NB_PRIVILEGE >0 OR NB_PRIVILEGE_TRESOR>0) AND T_messages_ORT_TRANSFERT_CAPGEMINI.id_rapport = gen.id_rapport)",
                @"UPDATE T_messages_ORT SET top_selection='O'  WHERE date_reception Is null AND EXISTS (SELECT 1 FROM T_BODACC as Bodac WHERE lower(ref_evenement) = 'modification sur les représentants' AND datediff(month,date_parution,getdate())<=6              AND T_messages_ORT.id_rapport = Bodac.id_rapport) AND EXISTS (SELECT 1 FROM T_BODACC as Bodac WHERE lower(ref_evenement) = 'modification de l''adresse de l''établissement principal'              AND datediff(month,date_parution,getdate())<=6 AND T_messages_ORT.id_rapport = Bodac.id_rapport)",
                @"UPDATE T_messages_ORT_TRANSFERT_CAPGEMINI SET top_selection='O'  WHERE date_reception Is null AND EXISTS (SELECT 1 FROM T_BODACC as Bodac WHERE lower(ref_evenement) = 'modification sur les représentants' AND datediff(month,date_parution,getdate())<=6              AND T_messages_ORT_TRANSFERT_CAPGEMINI.id_rapport = Bodac.id_rapport) AND EXISTS (SELECT 1 FROM T_BODACC as Bodac WHERE lower(ref_evenement) = 'modification de l''adresse de l''établissement principal'              AND datediff(month,date_parution,getdate())<=6 AND T_messages_ORT_TRANSFERT_CAPGEMINI.id_rapport = Bodac.id_rapport)",
                @"UPDATE T_messages_ORT SET top_selection='O'  WHERE date_reception Is null AND EXISTS (SELECT 1 FROM T_BODACC as Bodac WHERE lower(ref_evenement) = 'modification sur les représentants' AND datediff(month,date_parution,getdate())<=6              AND T_messages_ORT.id_rapport = Bodac.id_rapport)",
                @"UPDATE T_messages_ORT_TRANSFERT_CAPGEMINI SET top_selection='O'  WHERE date_reception Is null AND EXISTS (SELECT 1 FROM T_BODACC as Bodac WHERE lower(ref_evenement) = 'modification sur les représentants' AND datediff(month,date_parution,getdate())<=6              AND T_messages_ORT_TRANSFERT_CAPGEMINI.id_rapport = Bodac.id_rapport)",
                @"UPDATE T_messages_ORT SET top_selection='O'  WHERE(date_reception Is null) AND EXISTS(Select 1 FROM T_GENERAL as gen WHERE (PROC_COLLECTIVE ='O') AND T_messages_ORT.id_rapport = gen.id_rapport)",
                @"UPDATE T_messages_ORT_TRANSFERT_CAPGEMINI SET top_selection='O'  WHERE(date_reception Is null) AND EXISTS(Select 1 FROM T_GENERAL as gen WHERE (PROC_COLLECTIVE ='O') AND T_messages_ORT_TRANSFERT_CAPGEMINI.id_rapport = gen.id_rapport)",
                @"UPDATE T_messages_ORT SET top_selection='N' WHERE(date_reception Is null) AND cast(anc_cotation as varchar) = cast(new_cotation as varchar) AND ( EXISTS(Select 1 FROM T_GENERAL as gen WHERE (PROC_COLLECTIVE ='N' OR PROC_COLLECTIVE IS NULL) AND T_messages_ORT.id_rapport = gen.id_rapport) AND EXISTS(Select 1 FROM T_GENERAL as gen WHERE (NB_PRIVILEGE =0 OR NB_PRIVILEGE IS NULL) AND (NB_PRIVILEGE_TRESOR=0 OR NB_PRIVILEGE_TRESOR IS NULL)            AND T_messages_ORT.id_rapport = gen.id_rapport) AND NOT EXISTS (SELECT 1 FROM T_BODACC as Bodac WHERE lower(ref_evenement) = 'modification sur les représentants' AND datediff(month,date_parution,getdate())<=6              AND T_messages_ORT.id_rapport = Bodac.id_rapport)    )",
                @"UPDATE T_messages_ORT_TRANSFERT_CAPGEMINI SET top_selection='N' WHERE(date_reception Is null) AND cast(anc_cotation as varchar) = cast(new_cotation as varchar) AND ( EXISTS(Select 1 FROM T_GENERAL as gen WHERE (PROC_COLLECTIVE ='N' OR PROC_COLLECTIVE IS NULL) AND T_messages_ORT_TRANSFERT_CAPGEMINI.id_rapport = gen.id_rapport) AND EXISTS(Select 1 FROM T_GENERAL as gen WHERE (NB_PRIVILEGE =0 OR NB_PRIVILEGE IS NULL) AND (NB_PRIVILEGE_TRESOR=0 OR NB_PRIVILEGE_TRESOR IS NULL)            AND T_messages_ORT_TRANSFERT_CAPGEMINI.id_rapport = gen.id_rapport) AND NOT EXISTS (SELECT 1 FROM T_BODACC as Bodac WHERE lower(ref_evenement) = 'modification sur les représentants' AND datediff(month,date_parution,getdate())<=6              AND T_messages_ORT_TRANSFERT_CAPGEMINI.id_rapport = Bodac.id_rapport)    )",
                @"UPDATE T_messages_ORT SET top_selection='N' WHERE(date_reception Is null) AND (new_cotation<>'NA' AND anc_cotation<>'NA') AND cast(anc_cotation as integer) > cast(new_cotation as integer) AND cast(new_cotation as integer)>=7 AND ( EXISTS(Select 1 FROM T_GENERAL as gen WHERE (PROC_COLLECTIVE ='N' OR PROC_COLLECTIVE IS NULL) AND T_messages_ORT.id_rapport = gen.id_rapport) AND EXISTS(Select 1 FROM T_GENERAL as gen WHERE (NB_PRIVILEGE =0 OR NB_PRIVILEGE IS NULL) AND( NB_PRIVILEGE_TRESOR=0 OR NB_PRIVILEGE_TRESOR IS NULL)            AND T_messages_ORT.id_rapport = gen.id_rapport) AND NOT EXISTS (SELECT 1 FROM T_BODACC as Bodac WHERE lower(ref_evenement) = 'modification sur les représentants' AND datediff(month,date_parution,getdate())<=6              AND T_messages_ORT.id_rapport = Bodac.id_rapport) AND NOT EXISTS (SELECT 1 FROM T_BODACC as Bodac WHERE lower(ref_evenement) = 'modification de l''adresse de l''établissement principal'              AND datediff(month,date_parution,getdate())<=6  AND T_messages_ORT.id_rapport = Bodac.id_rapport)    )",
                @"UPDATE T_messages_ORT_TRANSFERT_CAPGEMINI SET top_selection='N' WHERE(date_reception Is null) AND (new_cotation<>'NA' AND anc_cotation<>'NA') AND cast(anc_cotation as integer) > cast(new_cotation as integer) AND cast(new_cotation as integer)>=7 AND ( EXISTS(Select 1 FROM T_GENERAL as gen WHERE (PROC_COLLECTIVE ='N' OR PROC_COLLECTIVE IS NULL) AND T_messages_ORT_TRANSFERT_CAPGEMINI.id_rapport = gen.id_rapport) AND EXISTS(Select 1 FROM T_GENERAL as gen WHERE (NB_PRIVILEGE =0 OR NB_PRIVILEGE IS NULL) AND( NB_PRIVILEGE_TRESOR=0 OR NB_PRIVILEGE_TRESOR IS NULL)            AND T_messages_ORT_TRANSFERT_CAPGEMINI.id_rapport = gen.id_rapport) AND NOT EXISTS (SELECT 1 FROM T_BODACC as Bodac WHERE lower(ref_evenement) = 'modification sur les représentants' AND datediff(month,date_parution,getdate())<=6              AND T_messages_ORT_TRANSFERT_CAPGEMINI.id_rapport = Bodac.id_rapport) AND NOT EXISTS (SELECT 1 FROM T_BODACC as Bodac WHERE lower(ref_evenement) = 'modification de l''adresse de l''établissement principal'              AND datediff(month,date_parution,getdate())<=6  AND T_messages_ORT_TRANSFERT_CAPGEMINI.id_rapport = Bodac.id_rapport)    )",
                @"UPDATE T_messages_ORT SET top_selection='O' WHERE cast(anc_cotation as integer) > cast(new_cotation as integer) AND cast(new_cotation as integer)<=6 AND date_reception is null AND (new_cotation<>'NA' AND anc_cotation<>'NA')",
                @"UPDATE T_messages_ORT_TRANSFERT_CAPGEMINI SET top_selection='O' WHERE cast(anc_cotation as integer) > cast(new_cotation as integer) AND cast(new_cotation as integer)<=6 AND date_reception is null AND (new_cotation<>'NA' AND anc_cotation<>'NA')",
                @"UPDATE T_messages_ORT SET top_selection='O' WHERE  ( (new_cotation<>'NA' AND cast(new_cotation as integer) = 0) OR new_cotation = 'NA') AND date_reception is null",
                @"UPDATE T_messages_ORT_TRANSFERT_CAPGEMINI SET top_selection='O' WHERE  ( (new_cotation<>'NA' AND cast(new_cotation as integer) = 0) OR new_cotation = 'NA') AND date_reception is null",
                @"UPDATE T_messages_ORT SET client =(select TOP 1 dss.branch_customer_nbr from openrowset('SQLOLEDB', {OPENROWSET}, 'Select credit_limit,customer_location.branch_customer_nbr,tax_exempt_nbr FROM  customer_location  INNER JOIN customer ON customer_location.branch_nbr = dbo.customer.branch_nbr AND customer_location.customer_nbr = customer.customer_nbr WHERE (customer_location.suffix = ''000'') AND (customer_location.branch_nbr = ''21'' and StatusCustFlg <> ''D'')') dss where dss.tax_exempt_nbr like '%'+T_messages_ORT.num_siren+'%' order by credit_limit desc) WHERE top_selection='O' AND date_reception is null",
                @"UPDATE T_messages_ORT_TRANSFERT_CAPGEMINI SET client =(select TOP 1 dss.branch_customer_nbr from openrowset('SQLOLEDB', {OPENROWSET}, 'Select credit_limit,customer_location.branch_customer_nbr,tax_exempt_nbr FROM  customer_location  INNER JOIN customer ON customer_location.branch_nbr = dbo.customer.branch_nbr AND customer_location.customer_nbr = customer.customer_nbr WHERE (customer_location.suffix = ''000'') AND (customer_location.branch_nbr = ''21'' and StatusCustFlg <> ''D'')') dss where dss.tax_exempt_nbr like '%'+T_messages_ORT_TRANSFERT_CAPGEMINI.num_siren+'%' order by credit_limit desc) WHERE top_selection='O' AND date_reception is null",
                @"UPDATE T_messages_ORT SET top_selection='N' WHERE (client is null OR client ='' OR top_selection is null) AND date_reception is null",
                @"UPDATE T_messages_ORT_TRANSFERT_CAPGEMINI SET top_selection='N' WHERE (client is null OR client ='' OR top_selection is null) AND date_reception is null",
                @"update T_messages_ORT set credit_limit =(select dss.credit_limit from openrowset('SQLOLEDB', {OPENROWSET}, 'Select branch_customer_nbr,credit_limit FROM  customer where CompanyCd = ''FR''') dss where dss.branch_customer_nbr=T_messages_ORT.client) where top_selection='O' and date_reception is null",
                @"update T_messages_ORT_TRANSFERT_CAPGEMINI set credit_limit =(select dss.credit_limit from openrowset('SQLOLEDB', {OPENROWSET}, 'Select branch_customer_nbr,credit_limit FROM  customer where CompanyCd = ''FR''') dss where dss.branch_customer_nbr=T_messages_ORT_TRANSFERT_CAPGEMINI.client) where top_selection='O' and date_reception is null",
                @"update T_messages_ORT set credit_limit=0 where credit_limit is null and date_reception is null",
                @"update T_messages_ORT_TRANSFERT_CAPGEMINI set credit_limit=0 where credit_limit is null and date_reception is null",
                @"delete from T_calcul_note where indice in (select indice from T_messages_ORT where date_reception is null)",
                @"delete from T_calcul_note_TRANSFERT_CAPGEMINI where indice in (select indice from T_messages_ORT_TRANSFERT_CAPGEMINI  where date_reception is null)",
                @"UPDATE T_messages_ORT set top_selection='2' where date_reception is null AND (top_selection='O' and ((new_cotation<>'NA' AND cast(new_cotation as integer) = 0) OR new_cotation = 'NA')) OR (top_selection='O' AND EXISTS(Select 1 FROM T_GENERAL as gen WHERE PROC_COLLECTIVE ='O' AND T_messages_ORT.id_rapport = gen.id_rapport))",
                @"UPDATE T_messages_ORT_TRANSFERT_CAPGEMINI set top_selection='2' where date_reception is null AND (top_selection='O' and ((new_cotation<>'NA' AND cast(new_cotation as integer) = 0) OR new_cotation = 'NA')) OR (top_selection='O' AND EXISTS(Select 1 FROM T_GENERAL as gen WHERE PROC_COLLECTIVE ='O' AND T_messages_ORT_TRANSFERT_CAPGEMINI.id_rapport = gen.id_rapport))",
                @"update T_messages_ORT set top_selection='1' where top_selection='O' AND date_reception is null AND credit_limit>1 AND new_cotation<>'NA' AND cast(new_cotation as integer) <> 0 ",
                @"update T_messages_ORT_TRANSFERT_CAPGEMINI set top_selection='1' where top_selection='O' AND date_reception is null AND credit_limit>1 AND new_cotation<>'NA' AND cast(new_cotation as integer) <> 0 ",
                @"UPDATE T_messages_ORT SET top_selection='1' WHERE(date_reception Is null) AND top_selection='O' AND credit_limit>1 AND (( EXISTS(Select 1 FROM T_GENERAL as gen WHERE (NB_PRIVILEGE >0 OR NB_PRIVILEGE_TRESOR>0) AND T_messages_ORT.id_rapport = gen.id_rapport)   OR ( EXISTS (SELECT 1 FROM T_BODACC as Bodac WHERE lower(ref_evenement) = 'modification sur les représentants' AND datediff(month,date_parution,getdate())<=6              AND T_messages_ORT.id_rapport = Bodac.id_rapport)      AND  EXISTS (SELECT 1 FROM T_BODACC as Bodac WHERE lower(ref_evenement) = 'modification de l''adresse de l''établissement principal'              AND datediff(month,date_parution,getdate())<=6  AND T_messages_ORT.id_rapport = Bodac.id_rapport)      )   OR EXISTS(Select 1 FROM T_GENERAL as gen WHERE (PROC_COLLECTIVE ='O') AND T_messages_ORT.id_rapport = gen.id_rapport)    ))",
                @"UPDATE T_messages_ORT_TRANSFERT_CAPGEMINI SET top_selection='1' WHERE(date_reception Is null) AND top_selection='O' AND credit_limit>1 AND (( EXISTS(Select 1 FROM T_GENERAL as gen WHERE (NB_PRIVILEGE >0 OR NB_PRIVILEGE_TRESOR>0) AND T_messages_ORT_TRANSFERT_CAPGEMINI.id_rapport = gen.id_rapport)   OR ( EXISTS (SELECT 1 FROM T_BODACC as Bodac WHERE lower(ref_evenement) = 'modification sur les représentants' AND datediff(month,date_parution,getdate())<=6              AND T_messages_ORT_TRANSFERT_CAPGEMINI.id_rapport = Bodac.id_rapport)      AND  EXISTS (SELECT 1 FROM T_BODACC as Bodac WHERE lower(ref_evenement) = 'modification de l''adresse de l''établissement principal'              AND datediff(month,date_parution,getdate())<=6  AND T_messages_ORT_TRANSFERT_CAPGEMINI.id_rapport = Bodac.id_rapport)      )   OR EXISTS(Select 1 FROM T_GENERAL as gen WHERE (PROC_COLLECTIVE ='O') AND T_messages_ORT_TRANSFERT_CAPGEMINI.id_rapport = gen.id_rapport)    ))",
                @"UPDATE T_messages_ORT set montant_balance =(select dss.TotalBalanceAmt from openrowset('SQLOLEDB', {OPENROWSET}, 'Select branch_customer_nbr,TotalBalanceAmt FROM  customer') dss where dss.branch_customer_nbr=T_messages_ORT.client) where (top_selection='1' OR top_selection='2') and date_reception is null",
                @"UPDATE T_messages_ORT_TRANSFERT_CAPGEMINI set montant_balance =(select dss.TotalBalanceAmt from openrowset('SQLOLEDB', {OPENROWSET}, 'Select branch_customer_nbr,TotalBalanceAmt FROM  customer') dss where dss.branch_customer_nbr=T_messages_ORT_TRANSFERT_CAPGEMINI.client) where (top_selection='1' OR top_selection='2') and date_reception is null",
                @"UPDATE T_messages_ORT set montant_balance = 0 WHERE montant_balance is null AND date_reception is null AND (top_selection='1' OR top_selection='2')",
                @"UPDATE T_messages_ORT_TRANSFERT_CAPGEMINI set montant_balance = 0 WHERE montant_balance is null AND date_reception is null AND (top_selection='1' OR top_selection='2')",
                @"UPDATE T_messages_ORT set note=note+3 where top_selection='1' AND date_reception is null AND ( (new_cotation<>'NA' AND cast(new_cotation as integer) <=3) OR new_cotation ='NA')",
                @"UPDATE T_messages_ORT_TRANSFERT_CAPGEMINI set note=note+3 where top_selection='1' AND date_reception is null AND ( (new_cotation<>'NA' AND cast(new_cotation as integer) <=3) OR new_cotation ='NA')",
                @"insert into T_calcul_note select indice,'COTATION',3 from T_messages_ORT  where top_selection='1' AND date_reception is null AND ( (new_cotation<>'NA' AND cast(new_cotation as integer) <=3) OR new_cotation ='NA')",
                @"insert into T_calcul_note_TRANSFERT_CAPGEMINI select indice,'COTATION',3 from T_messages_ORT_TRANSFERT_CAPGEMINI  where top_selection='1' AND date_reception is null AND ( (new_cotation<>'NA' AND cast(new_cotation as integer) <=3) OR new_cotation ='NA')",
                @"UPDATE T_messages_ORT set note=note+2 where top_selection='1' AND date_reception is null AND cast(new_cotation as integer) >3 AND cast(new_cotation as integer) <=4 AND new_cotation<>'NA'",
                @"UPDATE T_messages_ORT_TRANSFERT_CAPGEMINI set note=note+2 where top_selection='1' AND date_reception is null AND cast(new_cotation as integer) >3 AND cast(new_cotation as integer) <=4 AND new_cotation<>'NA'",
                @"insert into T_calcul_note select indice,'COTATION',2 from T_messages_ORT  where top_selection='1' AND date_reception is null AND cast(new_cotation as integer) >3 AND cast(new_cotation as integer) <=4 AND new_cotation<>'NA'",
                @"insert into T_calcul_note_TRANSFERT_CAPGEMINI select indice,'COTATION',2 from T_messages_ORT_TRANSFERT_CAPGEMINI  where top_selection='1' AND date_reception is null AND cast(new_cotation as integer) >3 AND cast(new_cotation as integer) <=4 AND new_cotation<>'NA'",
                @"update T_messages_ORT set note=note+1 where top_selection='1' and cast(new_cotation as integer) >4 and cast(new_cotation as integer) <7 AND new_cotation<>'NA' and date_reception is null",
                @"update T_messages_ORT_TRANSFERT_CAPGEMINI set note=note+1 where top_selection='1' and cast(new_cotation as integer) >4 and cast(new_cotation as integer) <7 AND new_cotation<>'NA' and date_reception is null",
                @"insert into T_calcul_note select indice,'COTATION',1 from T_messages_ORT  where top_selection='1' and cast(new_cotation as integer) >4 and cast(new_cotation as integer) <7 AND new_cotation<>'NA' and date_reception is null",
                @"insert into T_calcul_note_TRANSFERT_CAPGEMINI select indice,'COTATION',1 from T_messages_ORT_TRANSFERT_CAPGEMINI  where top_selection='1' and cast(new_cotation as integer) >4 and cast(new_cotation as integer) <7 AND new_cotation<>'NA' and date_reception is null",
                @"update T_messages_ORT set note=note+2 where valeur_score>=6  and date_reception is null and  top_selection='1' ",
                @"update T_messages_ORT_TRANSFERT_CAPGEMINI set note=note+2 where valeur_score>=6  and date_reception is null and  top_selection='1' ",
                @"insert into T_calcul_note select indice,'SCORE',2 from T_messages_ORT  where valeur_score>=6  and date_reception is null and  top_selection='1'",
                @"insert into T_calcul_note_TRANSFERT_CAPGEMINI select indice,'SCORE',2 from T_messages_ORT_TRANSFERT_CAPGEMINI  where valeur_score>=6  and date_reception is null and  top_selection='1'",
                @"update T_messages_ORT set note=note+1 where valeur_score>=3 and valeur_score<6 and date_reception is null and  top_selection='1' ",
                @"update T_messages_ORT_TRANSFERT_CAPGEMINI set note=note+1 where valeur_score>=3 and valeur_score<6 and date_reception is null and  top_selection='1' ",
                @"insert into T_calcul_note select indice,'SCORE',1 from T_messages_ORT  where valeur_score>=3 and valeur_score<6 and date_reception is null and  top_selection='1' ",
                @"insert into T_calcul_note_TRANSFERT_CAPGEMINI select indice,'SCORE',1 from T_messages_ORT_TRANSFERT_CAPGEMINI  where valeur_score>=3 and valeur_score<6 and date_reception is null and  top_selection='1' ",
                @"update T_messages_ORT set note=note+1 where exists (select TOP 1 DECISIONNING_CREDIT from T_score INNER JOIN t_general ON T_score.ID_RAPPORT = T_General.ID_RAPPORT where DECISIONNING_CREDIT not like 'Accord|%' and T_general.S_SIREN = T_messages_ORT.num_siren ORDER BY T_score.date_score DESC) and date_reception is null and  top_selection='1' ",
                @"update T_messages_ORT_TRANSFERT_CAPGEMINI set note=note+1 where exists (select TOP 1 DECISIONNING_CREDIT from T_score INNER JOIN t_general ON T_score.ID_RAPPORT = T_General.ID_RAPPORT where DECISIONNING_CREDIT not like 'Accord|%' and T_general.S_SIREN = T_messages_ORT_TRANSFERT_CAPGEMINI.num_siren ORDER BY T_score.date_score DESC) and date_reception is null and  top_selection='1' ",
                @"insert into T_calcul_note select indice,'DECISIONNING',1 from T_messages_ORT  where exists (select TOP 1 DECISIONNING_CREDIT from T_score INNER JOIN t_general ON T_score.ID_RAPPORT = T_General.ID_RAPPORT where DECISIONNING_CREDIT not like 'Accord|%' and T_general.S_SIREN = T_messages_ORT.num_siren ORDER BY T_score.date_score DESC) and date_reception is null and  top_selection='1'",
                @"insert into T_calcul_note_TRANSFERT_CAPGEMINI select indice,'DECISIONNING',1 from T_messages_ORT_TRANSFERT_CAPGEMINI  where exists (select TOP 1 DECISIONNING_CREDIT from T_score INNER JOIN t_general ON T_score.ID_RAPPORT = T_General.ID_RAPPORT where DECISIONNING_CREDIT not like 'Accord|%' and T_general.S_SIREN = T_messages_ORT_TRANSFERT_CAPGEMINI.num_siren ORDER BY T_score.date_score DESC) and date_reception is null and  top_selection='1'",
                @"update T_messages_ORT set potientiel_accorde = (SELECT TOP 1 CASE WHEN CHARINDEX(' ke|', dbo.T_score.DECISIONNING_CREDIT) - CHARINDEX('jusque ', dbo.T_score.DECISIONNING_CREDIT) - 7 <= 0 THEN 0 ELSE SUBSTRING(DECISIONNING_CREDIT, CHARINDEX('jusque ', DECISIONNING_CREDIT) + 7, CHARINDEX(' ke|', DECISIONNING_CREDIT) - CHARINDEX('jusque ', DECISIONNING_CREDIT) - 7) END AS Montant FROM T_score INNER JOIN T_General ON T_score.ID_RAPPORT = T_General.ID_RAPPORT  WHERE T_General.S_SIREN = T_messages_ORT.num_siren and (DECISIONNING_CREDIT IS NOT NULL) AND (DECISIONNING_CREDIT LIKE 'Accord|potentiel jusque %') ORDER BY T_General.id_rapport) where top_selection='1' and date_reception is null",
                @"update T_messages_ORT_TRANSFERT_CAPGEMINI set potientiel_accorde = (SELECT TOP 1 CASE WHEN CHARINDEX(' ke|', dbo.T_score.DECISIONNING_CREDIT) - CHARINDEX('jusque ', dbo.T_score.DECISIONNING_CREDIT) - 7 <= 0 THEN 0 ELSE SUBSTRING(DECISIONNING_CREDIT, CHARINDEX('jusque ', DECISIONNING_CREDIT) + 7, CHARINDEX(' ke|', DECISIONNING_CREDIT) - CHARINDEX('jusque ', DECISIONNING_CREDIT) - 7) END AS Montant FROM T_score INNER JOIN T_General ON T_score.ID_RAPPORT = T_General.ID_RAPPORT  WHERE T_General.S_SIREN = T_messages_ORT_TRANSFERT_CAPGEMINI.num_siren and (DECISIONNING_CREDIT IS NOT NULL) AND (DECISIONNING_CREDIT LIKE 'Accord|potentiel jusque %') ORDER BY T_General.id_rapport) where top_selection='1' and date_reception is null",
                @"update T_messages_ORT set potientiel_accorde = 1000* potientiel_accorde where potientiel_accorde is not null and date_reception is null",
                @"update T_messages_ORT_TRANSFERT_CAPGEMINI set potientiel_accorde = 1000* potientiel_accorde where potientiel_accorde is not null and date_reception is null",
                @"update T_messages_ORT set note=note+2 where top_selection='1' and potientiel_accorde<credit_limit and potientiel_accorde is not null and date_reception is null",
                @"update T_messages_ORT_TRANSFERT_CAPGEMINI set note=note+2 where top_selection='1' and potientiel_accorde<credit_limit and potientiel_accorde is not null and date_reception is null",
                @"insert into T_calcul_note select indice,'DECISIONNING',2 from T_messages_ORT  where top_selection='1' and potientiel_accorde<credit_limit and potientiel_accorde is not null and date_reception is null",
                @"insert into T_calcul_note_TRANSFERT_CAPGEMINI select indice,'DECISIONNING',2 from T_messages_ORT_TRANSFERT_CAPGEMINI  where top_selection='1' and potientiel_accorde<credit_limit and potientiel_accorde is not null and date_reception is null",
                @"update T_messages_ORT set potientiel_accorde=0 where top_selection='1' and potientiel_accorde is null and date_reception is null",
                @"update T_messages_ORT_TRANSFERT_CAPGEMINI set potientiel_accorde=0 where top_selection='1' and potientiel_accorde is null and date_reception is null",
                @"UPDATE T_messages_ORT SET note=note+1  WHERE(date_reception Is null) AND  top_selection='1' AND EXISTS(Select 1 FROM T_GENERAL as gen WHERE (NB_PRIVILEGE >0 OR NB_PRIVILEGE_TRESOR>0) AND T_messages_ORT.id_rapport = gen.id_rapport)",
                @"UPDATE T_messages_ORT_TRANSFERT_CAPGEMINI SET note=note+1  WHERE(date_reception Is null) AND  top_selection='1' AND EXISTS(Select 1 FROM T_GENERAL as gen WHERE (NB_PRIVILEGE >0 OR NB_PRIVILEGE_TRESOR>0) AND T_messages_ORT_TRANSFERT_CAPGEMINI.id_rapport = gen.id_rapport)",
                @"insert into T_calcul_note select indice,'PRESENCE DE PRIVILEGES',1 from T_messages_ORT  where  date_reception is null  and  top_selection='1' ",
                @"insert into T_calcul_note_TRANSFERT_CAPGEMINI select indice,'PRESENCE DE PRIVILEGES',1 from T_messages_ORT_TRANSFERT_CAPGEMINI  where  date_reception is null  and  top_selection='1' ",
                @"update T_messages_ORT set note=note+1 where top_selection='1' and exists (select dss.OpenOrderAmt from openrowset('SQLOLEDB', {OPENROWSET}, 'select branch_customer_nbr,OpenOrderAmt from customer where OpenOrderAmt >0 and branch_nbr=''21''') dss where dss.branch_customer_nbr=T_messages_ORT.client) and date_reception is null",
                @"update T_messages_ORT_TRANSFERT_CAPGEMINI set note=note+1 where top_selection='1' and exists (select dss.OpenOrderAmt from openrowset('SQLOLEDB', {OPENROWSET}, 'select branch_customer_nbr,OpenOrderAmt from customer where OpenOrderAmt >0 and branch_nbr=''21''') dss where dss.branch_customer_nbr=T_messages_ORT_TRANSFERT_CAPGEMINI.client) and date_reception is null",
                @"insert into T_calcul_note select indice,'OPEN ORDER',1 from T_messages_ORT  where top_selection='1' and exists (select dss.OpenOrderAmt from openrowset('SQLOLEDB', {OPENROWSET}, 'select branch_customer_nbr,OpenOrderAmt from customer where OpenOrderAmt >0 and branch_nbr=''21''') dss where dss.branch_customer_nbr=T_messages_ORT.client) and date_reception is null",
                @"insert into T_calcul_note_TRANSFERT_CAPGEMINI select indice,'OPEN ORDER',1 from T_messages_ORT_TRANSFERT_CAPGEMINI  where top_selection='1' and exists (select dss.OpenOrderAmt from openrowset('SQLOLEDB', {OPENROWSET}, 'select branch_customer_nbr,OpenOrderAmt from customer where OpenOrderAmt >0 and branch_nbr=''21''') dss where dss.branch_customer_nbr=T_messages_ORT_TRANSFERT_CAPGEMINI.client) and date_reception is null",
                @"update T_messages_ORT set note=note+2 where top_selection='1' and exists (select dss.past_due_31_60 from openrowset('SQLOLEDB', {OPENROWSET}, 'select branch_customer_nbr,past_due_31_60, past_due_61_90, past_due_91 from customer where (past_due_31_60 > 0 or past_due_61_90 > 0 or past_due_91 > 0 ) and branch_nbr=''21''') dss where dss.branch_customer_nbr=T_messages_ORT.client) and date_reception is null",
                @"update T_messages_ORT_TRANSFERT_CAPGEMINI set note=note+2 where top_selection='1' and exists (select dss.past_due_31_60 from openrowset('SQLOLEDB', {OPENROWSET}, 'select branch_customer_nbr,past_due_31_60, past_due_61_90, past_due_91 from customer where (past_due_31_60 > 0 or past_due_61_90 > 0 or past_due_91 > 0 ) and branch_nbr=''21''') dss where dss.branch_customer_nbr=T_messages_ORT_TRANSFERT_CAPGEMINI.client) and date_reception is null",
                @"insert into T_calcul_note select indice,'BALANCE AGEES IMFR',2 from T_messages_ORT  where top_selection='1' and exists (select dss.past_due_31_60 from openrowset('SQLOLEDB', {OPENROWSET}, 'select branch_customer_nbr,past_due_31_60, past_due_61_90, past_due_91 from customer where (past_due_31_60 > 0 or past_due_61_90 > 0 or past_due_91 > 0 ) and branch_nbr=''21''') dss where dss.branch_customer_nbr=T_messages_ORT.client) and date_reception is null",
                @"insert into T_calcul_note_TRANSFERT_CAPGEMINI select indice,'BALANCE AGEES IMFR',2 from T_messages_ORT_TRANSFERT_CAPGEMINI  where top_selection='1' and exists (select dss.past_due_31_60 from openrowset('SQLOLEDB', {OPENROWSET}, 'select branch_customer_nbr,past_due_31_60, past_due_61_90, past_due_91 from customer where (past_due_31_60 > 0 or past_due_61_90 > 0 or past_due_91 > 0 ) and branch_nbr=''21''') dss where dss.branch_customer_nbr=T_messages_ORT_TRANSFERT_CAPGEMINI.client) and date_reception is null",
                @"update T_messages_ORT set note=note+1 where top_selection='1' and exists (select dss.past_due_16_30 from openrowset('SQLOLEDB', {OPENROWSET}, 'select branch_customer_nbr,past_due_16_30 from customer where past_due_16_30 >0 and branch_nbr=''21''') dss where dss.branch_customer_nbr=T_messages_ORT.client) and date_reception is null",
                @"update T_messages_ORT_TRANSFERT_CAPGEMINI set note=note+1 where top_selection='1' and exists (select dss.past_due_16_30 from openrowset('SQLOLEDB', {OPENROWSET}, 'select branch_customer_nbr,past_due_16_30 from customer where past_due_16_30 >0 and branch_nbr=''21''') dss where dss.branch_customer_nbr=T_messages_ORT_TRANSFERT_CAPGEMINI.client) and date_reception is null",
                @"insert into T_calcul_note select indice,'BALANCE AGEES IMFR',1 from T_messages_ORT  where top_selection='1' and exists (select dss.past_due_16_30 from openrowset('SQLOLEDB', {OPENROWSET}, 'select branch_customer_nbr,past_due_16_30 from customer where past_due_16_30 >0 and branch_nbr=''21''') dss where dss.branch_customer_nbr=T_messages_ORT.client) and date_reception is null",
                @"insert into T_calcul_note_TRANSFERT_CAPGEMINI select indice,'BALANCE AGEES IMFR',1 from T_messages_ORT_TRANSFERT_CAPGEMINI  where top_selection='1' and exists (select dss.past_due_16_30 from openrowset('SQLOLEDB', {OPENROWSET}, 'select branch_customer_nbr,past_due_16_30 from customer where past_due_16_30 >0 and branch_nbr=''21''') dss where dss.branch_customer_nbr=T_messages_ORT_TRANSFERT_CAPGEMINI.client) and date_reception is null",
                @"UPDATE T_messages_ORT SET note=note+2  WHERE date_reception Is null AND  top_selection='1' AND  EXISTS (SELECT 1 FROM T_BODACC as Bodac WHERE lower(ref_evenement) = 'modification sur les représentants' AND datediff(month,date_parution,getdate())<=6              AND T_messages_ORT.id_rapport = Bodac.id_rapport)",
                @"UPDATE T_messages_ORT_TRANSFERT_CAPGEMINI SET note=note+2  WHERE date_reception Is null AND  top_selection='1' AND  EXISTS (SELECT 1 FROM T_BODACC as Bodac WHERE lower(ref_evenement) = 'modification sur les représentants' AND datediff(month,date_parution,getdate())<=6              AND T_messages_ORT_TRANSFERT_CAPGEMINI.id_rapport = Bodac.id_rapport)",
                @"insert into T_calcul_note select indice,'CHANGEMENT REPRESENTANT BODACC',2 from T_messages_ORT  where top_selection='1' and date_reception Is null ",
                @"insert into T_calcul_note_TRANSFERT_CAPGEMINI select indice,'CHANGEMENT REPRESENTANT BODACC',2 from T_messages_ORT_TRANSFERT_CAPGEMINI  where top_selection='1' and date_reception Is null ",
                @"UPDATE T_messages_ORT SET note=note+3  WHERE date_reception Is null AND  top_selection='1' AND EXISTS(Select 1 FROM T_GENERAL as gen WHERE (PROC_COLLECTIVE ='O') AND T_messages_ORT.id_rapport = gen.id_rapport)",
                @"UPDATE T_messages_ORT_TRANSFERT_CAPGEMINI SET note=note+3  WHERE date_reception Is null AND  top_selection='1' AND EXISTS(Select 1 FROM T_GENERAL as gen WHERE (PROC_COLLECTIVE ='O') AND T_messages_ORT_TRANSFERT_CAPGEMINI.id_rapport = gen.id_rapport)",
                @"insert into T_calcul_note select indice,'PROCEDURES COLLECTIVES',3 from T_messages_ORT  where top_selection='1' and date_reception Is null ",
                @"insert into T_calcul_note_TRANSFERT_CAPGEMINI select indice,'PROCEDURES COLLECTIVES',3 from T_messages_ORT_TRANSFERT_CAPGEMINI  where top_selection='1' and date_reception Is null ",
                @"update T_messages_ORT set note=note+1 WHERE top_selection='2' AND credit_limit > 1 AND date_reception is null",
                @"update T_messages_ORT_TRANSFERT_CAPGEMINI set note=note+1 WHERE top_selection='2' AND credit_limit > 1 AND date_reception is null",
                @"insert into T_calcul_note select indice,'CREDIT LIMIT',1 from T_messages_ORT  WHERE top_selection='2'  AND credit_limit > 1 AND date_reception is null",
                @"insert into T_calcul_note_TRANSFERT_CAPGEMINI select indice,'CREDIT LIMIT',1 from T_messages_ORT_TRANSFERT_CAPGEMINI  WHERE top_selection='2'  AND credit_limit > 1 AND date_reception is null",
                @"update T_messages_ORT set note=note+1 where top_selection='2' and exists (select dss.total_past_due from openrowset('SQLOLEDB', {OPENROWSET}, 'select branch_customer_nbr,total_past_due from customer where total_past_due >0 and branch_nbr=''21''') dss where dss.branch_customer_nbr=T_messages_ORT.client) and date_reception is null",
                @"update T_messages_ORT_TRANSFERT_CAPGEMINI set note=note+1 where top_selection='2' and exists (select dss.total_past_due from openrowset('SQLOLEDB', {OPENROWSET}, 'select branch_customer_nbr,total_past_due from customer where total_past_due >0 and branch_nbr=''21''') dss where dss.branch_customer_nbr=T_messages_ORT_TRANSFERT_CAPGEMINI.client) and date_reception is null",
                @"insert into T_calcul_note select indice,'BALANCE AGEES',1 from T_messages_ORT  where top_selection='2' and exists (select dss.total_past_due from openrowset('SQLOLEDB', {OPENROWSET}, 'select branch_customer_nbr,total_past_due from customer where total_past_due >0 and branch_nbr=''21''') dss where dss.branch_customer_nbr=T_messages_ORT.client) and date_reception is null",
                @"insert into T_calcul_note_TRANSFERT_CAPGEMINI select indice,'BALANCE AGEES',1 from T_messages_ORT_TRANSFERT_CAPGEMINI  where top_selection='2' and exists (select dss.total_past_due from openrowset('SQLOLEDB', {OPENROWSET}, 'select branch_customer_nbr,total_past_due from customer where total_past_due >0 and branch_nbr=''21''') dss where dss.branch_customer_nbr=T_messages_ORT_TRANSFERT_CAPGEMINI.client) and date_reception is null",
                @"update T_messages_ORT set note=note+1 where top_selection='2' and exists (select dss.OpenOrderAmt from openrowset('SQLOLEDB', {OPENROWSET}, 'select branch_customer_nbr,OpenOrderAmt from customer where OpenOrderAmt >0 and branch_nbr=''21''') dss where dss.branch_customer_nbr=T_messages_ORT.client) and date_reception is null",
                @"update T_messages_ORT_TRANSFERT_CAPGEMINI set note=note+1 where top_selection='2' and exists (select dss.OpenOrderAmt from openrowset('SQLOLEDB', {OPENROWSET}, 'select branch_customer_nbr,OpenOrderAmt from customer where OpenOrderAmt >0 and branch_nbr=''21''') dss where dss.branch_customer_nbr=T_messages_ORT_TRANSFERT_CAPGEMINI.client) and date_reception is null",
                @"insert into T_calcul_note select indice,'OPEN ORDER',1 from T_messages_ORT  where top_selection='2' and exists (select dss.OpenOrderAmt from openrowset('SQLOLEDB', {OPENROWSET}, 'select branch_customer_nbr,OpenOrderAmt from customer where OpenOrderAmt >0 and branch_nbr=''21''') dss where dss.branch_customer_nbr=T_messages_ORT.client) and date_reception is null",
                @"insert into T_calcul_note_TRANSFERT_CAPGEMINI select indice,'OPEN ORDER',1 from T_messages_ORT_TRANSFERT_CAPGEMINI  where top_selection='2' and exists (select dss.OpenOrderAmt from openrowset('SQLOLEDB', {OPENROWSET}, 'select branch_customer_nbr,OpenOrderAmt from customer where OpenOrderAmt >0 and branch_nbr=''21''') dss where dss.branch_customer_nbr=T_messages_ORT_TRANSFERT_CAPGEMINI.client) and date_reception is null",
                @"UPDATE T_messages_ORT SET pole_analyse = 'SMB' WHERE date_reception Is null AND  top_selection='1'  AND credit_limit <=250000 AND (cast(new_cotation as integer) BETWEEN 1 and 5 OR new_cotation ='NA') AND exists (select dss.branch_customer_nbr from openrowset('SQLOLEDB', {OPENROWSET}, 'Select branch_customer_nbr FROM  customer WHERE analyst_risk in (''SMB'',''EXP'')') dss where dss.branch_customer_nbr=T_messages_ORT.client)",
                @"UPDATE T_messages_ORT SET pole_analyse = 'MG',personne_affectee='15'  WHERE date_reception Is null And top_selection ='1' AND (cast(new_cotation as integer) BETWEEN 1 and 5 OR new_cotation ='NA')  And exists(select dss.branch_customer_nbr from openrowset('SQLOLEDB', {OPENROWSET}, 'Select branch_customer_nbr FROM  customer where analyst_risk in (''XER'')') dss where dss.branch_customer_nbr=T_messages_ORT.client)",
                @"UPDATE T_messages_ORT  SET pole_analyse = 'CF',personne_affectee='13'  WHERE date_reception Is null And top_selection ='1' AND (cast(new_cotation as integer) BETWEEN 1 and 5 OR new_cotation ='NA')  And exists(select dss.branch_customer_nbr from openrowset('SQLOLEDB', {OPENROWSET}, 'Select branch_customer_nbr FROM  customer where analyst_risk in (''GMS'') ') dss where dss.branch_customer_nbr=T_messages_ORT.client)  Or ( (date_reception Is null And  top_selection='1' and credit_limit >250000 and credit_limit <=1000000) AND (cast(new_cotation as integer) BETWEEN 1 and 5 OR new_cotation ='NA') AND (personne_affectee is null or personne_affectee = '')  And exists (select dss.branch_customer_nbr from openrowset('SQLOLEDB', {OPENROWSET}, 'Select branch_customer_nbr FROM  customer  where analyst_risk in(''SMB'',''EXP'')') dss where dss.branch_customer_nbr=T_messages_ORT.client)) ",
                @"UPDATE T_messages_ORT SET pole_analyse = 'TC',personne_affectee='4'  WHERE date_reception Is null AND  top_selection='1' AND (cast(new_cotation as integer) BETWEEN 1 and 5 OR new_cotation ='NA') AND exists (select dss.branch_customer_nbr from openrowset('SQLOLEDB', {OPENROWSET}, 'Select branch_customer_nbr FROM  customer where analyst_risk in(''KEY'')') dss where dss.branch_customer_nbr=T_messages_ORT.client)  OR ( (date_reception Is null AND  top_selection='1' and credit_limit >1000000) AND (cast(new_cotation as integer) BETWEEN 1 and 5 OR new_cotation ='NA') AND (personne_affectee is null or personne_affectee = '') AND exists (select dss.branch_customer_nbr from openrowset('SQLOLEDB', {OPENROWSET}, 'Select branch_customer_nbr FROM  customer  where analyst_risk in(''SMB'',''EXP'')') dss where dss.branch_customer_nbr=T_messages_ORT.client)) ",
                @"UPDATE T_messages_ORT SET pole_analyse = 'IA',personne_affectee='14'  WHERE date_reception Is null AND  top_selection='1' AND (cast(new_cotation as integer) BETWEEN 1 and 5 OR new_cotation ='NA') AND exists (select dss.branch_customer_nbr from openrowset('SQLOLEDB', {OPENROWSET}, 'Select branch_customer_nbr FROM  customer where analyst_risk in (''APP'',''DCP'')') dss where dss.branch_customer_nbr=T_messages_ORT.client)",
                @"UPDATE T_messages_ORT SET pole_analyse = 'PANEURO',personne_affectee='7'  WHERE ( date_reception Is null AND  top_selection='1' AND (cast(new_cotation as integer) BETWEEN 1 and 5 OR new_cotation ='NA') AND (personne_affectee is null or personne_affectee = '') AND exists (select dss.branch_customer_nbr from openrowset('SQLOLEDB', {OPENROWSET}, 'Select branch_customer_nbr FROM  customer  where analyst_risk in(''PAN'')') dss where dss.branch_customer_nbr=T_messages_ORT.client)) ",
                @"UPDATE T_messages_ORT SET pole_analyse = 'ALERTE CONCURRENT',personne_affectee='12'  WHERE ((date_reception Is null AND  top_selection='1') AND (cast(new_cotation as integer) BETWEEN 1 and 5 OR new_cotation ='NA') AND (personne_affectee is null or personne_affectee = '') AND exists (select dss.branch_customer_nbr from openrowset('SQLOLEDB', {OPENROWSET}, 'Select branch_customer_nbr FROM  customer  where analyst_risk in(''CON'')') dss where dss.branch_customer_nbr=T_messages_ORT.client)) ",
                @"update T_messages_ORT set pole_analyse = 'ALERTE ELLISPHERE',personne_affectee='9' where top_selection='2' and date_reception is null",
                @"UPDATE T_messages_ORT SET  pole_analyse = 'ALERTE FRAUDE',personne_affectee='10'  WHERE (date_reception Is null AND  top_selection='1'and credit_limit <=50000)  AND EXISTS (SELECT 1 FROM T_BODACC as Bodac WHERE lower(ref_evenement) = 'modification sur les représentants' AND datediff(month,date_parution,getdate())<=6  AND T_messages_ORT.id_rapport = Bodac.id_rapport) AND EXISTS (SELECT 1 FROM T_BODACC as Bodac WHERE lower(ref_evenement) = 'modification de l''adresse de l''établissement principal'  AND datediff(month,date_parution,getdate())<=6 AND T_messages_ORT.id_rapport = Bodac.id_rapport)",
                @"UPDATE T_messages_ORT SET  pole_analyse = 'ALERTE REPRESENTANT',personne_affectee='11'  WHERE (date_reception Is null AND  top_selection='1' and credit_limit >50000)  AND EXISTS (SELECT 1 FROM T_BODACC as Bodac WHERE lower(ref_evenement) = 'modification sur les représentants' AND datediff(month,date_parution,getdate())<=6  AND T_messages_ORT.id_rapport = Bodac.id_rapport)",
                @"UPDATE T_messages_ORT SET pole_analyse = 'SMB' WHERE date_reception Is null AND  top_selection='1' and credit_limit <=250000 AND (cast(new_cotation as integer) BETWEEN 1 and 5 OR new_cotation ='NA') AND (personne_affectee is null or personne_affectee = '') ",
                @"UPDATE T_messages_ORT_TRANSFERT_CAPGEMINI SET pole_analyse = 'ALERTE CONCURRENT',personne_affectee='12'  WHERE ((date_reception Is null AND  top_selection='1') AND (cast(new_cotation as integer) BETWEEN 1 and 5 OR new_cotation ='NA') AND (personne_affectee is null or personne_affectee = '') AND exists (select dss.branch_customer_nbr from openrowset('SQLOLEDB', {OPENROWSET}, 'Select branch_customer_nbr FROM  customer  where analyst_risk in(''CON'')') dss where dss.branch_customer_nbr=T_messages_ORT_TRANSFERT_CAPGEMINI.client)) ",
                @"update T_messages_ORT_TRANSFERT_CAPGEMINI set pole_analyse = 'ALERTE ELLISPHERE',personne_affectee='9' where top_selection='2' and date_reception is null",
                @"UPDATE T_messages_ORT_TRANSFERT_CAPGEMINI SET  pole_analyse = 'ALERTE FRAUDE',personne_affectee='10'  WHERE (date_reception Is null AND  top_selection='1'and credit_limit <=50000)  AND EXISTS (SELECT 1 FROM T_BODACC as Bodac WHERE lower(ref_evenement) = 'modification sur les représentants' AND datediff(month,date_parution,getdate())<=6  AND T_messages_ORT_TRANSFERT_CAPGEMINI.id_rapport = Bodac.id_rapport) AND EXISTS (SELECT 1 FROM T_BODACC as Bodac WHERE lower(ref_evenement) = 'modification de l''adresse de l''établissement principal'  AND datediff(month,date_parution,getdate())<=6 AND T_messages_ORT_TRANSFERT_CAPGEMINI.id_rapport = Bodac.id_rapport)",
                @"UPDATE T_messages_ORT_TRANSFERT_CAPGEMINI SET  pole_analyse = 'ALERTE REPRESENTANT',personne_affectee='11'  WHERE (date_reception Is null AND  top_selection='1' and credit_limit >50000)  AND EXISTS (SELECT 1 FROM T_BODACC as Bodac WHERE lower(ref_evenement) = 'modification sur les représentants' AND datediff(month,date_parution,getdate())<=6  AND T_messages_ORT_TRANSFERT_CAPGEMINI.id_rapport = Bodac.id_rapport)",
                @"UPDATE T_messages_ORT_TRANSFERT_CAPGEMINI SET pole_analyse = 'CAP GEMINI',personne_affectee='1'  WHERE date_reception Is null And top_selection ='1' AND (cast(new_cotation as integer) BETWEEN 1 and 5 OR new_cotation ='NA')  And exists(select dss.branch_customer_nbr from openrowset('SQLOLEDB', {OPENROWSET}, 'Select branch_customer_nbr FROM  customer where analyst_risk in (''CAP'',''CDC'',''MOO'')') dss where dss.branch_customer_nbr=T_messages_ORT_TRANSFERT_CAPGEMINI.client)",
                @"UPDATE T_messages_ORT_TRANSFERT_CAPGEMINI SET pole_analyse = 'PANEURO',personne_affectee='7'  WHERE ( date_reception Is null AND  top_selection='1' AND (cast(new_cotation as integer) BETWEEN 1 and 5 OR new_cotation ='NA') AND (personne_affectee is null or personne_affectee = '') AND exists (select dss.branch_customer_nbr from openrowset('SQLOLEDB', {OPENROWSET}, 'Select branch_customer_nbr FROM  customer  where analyst_risk in (''PAN'')') dss where dss.branch_customer_nbr=T_messages_ORT_TRANSFERT_CAPGEMINI.client)) ",
                @"UPDATE T_messages_ORT_TRANSFERT_CAPGEMINI  SET pole_analyse = 'CF',personne_affectee='13'  WHERE date_reception Is null And top_selection ='1' AND (cast(new_cotation as integer) BETWEEN 1 and 5 OR new_cotation ='NA')  And exists(select dss.branch_customer_nbr from openrowset('SQLOLEDB', {OPENROWSET}, 'Select branch_customer_nbr FROM  customer where analyst_risk in (''CS1'',''RET'') ') dss where dss.branch_customer_nbr=T_messages_ORT_TRANSFERT_CAPGEMINI.client) ",
                @"UPDATE T_messages_ORT_TRANSFERT_CAPGEMINI SET pole_analyse = 'TC',personne_affectee='4'  WHERE date_reception Is null AND  top_selection='1' AND (cast(new_cotation as integer) BETWEEN 1 and 5 OR new_cotation ='NA') AND exists (select dss.branch_customer_nbr from openrowset('SQLOLEDB', {OPENROWSET}, 'Select branch_customer_nbr FROM  customer where analyst_risk in (''KEY'')') dss where dss.branch_customer_nbr=T_messages_ORT_TRANSFERT_CAPGEMINI.client) ",
                @"UPDATE T_messages_ORT_TRANSFERT_CAPGEMINI SET pole_analyse = 'IA',personne_affectee='14'  WHERE date_reception Is null AND  top_selection='1' AND (cast(new_cotation as integer) BETWEEN 1 and 5 OR new_cotation ='NA') AND exists (select dss.branch_customer_nbr from openrowset('SQLOLEDB', {OPENROWSET}, 'Select branch_customer_nbr FROM  customer where analyst_risk in (''APP'',''CS2'')') dss where dss.branch_customer_nbr=T_messages_ORT_TRANSFERT_CAPGEMINI.client)",
                @"UPDATE T_messages_ORT_TRANSFERT_CAPGEMINI SET pole_analyse = 'GLOBAL',personne_affectee='99'  WHERE date_reception Is null AND  top_selection='1' AND (cast(new_cotation as integer) BETWEEN 1 and 5 OR new_cotation ='NA') AND exists (select dss.branch_customer_nbr from openrowset('SQLOLEDB', {OPENROWSET}, 'Select branch_customer_nbr FROM  customer where analyst_risk not in (''CAP'',''CDC'',''MOO'',''APP'',''CS2'',''CS1'',''RET'',''KEY'',''PAN'',''COM'',''PER'')') dss where dss.branch_customer_nbr=T_messages_ORT_TRANSFERT_CAPGEMINI.client)"
                };
                foreach (string historicalQuery in historicalQueries)
                    ExecuteHistoricalSql(connection, historicalQuery.Replace("{OPENROWSET}", dss_con_openrowset));
                ApplyHistoricalSmbRoundRobin(connection);
                ExecuteHistoricalSql(connection, "UPDATE dbo.T_messages_ORT SET date_reception=@P WHERE date_reception IS NULL; UPDATE dbo.T_messages_ORT_TRANSFERT_CAPGEMINI SET date_reception=@P WHERE date_reception IS NULL;", new SqlParameter("@P", SqlDbType.Char, 8) { Value = period });
                string month = archiveDate.ToString("yyyyMM", CultureInfo.InvariantCulture);
                string folder = Path.Combine(path_archives, month);
                Directory.CreateDirectory(folder);
                ExecuteHistoricalSql(connection, "INSERT INTO dbo.T_surveillances VALUES(@F,@D,@M);", new SqlParameter("@F", SqlDbType.NVarChar, 500) { Value = fileName }, new SqlParameter("@D", SqlDbType.Char, 8) { Value = archiveDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture) }, new SqlParameter("@M", SqlDbType.VarChar, 6) { Value = month });
                string archivePath = Path.Combine(folder, fileName);
                File.Move(fullPath, archivePath);
                return archivePath;
            }
        }

        private static string ExtractHistoricalScoreFraudPeriod(string fileName)
        {
            string value = Path.GetFileName(fileName) ?? "";
            if (value.StartsWith("surv_score_", StringComparison.OrdinalIgnoreCase)) value = value.Substring("surv_score_".Length);
            if (value.Length < 8) throw new InvalidDataException("Unable to extract ScoreFraud period from " + fileName);
            return value.Substring(0, 8);
        }

        private static string ReadXmlValue(XmlNode node, string xpath)
        {
            return node?.SelectSingleNode(xpath)?.InnerText ?? "";
        }

        private static void InsertHistoricalScoreFraudReport(SqlConnection connection, string table, XmlNode report)
        {
            var values = new List<KeyValuePair<string, string>>
            {
                new KeyValuePair<string,string>("nom_client",ReadXmlValue(report,"officialCompanyName")),
                new KeyValuePair<string,string>("commentaires_ORT",ReadXmlValue(report,"variationScoreMotive")),
                new KeyValuePair<string,string>("num_siren",ReadXmlValue(report,"companyId")),
                new KeyValuePair<string,string>("anc_cotation",ReadXmlValue(report,"previousScore/score")),
                new KeyValuePair<string,string>("new_cotation",ReadXmlValue(report,"actualScore/score")),
                new KeyValuePair<string,string>("avis_credit",ReadXmlValue(report,"lastCreditOpinion/amount")),
                new KeyValuePair<string,string>("new_avis_credit",ReadXmlValue(report,"creditOpinion/amount"))
            }.Where(x => !string.IsNullOrEmpty(x.Value)).ToList();
            if (values.Count == 0) return;
            string cols = string.Join(",", values.Select(x => x.Key));
            string pars = string.Join(",", values.Select((x, i) => "@P" + i));
            using (var cmd = new SqlCommand("INSERT INTO " + table + " (" + cols + ") VALUES (" + pars + ");", connection))
            {
                cmd.CommandTimeout = 300;
                for (int i = 0; i < values.Count; i++) cmd.Parameters.Add("@P" + i, SqlDbType.NVarChar, -1).Value = values[i].Value;
                cmd.ExecuteNonQuery();
            }
        }

        private void EnrichHistoricalScoreFraudRows(SqlConnection connection)
        {
            var rows = new List<Tuple<int, string>>();
            using (var cmd = new SqlCommand("SELECT indice,ISNULL(num_siren,'') FROM dbo.T_messages_ORT WHERE date_reception IS NULL;", connection))
            { cmd.CommandTimeout = 300; using (SqlDataReader r = cmd.ExecuteReader()) while (r.Read()) rows.Add(Tuple.Create(Convert.ToInt32(r[0]), Convert.ToString(r[1]))); }
            foreach (var row in rows)
            {
                string siren = row.Item2 ?? ""; if (siren.Length > 9) siren = siren.Substring(0, 9); if (siren.Length == 0) continue;
                try
                {
                    string response = CallScoreFraudWebService(siren); if (string.IsNullOrWhiteSpace(response)) continue;
                    var xml = new XmlDocument(); xml.LoadXml(response); if (!string.Equals(xml.DocumentElement?.Name, "response", StringComparison.OrdinalIgnoreCase)) continue;
                    string report = xml.SelectSingleNode("//response/id_rapport")?.InnerText ?? ""; string score = xml.SelectSingleNode("//response/score_fraude")?.InnerText ?? "";
                    ExecuteHistoricalSql(connection, "UPDATE dbo.T_messages_ORT SET id_rapport=@R,valeur_score=@S WHERE indice=@I;", new SqlParameter("@R", SqlDbType.VarChar, 100) { Value = report }, new SqlParameter("@S", SqlDbType.VarChar, 100) { Value = score }, new SqlParameter("@I", SqlDbType.Int) { Value = row.Item1 });
                }
                catch (Exception ex) { WriteLog("       ScoreFraud webservice error for SIREN " + siren + " : " + ex.Message); }
            }
        }

        private string CallScoreFraudWebService(string siren)
        {
            string separator = uri_webservice.Contains("?") ? "&" : "?";
            var request = (HttpWebRequest)WebRequest.Create(uri_webservice.Trim() + separator + "siren=" + Uri.EscapeDataString(siren) + "&type=1");
            request.Credentials = CredentialCache.DefaultCredentials; request.Timeout = 120000; request.ReadWriteTimeout = 120000;
            using (var response = (HttpWebResponse)request.GetResponse()) using (Stream stream = response.GetResponseStream())
            { if (stream == null) return ""; using (var reader = new StreamReader(stream)) return reader.ReadToEnd(); }
        }

        private static void ApplyHistoricalSmbRoundRobin(SqlConnection connection)
        {
            var ids = new List<int>();
            using (var cmd = new SqlCommand("SELECT indice FROM dbo.T_messages_ORT WHERE date_reception IS NULL AND pole_analyse='SMB' AND (personne_affectee IS NULL OR personne_affectee='');", connection))
            using (SqlDataReader r = cmd.ExecuteReader()) while (r.Read()) ids.Add(Convert.ToInt32(r[0]));
            for (int i = 0; i < ids.Count; i++) ExecuteHistoricalSql(connection, "UPDATE dbo.T_messages_ORT SET personne_affectee=@P WHERE indice=@I;", new SqlParameter("@P", SqlDbType.VarChar, 10) { Value = (i + 1) % 2 == 0 ? "14" : "15" }, new SqlParameter("@I", SqlDbType.Int) { Value = ids[i] });
        }

        private static void ExecuteHistoricalSql(SqlConnection connection, string sql, params SqlParameter[] parameters)
        {
            using (var cmd = new SqlCommand(sql, connection)) { cmd.CommandTimeout = 300; if (parameters != null && parameters.Length > 0) cmd.Parameters.AddRange(parameters); cmd.ExecuteNonQuery(); }
        }

        private Message GetScoreFraudMessage(string messageId)
        {
            return ExecuteGraphWithRetry(() => graphService.Users[mailboxAddress].Messages[messageId].GetAsync(config =>
            {
                AddImmutableHeader(config.Headers);
                config.QueryParameters.Select = new[] { "id", "subject", "receivedDateTime", "hasAttachments" };
            }).GetAwaiter().GetResult(), "Get ScoreFraud message");
        }

        private AttachmentCollectionResponse GetScoreFraudAttachments(string messageId)
        {
            return ExecuteGraphWithRetry(() => graphService.Users[mailboxAddress].Messages[messageId].Attachments.GetAsync(config =>
            {
                AddImmutableHeader(config.Headers);
                config.QueryParameters.Top = 999;
                // Do not select contentBytes on the generic Attachment collection.
            }).GetAwaiter().GetResult(), "Read ScoreFraud attachments");
        }

        private FileAttachment GetFileAttachmentContent(string messageId, Microsoft.Graph.Models.Attachment attachment)
        {
            if (attachment is FileAttachment loaded && loaded.ContentBytes != null) return loaded;
            if (string.IsNullOrWhiteSpace(attachment?.Id)) return null;
            return ExecuteGraphWithRetry(() => graphService.Users[mailboxAddress].Messages[messageId].Attachments[attachment.Id].GetAsync(config => AddImmutableHeader(config.Headers)).GetAwaiter().GetResult() as FileAttachment, "Download ScoreFraud attachment " + (attachment.Name ?? attachment.Id));
        }

        private void MarkMessageAsReadAndMove(string messageId, string destinationFolderId)
        {
            ExecuteGraphWithRetry(() =>
            {
                graphService.Users[mailboxAddress].Messages[messageId].PatchAsync(new Message { IsRead = true }, config => AddImmutableHeader(config.Headers)).GetAwaiter().GetResult();
                return true;
            }, "Mark ScoreFraud message as read");
            ExecuteGraphWithRetry(() =>
            {
                var request = new Microsoft.Graph.Users.Item.Messages.Item.Move.MovePostRequestBody { DestinationId = destinationFolderId };
                graphService.Users[mailboxAddress].Messages[messageId].Move.PostAsync(request, config => AddImmutableHeader(config.Headers)).GetAwaiter().GetResult();
                return true;
            }, "Move ScoreFraud message to " + outputFolderName);
        }

        private void PermanentlyDeleteMessage(string messageId)
        {
            ExecuteGraphWithRetry(() =>
            {
                graphService.Users[mailboxAddress].Messages[messageId].PermanentDelete.PostAsync(config => AddImmutableHeader(config.Headers)).GetAwaiter().GetResult();
                return true;
            }, "Permanently delete ScoreFraud CSV message");
        }

        private List<Message> GetCandidateMessages(string folderId)
        {
            int top;
            if (!int.TryParse(number_Of_Mails, out top) || top <= 0) top = 10;
            DateTime date = mailboxFilterDate > new DateTime(1900, 1, 1) ? mailboxFilterDate : ParseStartDate();
            MessageCollectionResponse res = ExecuteGraphWithRetry(() => graphService.Users[mailboxAddress].MailFolders[folderId].Messages.GetAsync(q =>
            {
                AddImmutableHeader(q.Headers);
                q.QueryParameters.Top = top;
                q.QueryParameters.Orderby = new[]
                {
                    "receivedDateTime asc"
                }
                ;
                string filter =
                    "receivedDateTime gt " +
                    date.ToUniversalTime().ToString(
                        "yyyy-MM-ddTHH:mm:ss.fffZ",
                        CultureInfo.InvariantCulture);

                // L'ancien appel Read_Email(True) ne lisait que les messages non lus.
                if (mailboxName.Equals(
                        score_fraud_mailbox_name,
                        StringComparison.OrdinalIgnoreCase))
                {
                    filter += " and isRead eq false";
                }

                q.QueryParameters.Filter = filter;
                q.QueryParameters.Select = new[]
                {
                    "id","subject","receivedDateTime","internetMessageId","from"
                }
                ;
                q.QueryParameters.Expand = new[]
                {
                    "singleValueExtendedProperties($filter=id eq '"+EscapeODataString(processed_property_id)+"')"
                }
                ;
            }
            ).GetAwaiter().GetResult(), "List credit messages");
            var byId = (res?.Value ?? new List<Message>()).Where(x => !string.IsNullOrWhiteSpace(x.Id)).ToDictionary(x => x.Id, x => x, StringComparer.Ordinal);
            foreach (string id in GetDeferredMessageIds()) if (!byId.ContainsKey(id)) try
                {
                    Message m = graphService.Users[mailboxAddress].Messages[id].GetAsync(q =>
                    {
                        AddImmutableHeader(q.Headers);
                        q.QueryParameters.Select = new[]
                        {
                        "id","subject","receivedDateTime","internetMessageId","from"
                    }
                        ;
                    }
                    ).GetAwaiter().GetResult();
                    if (m != null) byId[m.Id] = m;
                }
                catch (Exception ex)
                {
                    WriteLog("       Unable to reload pending message " + id + " : " + ex.Message);
                }
            return byId.Values.OrderBy(x => x.ReceivedDateTime).ToList();
        }

        private ProcessingOutcome ProcessCreditMessage(Message email, string outputFolderId)
        {
            string sender = GetSender(email);
            if (!allowedSenders.Contains(sender))
            {
                UpsertProcessingState(email, "I", null, "Sender not allowed");
                return ProcessingOutcome.Ignored;
            }
            string body = email.Body?.Content ?? "";
            if (body.IndexOf("PAIEMENT DE FACTURE", StringComparison.OrdinalIgnoreCase) < 0)
            {
                if (body.IndexOf("reference of the customer concerned: 21000007", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    FinalizeGraphMessage(email.Id, outputFolderId);
                    UpsertProcessingState(email, "S", null, "English test message archived");
                    return ProcessingOutcome.Completed;
                }
                UpsertProcessingState(email, "I", null, "Not a payment invoice message");
                return ProcessingOutcome.Ignored;
            }
            CreditMailData data = ParseCreditMail(email);
            if (string.IsNullOrWhiteSpace(data.AuthorizationNumber))
            {
                FinalizeGraphMessage(email.Id, outputFolderId);
                UpsertProcessingState(email, "S", null, "Authorization number missing - archived without business processing");
                return ProcessingOutcome.Completed;
            }
            if (AuthorizationExists(data.AuthorizationNumber) || data.CustomerCode.Equals("21000007", StringComparison.OrdinalIgnoreCase))
            {
                FinalizeGraphMessage(email.Id, outputFolderId);
                UpsertProcessingState(email, "S", null, "Already processed or test customer");
                return ProcessingOutcome.Completed;
            }
            bool defer;
            string resolved = ResolveDeliveryOrInvoice(data, email.ReceivedDateTime?.LocalDateTime ?? DateTime.Now, out defer);
            if (defer)
            {
                UpsertProcessingState(email, "P", null, "BL/invoice not found during retry window");
                return ProcessingOutcome.Deferred;
            }
            data.DeliveryOrInvoice = resolved;
            byte[] mime = GetMimeContent(email.Id);
            int requestId = InsertCreditRequestAndMail(email, data, mime);
            UpdateCreditManager(requestId);
            UpdateTransactionStatus(requestId, data.TransactionStatus);
            ProcessContentieux(requestId);
            SetFinalStatus(requestId);
            FinalizeGraphMessage(email.Id, outputFolderId);
            UpsertProcessingState(email, "S", requestId, "Completed");
            WriteLog("       Credit request inserted and archived. ID : " + requestId);
            return ProcessingOutcome.Completed;
        }

        private CreditMailData ParseCreditMail(Message email)
        {
            string b = (email.Body?.Content ?? "").Replace("\r\n", "\n");
            var d = new CreditMailData
            {
                CustomerName = ExtractBetween(b, "Nom du client concerné :", "Référence du client concerné").ToUpperInvariant(),
                CustomerCode = ExtractBetween(b, "Référence du client concerné :", "Adresse e-mail du client concerné").ToUpperInvariant(),
                CustomerEmail = ExtractBetween(b, "Adresse e-mail du client concerné :", "Le résultat du paiement").ToUpperInvariant(),
                TransactionStatus = ExtractBetween(b, "énoncées ci-dessous est :", "La demande de paiement"),
                DeliveryOrInvoice = ExtractBetween(b, "NUMERO DE BL :", "Montant TTC :").ToUpperInvariant(),
                AuthorizationNumber = ExtractBetween(b, "Numero d'autorisation :", "Commentaire :").ToUpperInvariant(),
                Comments = ExtractAfter(b, "Commentaire :").ToUpperInvariant(),
                RequestDate = email.CreatedDateTime?.LocalDateTime ?? email.ReceivedDateTime?.LocalDateTime ?? DateTime.Now
            }
            ;
            string a = ExtractBetween(b, "MONTANT TTC :", "EUR").Replace("€", "").Replace("", "").Replace(" ", "").Replace(',', '.');
            if (!double.TryParse(a, NumberStyles.Any, CultureInfo.InvariantCulture, out double amount)) throw new InvalidDataException("Invalid amount : " + a);
            d.Amount = amount;
            if (d.CustomerCode.Length < 2) throw new InvalidDataException("Customer code is missing");
            d.CustomerBranch = d.CustomerCode.Substring(0, 2);
            return d;
        }

        private string ResolveDeliveryOrInvoice(CreditMailData d, DateTime received, out bool defer)
        {
            defer = false;
            string v = (d.DeliveryOrInvoice ?? "").Trim();
            if (d.CustomerBranch.Equals(Left(v, 2), StringComparison.OrdinalIgnoreCase) && v.Length > 12 && Left(v, 12).EndsWith(",")) v = Left(v, 11);
            if (d.CustomerBranch.Equals(Left(v, 2), StringComparison.OrdinalIgnoreCase) && v.Length == 11 && v[2] == '-' && v[8] == '-') v = v.Replace("-", "");
            bool same = d.CustomerBranch.Equals(Left(v, 2), StringComparison.OrdinalIgnoreCase), seven = same && v.Length == 7;
            if (same && !seven) return v;
            string found = FindDeliveryInGestionCdes(d.CustomerCode, v);
            if (!string.IsNullOrWhiteSpace(found)) return found;
            if (seven)
            {
                string inv = FindInvoiceInDss(d.CustomerCode, v);
                if (!string.IsNullOrWhiteSpace(inv)) return inv;
                if ((DateTime.Now - received).TotalHours < retry_delay_hours)
                {
                    defer = true;
                    return "";
                }
                return v;
            }
            if ((DateTime.Now - received).TotalHours < retry_delay_hours)
            {
                defer = true;
                return "";
            }
            return d.CustomerBranch + "00000";
        }

        private string FindDeliveryInGestionCdes(string customer, string order)
        {
            const string sql = @"SELECT TOP (1) CAST(o.BR_NBR AS varchar)+CAST(o.ORDR_NBR AS varchar)+CAST(o.DIST_NBR AS varchar)+CAST(o.SHIP_NBR AS varchar) FROM OPENQUERY(DWHP_IMT,'SELECT s.BR_NBR,s.ORDR_NBR,s.DIST_NBR,s.SHIP_NBR,h.BILL_CUST_NBR,h.CUST_ORDR_NBR FROM DSSDATA.UVW_ORDERHEADFR h INNER JOIN DSSDATA.UVW_ORDERSHIPFR s ON h.BR_NBR=s.BR_NBR AND h.ORDR_NBR=s.ORDR_NBR AND h.ORDR_DATE=s.ORDR_DT') o WHERE o.BILL_CUST_NBR=@C AND o.BR_NBR='21' AND(o.CUST_ORDR_NBR=@O OR CAST(o.BR_NBR AS varchar)+CAST(o.ORDR_NBR AS varchar)=@O);";
            return ExecuteScalarString(sql_gestion_cdes, sql, new SqlParameter("@C", SqlDbType.VarChar, 10)
            {
                Value = Right(customer, 6)
            }
            , new SqlParameter("@O", SqlDbType.VarChar, 50)
            {
                Value = order ?? ""
            }
            );
        }

        private string FindInvoiceInDss(string customer, string order)
        {
            return ExecuteScalarString(sql_dss_copie, "SELECT TOP(1) invoice_nbr FROM dbo.shipment_header WHERE order_nbr=@O AND branch_customer_nbr=@C AND invoice_date>DATEADD(year,-2,GETDATE()) ORDER BY invoice_date DESC;", new SqlParameter("@O", SqlDbType.VarChar, 50)
            {
                Value = order ?? ""
            }
            , new SqlParameter("@C", SqlDbType.VarChar, 10)
            {
                Value = customer ?? ""
            }
            );
        }

        private int InsertCreditRequestAndMail(Message email, CreditMailData d, byte[] mime)
        {
            using (var c = new SqlConnection(sql_gestion_cdes))
            {
                c.Open();
                using (var tx = c.BeginTransaction())
                {
                    try
                    {
                        int id;
                        using (var cmd = new SqlCommand(@"INSERT INTO dbo.T_deblocage_carte_bleue(nom_prenom,date_demande,num_bl,montant_CB,montant_bl,code_client,modalite_reglement,id_statut,email_demandeur,statut_transaction,num_autorisation)VALUES(@N,@D,@B,@M,0,@C,1,0,@E,@S,@A);SELECT CAST(SCOPE_IDENTITY() AS int);", c, tx))
                        {
                            cmd.Parameters.Add("@N", SqlDbType.VarChar, 50).Value = Truncate(d.CustomerName, 50);
                            cmd.Parameters.Add("@D", SqlDbType.DateTime).Value = d.RequestDate;
                            cmd.Parameters.Add("@B", SqlDbType.VarChar, 50).Value = Truncate(d.DeliveryOrInvoice, 50);
                            cmd.Parameters.Add("@M", SqlDbType.Float).Value = d.Amount;
                            cmd.Parameters.Add("@C", SqlDbType.VarChar, 8).Value = Truncate(d.CustomerCode, 8);
                            cmd.Parameters.Add("@E", SqlDbType.VarChar, 100).Value = Truncate(d.CustomerEmail, 100);
                            cmd.Parameters.Add("@S", SqlDbType.VarChar, 500).Value = Truncate(d.TransactionStatus, 500);
                            cmd.Parameters.Add("@A", SqlDbType.VarChar, 500).Value = Truncate(d.AuthorizationNumber, 500);
                            id = Convert.ToInt32(cmd.ExecuteScalar());
                        }
                        using (var cmd = new SqlCommand("dbo.USP_DEBLOCAGE_CB_mail_joint", c, tx))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.Add("@id_ligne", SqlDbType.Int).Value = id;
                            cmd.Parameters.Add("@nom_fichier", SqlDbType.NVarChar, 200).Value = Truncate(CleanFileName(email.Subject ?? "email") + ".eml", 200);
                            cmd.Parameters.Add("@fichier", SqlDbType.Image).Value = mime;
                            cmd.Parameters.Add("@ext", SqlDbType.NVarChar, 5).Value = "eml";
                            cmd.Parameters.Add("@sujet", SqlDbType.NVarChar, -1).Value = email.Subject ?? "";
                            cmd.Parameters.Add("@commentaire", SqlDbType.NVarChar, -1).Value = d.Comments;
                            cmd.ExecuteNonQuery();
                        }
                        tx.Commit();
                        return id;
                    }
                    catch
                    {
                        tx.Rollback();
                        throw;
                    }
                }
            }
        }

        private bool AuthorizationExists(string a)
        {
            return Convert.ToInt32(ExecuteScalarString(sql_gestion_cdes, "SELECT COUNT(*) FROM dbo.T_deblocage_carte_bleue WHERE UPPER(LTRIM(RTRIM(num_autorisation)))=UPPER(LTRIM(RTRIM(@A)));", new SqlParameter("@A", SqlDbType.VarChar, 500)
            {
                Value = a ?? ""
            }
            )) > 0;
        }

        private void UpdateCreditManager(int id)
        {
            ExecuteNonQuery(sql_gestion_cdes, "UPDATE dbo.T_deblocage_carte_bleue SET code_gestionnaire_credit=(SELECT TOP 1 credit_mgr_code FROM DSS_COPIE.dbo.customer WHERE customer_nbr=code_client COLLATE SQL_Latin1_General_CP1_CI_AS)WHERE indice=@I;", new SqlParameter("@I", SqlDbType.Int)
            {
                Value = id
            }
            );
        }

        private void UpdateTransactionStatus(int id, string s)
        {
            if (!string.Equals((s ?? "").Trim(), "Transaction acceptée", StringComparison.OrdinalIgnoreCase)) ExecuteNonQuery(sql_gestion_cdes, "UPDATE dbo.T_deblocage_carte_bleue SET id_statut=8,motif_refus=0 WHERE indice=@I;", new SqlParameter("@I", SqlDbType.Int)
            {
                Value = id
            }
            );
        }

        private void ProcessContentieux(int id)
        {
            string manager = ExecuteScalarString(sql_gestion_cdes, "SELECT ISNULL(code_gestionnaire_credit,'') FROM dbo.T_deblocage_carte_bleue WHERE indice=@I;", new SqlParameter("@I", SqlDbType.Int)
            {
                Value = id
            }
            );
            if (!credit_managers_contentieux_list.Contains(manager)) return;
            ExecuteNonQuery(sql_gestion_cdes, "UPDATE dbo.T_deblocage_carte_bleue SET id_statut=7,motif_refus=0,commentaires_macro='A traiter par le contentieux' WHERE indice=@I;", new SqlParameter("@I", SqlDbType.Int)
            {
                Value = id
            }
            );
            int hist = Convert.ToInt32(ExecuteScalarString(sql_gestion_cdes, "SELECT ISNULL(MAX(id_histo),0) FROM dbo.T_deblocage_carte_bleue_histo_mail WHERE indice_demande=@I;", new SqlParameter("@I", SqlDbType.Int)
            {
                Value = id
            }
            ));
            if (hist <= 0) throw new InvalidOperationException("No email history for request " + id);
            string link = "https://defrizwiis1041/credit/imfr_recouvrement/show_mail_CB_eml.aspx?id_histo=" + hist;
            SendGraphMail(contentieux_recipients, "Reception de Cartes Bleues pour les Credit Managers " + string.Join(",", credit_managers_contentieux_list), "Cliquez <a href='" + WebUtility.HtmlEncode(link) + "'>ici</a> pour consulter le mail.");
        }

        private void SetFinalStatus(int id)
        {
            ExecuteNonQuery(sql_gestion_cdes, "UPDATE dbo.T_deblocage_carte_bleue SET id_statut=1 WHERE indice=@I AND id_statut NOT IN(7,8);", new SqlParameter("@I", SqlDbType.Int)
            {
                Value = id
            }
            );
        }

        private void UpsertProcessingState(Message m, string status, int? request, string error)
        {
            const string sql = @"MERGE dbo.T_CreditMailProcessing t USING(SELECT @M MailboxId,@G GraphMessageId)s ON t.MailboxId=s.MailboxId AND t.GraphMessageId=s.GraphMessageId WHEN MATCHED THEN UPDATE SET InternetMessageId=@II,ReceivedDateTime=@R,CreditRequestId=COALESCE(@CI,t.CreditRequestId),ProcessingStatus=@S,ErrorMessage=@E,ProcessedAtUtc=CASE WHEN @S='S' THEN SYSUTCDATETIME() ELSE t.ProcessedAtUtc END WHEN NOT MATCHED THEN INSERT(MailboxId,GraphMessageId,InternetMessageId,ReceivedDateTime,CreditRequestId,ProcessingStatus,ErrorMessage,ProcessedAtUtc)VALUES(@M,@G,@II,@R,@CI,@S,@E,CASE WHEN @S='S' THEN SYSUTCDATETIME() ELSE NULL END);";
            ExecuteNonQuery(sql_connexion, sql, new SqlParameter("@M", SqlDbType.Int)
            {
                Value = mailboxId
            }
            , new SqlParameter("@G", SqlDbType.NVarChar, 500)
            {
                Value = m.Id ?? ""
            }
            , new SqlParameter("@II", SqlDbType.NVarChar, 500)
            {
                Value = m.InternetMessageId ?? ""
            }
            , new SqlParameter("@R", SqlDbType.DateTime2)
            {
                Value = m.ReceivedDateTime?.UtcDateTime ?? DateTime.UtcNow
            }
            , new SqlParameter("@CI", SqlDbType.Int)
            {
                Value = request.HasValue ? (object)request.Value : DBNull.Value
            }
            , new SqlParameter("@S", SqlDbType.Char, 1)
            {
                Value = status
            }
            , new SqlParameter("@E", SqlDbType.NVarChar, 2000)
            {
                Value = Truncate(error, 2000)
            }
            );
        }

        private List<string> GetDeferredMessageIds()
        {
            var list = new List<string>();
            using (var c = new SqlConnection(sql_connexion)) using (var cmd = new SqlCommand("SELECT GraphMessageId FROM dbo.T_CreditMailProcessing WHERE MailboxId=@I AND ProcessingStatus IN('P','E');", c))
            {
                cmd.Parameters.Add("@I", SqlDbType.Int).Value = mailboxId;
                c.Open();
                using (var r = cmd.ExecuteReader()) while (r.Read()) list.Add(Convert.ToString(r[0]));
            }
            return list;
        }

        private void FinalizeGraphMessage(string id, string dest)
        {
            if (!TryMarkMessageAsProcessed(id)) WriteLog("       Warning: SQL succeeded but Graph tracking failed.");
            ExecuteGraphWithRetry(() =>
            {
                graphService.Users[mailboxAddress].Messages[id].Move.PostAsync(new Microsoft.Graph.Users.Item.Messages.Item.Move.MovePostRequestBody
                {
                    DestinationId = dest
                }
                , q => AddImmutableHeader(q.Headers)).GetAwaiter().GetResult();
                return true;
            }
            , "Move processed credit message");
        }

        private bool TryMarkMessageAsProcessed(string id)
        {
            Exception last = null;
            for (int attempt = 1;
            attempt <= 3;
            attempt++) try
                {
                    graphService.Users[mailboxAddress].Messages[id].PatchAsync(new Message
                    {
                        IsRead = true,
                        SingleValueExtendedProperties = new List<SingleValueLegacyExtendedProperty>
                    {
                        new SingleValueLegacyExtendedProperty
                        {
                            Id=processed_property_id,Value=DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ",CultureInfo.InvariantCulture)
                        }
                    }
                    }
                    , q => AddImmutableHeader(q.Headers)).GetAwaiter().GetResult();
                    return true;
                }
                catch (Exception ex) when (IsChangeKeyConflict(ex))
                {
                    last = ex;
                    if (attempt >= 3) break;
                    System.Threading.Thread.Sleep(attempt * 500);
                }
                catch (Exception ex)
                {
                    last = ex;
                    break;
                }
            SendTechnicalAlert(nameof(TryMarkMessageAsProcessed), GetInnermostExceptionMessage(last), "MESSAGE TRACKING");
            return false;
        }

        private Message GetCompleteMessage(string id)
        {
            return ExecuteGraphWithRetry(() => graphService.Users[mailboxAddress].Messages[id].GetAsync(q =>
            {
                q.Headers.Add("Prefer", "outlook.body-content-type=\"text\", " + immutable_id_preference);
                q.QueryParameters.Select = new[]
                {
                    "id","subject","body","from","sender","receivedDateTime","createdDateTime","internetMessageId"
                }
                ;
            }
            ).GetAwaiter().GetResult(), "Get complete message");
        }

        private byte[] GetMimeContent(string id)
        {
            using (Stream input = ExecuteGraphWithRetry(() => graphService.Users[mailboxAddress].Messages[id].Content.GetAsync(q => AddImmutableHeader(q.Headers)).GetAwaiter().GetResult(), "Get MIME")) using (var output = new MemoryStream())
            {
                input.CopyTo(output);
                return output.ToArray();
            }
        }

        private MailFolder GetRequiredFolder(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Folder name is empty", nameof(name));
            if (name.Equals("Inbox", StringComparison.OrdinalIgnoreCase))
                return ExecuteGraphWithRetry(() => graphService.Users[mailboxAddress].MailFolders["inbox"].GetAsync(q => AddImmutableHeader(q.Headers)).GetAwaiter().GetResult(), "Read Inbox folder");
            MailFolderCollectionResponse res = ExecuteGraphWithRetry(() => graphService.Users[mailboxAddress].MailFolders["inbox"].ChildFolders.GetAsync(q => { AddImmutableHeader(q.Headers); q.QueryParameters.Top = 100; q.QueryParameters.Filter = "displayName eq '" + EscapeODataString(name) + "'"; }).GetAwaiter().GetResult(), "Find folder " + name);
            MailFolder f = res?.Value?.FirstOrDefault(x => string.Equals(x.DisplayName, name, StringComparison.OrdinalIgnoreCase));
            if (f == null) throw new DirectoryNotFoundException("Folder not found under Inbox : " + name);
            return f;
        }

        private T ExecuteGraphWithRetry<T>(Func<T> action, string operation)
        {
            const int maxAttempts = 3;
            Exception last = null;

            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                try
                {
                    return action();
                }
                catch (Exception ex) when (IsTransientGraphError(ex))
                {
                    last = ex;
                    if (attempt >= maxAttempts) break;

                    int delayMilliseconds = attempt * 5000;
                    WriteLog(
                        "       Temporary Graph error during " + operation +
                        ". Application attempt " + attempt + "/" + maxAttempts +
                        ". Retry in " + delayMilliseconds +
                        " ms. Error : " + GetInnermostExceptionMessage(ex));
                    System.Threading.Thread.Sleep(delayMilliseconds);
                }
            }

            throw new InvalidOperationException(
                "Graph operation failed after " + maxAttempts +
                " application attempt(s) : " + operation + " - " +
                GetInnermostExceptionMessage(last),
                last);
        }

        private void SendTechnicalAlert(string method, string error, string type)
        {
            try
            {
                SendGraphMail(email_in_case_of_technical_issue, global_application_name + " - Erreur " + type + " dans " + method, "<b>Mailbox :</b> " + WebUtility.HtmlEncode(mailboxAddress) + "<br/><b>Message :</b> " + WebUtility.HtmlEncode(error));
            }
            catch (Exception ex)
            {
                WriteLog("Alert error : " + ex.Message);
            }
        }

        private void SendGraphMail(string recipients, string subject, string html)
        {
            List<Recipient> to = BuildRecipients(recipients);
            if (to.Count == 0) throw new InvalidOperationException("Invalid recipients");
            var m = new Message
            {
                Subject = subject,
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = html
                }
                ,
                ToRecipients = to
            }
            ;
            graphService.Users[fr_graph_send_as].SendMail.PostAsync(new Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody
            {
                Message = m,
                SaveToSentItems = true
            }
            ).GetAwaiter().GetResult();
        }

        private void UpsertMailboxError(Exception ex)
        {
            WriteLog("   Error reading mailbox " + mailboxAddress + " : " + ex.Message);
            SendTechnicalAlert(nameof(Read_Email_with_Graph), mailboxAddress + " - " + ex.Message, "MAILBOX PROCESSING");
        }

        private void UpdateMailboxRefreshInformation(int id, DateTime? latest)
        {
            ExecuteNonQuery(sql_connexion, "UPDATE dbo.T_SharedMailboxes SET date_dernier_raf=GETDATE(),dt_heure_filtre=CASE WHEN @L IS NULL THEN dt_heure_filtre WHEN dt_heure_filtre IS NULL OR @L>dt_heure_filtre THEN @L ELSE dt_heure_filtre END WHERE id_mailboxe=@I;", new SqlParameter("@L", SqlDbType.DateTime)
            {
                Value = latest.HasValue ? (object)latest.Value : DBNull.Value
            }
            , new SqlParameter("@I", SqlDbType.Int)
            {
                Value = id
            }
            );
        }

        private string GetImcaParameter(string c, string p)
        {
            if (string.IsNullOrWhiteSpace(p)) return "";
            return ExecuteScalarString(c, "SELECT ISNULL(VALUE,'') FROM PCM_TAB_IMCA_PARAMETER_GLOBAL WHERE SK_VALID=0 AND PARAMETER=@P;", new SqlParameter("@P", SqlDbType.NVarChar, 255)
            {
                Value = p
            }
            );
        }

        private static void ExecuteNonQuery(string cs, string sql, params SqlParameter[] ps)
        {
            using (var c = new SqlConnection(cs)) using (var cmd = new SqlCommand(sql, c))
            {
                cmd.CommandTimeout = 300;
                if (ps != null) cmd.Parameters.AddRange(ps);
                c.Open();
                cmd.ExecuteNonQuery();
            }
        }

        private static string ExecuteScalarString(string cs, string sql, params SqlParameter[] ps)
        {
            using (var c = new SqlConnection(cs)) using (var cmd = new SqlCommand(sql, c))
            {
                cmd.CommandTimeout = 300;
                if (ps != null) cmd.Parameters.AddRange(ps);
                c.Open();
                return Convert.ToString(cmd.ExecuteScalar()).Trim();
            }
        }

        private void ValidateCountryConfiguration()
        {
            if (string.IsNullOrWhiteSpace(sql_connexion) ||
                string.IsNullOrWhiteSpace(sql_gestion_cdes) ||
                string.IsNullOrWhiteSpace(sql_dss_copie) ||
                string.IsNullOrWhiteSpace(sql_creation_compte) ||
                string.IsNullOrWhiteSpace(email_in_case_of_technical_issue) ||
                string.IsNullOrWhiteSpace(fr_graph_send_as) ||
                string.IsNullOrWhiteSpace(path_archives) ||
                string.IsNullOrWhiteSpace(dss_con_openrowset) ||
                string.IsNullOrWhiteSpace(uri_webservice))
            {
                throw new InvalidOperationException(
                    "Country configuration is incomplete");
            }
            if (credit_managers_contentieux_list.Count == 0 || BuildRecipients(contentieux_recipients).Count == 0) throw new InvalidOperationException("Contentieux configuration is invalid");
        }

        private void ValidateMailboxConfiguration()
        {
            if (mailboxId <= 0 || string.IsNullOrWhiteSpace(mailboxAddress) || string.IsNullOrWhiteSpace(inputFolderName) || string.IsNullOrWhiteSpace(outputFolderName)) throw new InvalidOperationException("Mailbox configuration is incomplete");
            if (mailboxName.Equals(credit_card_mailbox_name, StringComparison.OrdinalIgnoreCase) && allowedSenders.Count == 0) throw new InvalidOperationException("sender_allowed is empty for the Cartes Bleues mailbox");
        }

        private DateTime ParseStartDate()
        {
            return DateTime.TryParseExact(start_date_scan, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime d) ? d : new DateTime(1900, 1, 1);
        }

        private void WriteLog(string message)
        {
            try
            {
                Directory.CreateDirectory(logsFolder);

                string logPath = Path.Combine(
                    logsFolder,
                    "IMCA_" +
                    sessionName + "_" +
                    DateTime.Now.ToString("dd_MM_yyyy") + "_" +
                    country + "_" +
                    global_application_name +
                    ".txt");

                string line =
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") +
                    " - " +
                    (message ?? "") +
                    Environment.NewLine;

                lock (logSyncRoot)
                {
                    WriteLogLineWithCrossProcessLock(logPath, line);
                }
            }
            catch (Exception ex)
            {
                WriteEmergencyLog(
                    "WriteLog failure" +
                    " - Error : " + ex.Message +
                    " - Original message : " + (message ?? ""));
            }
        }

        private static void WriteLogLineWithCrossProcessLock(
            string logPath,
            string line)
        {
            string mutexName =
                "Local\\IMCA_CREDIT_LOG_" +
                GetStableLogNameHash(logPath);

            using (var mutex =
                new System.Threading.Mutex(false, mutexName))
            {
                bool lockTaken = false;

                try
                {
                    try
                    {
                        lockTaken = mutex.WaitOne(
                            TimeSpan.FromSeconds(10));
                    }
                    catch (System.Threading.AbandonedMutexException)
                    {
                        lockTaken = true;
                    }

                    if (!lockTaken)
                    {
                        throw new IOException(
                            "Unable to acquire the log mutex within 10 seconds");
                    }

                    const int maxAttempts = 5;
                    IOException lastWriteException = null;

                    for (int attempt = 1;
                         attempt <= maxAttempts;
                         attempt++)
                    {
                        try
                        {
                            byte[] content = Encoding.UTF8.GetBytes(line);

                            using (var stream = new FileStream(
                                logPath,
                                FileMode.Append,
                                FileAccess.Write,
                                FileShare.ReadWrite))
                            {
                                stream.Write(content, 0, content.Length);
                                stream.Flush();
                            }

                            return;
                        }
                        catch (IOException ex)
                        {
                            lastWriteException = ex;

                            if (attempt < maxAttempts)
                            {
                                System.Threading.Thread.Sleep(attempt * 100);
                            }
                        }
                    }

                    throw new IOException(
                        "Unable to write the log file after " +
                        maxAttempts +
                        " attempts : " +
                        logPath,
                        lastWriteException);
                }
                finally
                {
                    if (lockTaken)
                    {
                        try
                        {
                            mutex.ReleaseMutex();
                        }
                        catch (ApplicationException)
                        {
                        }
                    }
                }
            }
        }

        private void WriteEmergencyLog(string message)
        {
            try
            {
                string emergencyFolder =
                    string.IsNullOrWhiteSpace(tempFolder)
                        ? AppDomain.CurrentDomain.BaseDirectory
                        : tempFolder;

                Directory.CreateDirectory(emergencyFolder);

                string emergencyFile = Path.Combine(
                    emergencyFolder,
                    "IMCA_LOG_FAILURE_" +
                    global_application_name +
                    "_ACTION_ID_" +
                    System.Diagnostics.Process
                        .GetCurrentProcess()
                        .Id +
                    "_" +
                    DateTime.Now.ToString("dd_MM_yyyy") +
                    ".txt");

                string line =
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") +
                    " - " +
                    (message ?? "") +
                    Environment.NewLine;

                byte[] content = Encoding.UTF8.GetBytes(line);

                using (var stream = new FileStream(
                    emergencyFile,
                    FileMode.Append,
                    FileAccess.Write,
                    FileShare.ReadWrite))
                {
                    stream.Write(content, 0, content.Length);
                    stream.Flush();
                }
            }
            catch
            {
                // Last-resort protection: logging must never stop processing.
            }
        }

        private static string GetStableLogNameHash(string value)
        {
            unchecked
            {
                uint hash = 2166136261;

                foreach (char character in
                    (value ?? "").ToUpperInvariant())
                {
                    hash ^= character;
                    hash *= 16777619;
                }

                return hash.ToString("X8", CultureInfo.InvariantCulture);
            }
        }

        private static IEnumerable<string> SplitQuotedValues(string v)
        {
            return (v ?? "").Split(new[]
            {
                ';',','
            }
            , StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim().Trim('\'', '"')).Where(x => x.Length > 0);
        }

        private static IEnumerable<string> SplitValues(string v)
        {
            return (v ?? "").Split(new[]
            {
                ';',','
            }
            , StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).Where(x => x.Length > 0);
        }

        private static List<Recipient> BuildRecipients(string v)
        {
            return SplitValues(v).Where(IsEmail).Distinct(StringComparer.OrdinalIgnoreCase).Select(x => new Recipient
            {
                EmailAddress = new EmailAddress
                {
                    Address = x
                }
            }
            ).ToList();
        }

        private static string ExtractBetween(string s, string a, string b)
        {
            int i = (s ?? "").IndexOf(a, StringComparison.OrdinalIgnoreCase);
            if (i < 0) return "";
            i += a.Length;
            int j = s.IndexOf(b, i, StringComparison.OrdinalIgnoreCase);
            return j < 0 ? "" : s.Substring(i, j - i).Trim();
        }

        private static string ExtractAfter(string s, string a)
        {
            int i = (s ?? "").IndexOf(a, StringComparison.OrdinalIgnoreCase);
            return i < 0 ? "" : s.Substring(i + a.Length).Trim();
        }

        private static string GetSender(Message m)
        {
            return m?.From?.EmailAddress?.Address ?? m?.Sender?.EmailAddress?.Address ?? "";
        }

        private static bool IsMessageAlreadyProcessed(Message m)
        {
            return m?.SingleValueExtendedProperties?.Any(x => string.Equals(x.Id, processed_property_id, StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(x.Value)) == true;
        }

        private static bool IsEmail(string v)
        {
            return !string.IsNullOrWhiteSpace(v) && Regex.IsMatch(v, @"^[^\s@]+@[^\s@]+\.[^\s@]+$");
        }

        private static bool IsGraphObjectNotFound(Exception exception)
        {
            Exception current = exception;

            while (current != null)
            {
                string message = current.Message ?? "";

                if (message.IndexOf(
                        "specified object was not found",
                        StringComparison.OrdinalIgnoreCase) >= 0 ||
                    message.IndexOf(
                        "not found in the store",
                        StringComparison.OrdinalIgnoreCase) >= 0 ||
                    message.IndexOf(
                        "failed to get the correct properties",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }

                if (current is ApiException apiException &&
                    apiException.ResponseStatusCode == 404)
                {
                    return true;
                }

                current = current.InnerException;
            }

            return false;
        }

        private static bool IsChangeKeyConflict(Exception ex)
        {
            string t = ex?.ToString() ?? "";
            return t.IndexOf("change key", StringComparison.OrdinalIgnoreCase) >= 0 || t.IndexOf("ErrorIrresolvableConflict", StringComparison.OrdinalIgnoreCase) >= 0 || (ex is ApiException a && (a.ResponseStatusCode == 409 || a.ResponseStatusCode == 412));
        }

        private static bool IsTransientGraphError(Exception ex)
        {
            Exception current = ex;
            while (current != null)
            {
                if (current is HttpRequestException ||
                    current is TimeoutException ||
                    current is System.Threading.Tasks.TaskCanceledException)
                    return true;

                if (current is ApiException apiException &&
                    (apiException.ResponseStatusCode == 429 ||
                     apiException.ResponseStatusCode == 503 ||
                     apiException.ResponseStatusCode == 504))
                    return true;

                string text = current.Message ?? "";
                if (text.IndexOf("Too many retries performed", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    text.IndexOf("More than 3 retries encountered", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    text.IndexOf("TooManyRequests", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    text.IndexOf("ApplicationThrottled", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    text.IndexOf("temporarily unavailable", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    text.IndexOf("service unavailable", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    text.IndexOf("12002", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    text.IndexOf("timed out", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    text.IndexOf("An error occurred while sending the request", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;

                current = current.InnerException;
            }
            return false;
        }

        private static string GetInnermostExceptionMessage(Exception ex)
        {
            if (ex == null) return "";
            while (ex.InnerException != null) ex = ex.InnerException;
            return ex.Message ?? "";
        }

        private static void AddImmutableHeader(RequestHeaders h)
        {
            h.Add("Prefer", immutable_id_preference);
        }

        private static string EscapeODataString(string v)
        {
            return (v ?? "").Replace("'", "''");
        }

        private static string Truncate(string v, int l)
        {
            v = v ?? "";
            return v.Length <= l ? v : v.Substring(0, l);
        }

        private static string CleanFileName(string v)
        {
            return string.Join("_", (v ?? "email").Split(Path.GetInvalidFileNameChars())).Trim();
        }

        private static string Left(string v, int l)
        {
            v = v ?? "";
            return v.Substring(0, Math.Min(l, v.Length));
        }

        private static string Right(string v, int l)
        {
            v = v ?? "";
            return v.Substring(Math.Max(0, v.Length - l));
        }

        private static bool IsTrue(string v)
        {
            return string.Equals(v?.Trim(), "TRUE", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsRefreshDue(DateTime d, int m)
        {
            return m <= 0 || DateTime.Now >= d.AddMinutes(m);
        }

        private static string GetServicePath()
        {
            string l = System.Reflection.Assembly.GetEntryAssembly()?.Location;
            return string.IsNullOrWhiteSpace(l) ? AppDomain.CurrentDomain.BaseDirectory : Path.GetDirectoryName(l);
        }

        public void Maquettes_Et_Rapports(string sql_con, string logs, string tmp_folder, string session_name)
        {
            string root = GetServicePath();
            logsFolder = Path.Combine(root, logs ?? "");
            tempFolder = Path.Combine(root, tmp_folder ?? "");
            sessionName = session_name ?? "";
            Directory.CreateDirectory(logsFolder);
            Directory.CreateDirectory(tempFolder);

            JsonFile cfg = JsonConvert.DeserializeObject<JsonFile>(
                GetImcaParameter(sql_con, global_application_name) ?? "");

            if (cfg?.countries == null || cfg.countries.Count == 0)
            {
                throw new InvalidOperationException(
                    global_application_name + " parameters are empty or invalid");
            }

            foreach (Country item in cfg.countries)
            {
                ApplyCountryConfiguration(item, sql_con);
                if (!IsTrue(active))
                    continue;

                ValidateTemplatesReportsConfiguration();
                graphService = ConnectGraph();

                WriteLog(
                    "   Starting templates and reports processing" +
                    " - Resources folder : " + templatesReportsFolder +
                    " - Temp folder : " + tempFolder);

                try
                {
                    WriteLog(
                        "   Creating automatic opening account reminders and refusals");
                    CreateAutomaticOpeningRequests();

                    WriteLog(
                        "   Processing pending opening account templates");
                    ProcessPendingOpeningTemplates();

                    WriteLog(
                        "   Processing pending customer change templates");
                    ProcessPendingCustomerChangeTemplates();

                    WriteLog(
                        "   Checking scheduled opening account reports");
                    ProcessDueOpeningReports();

                    WriteLog(
                        "   Templates and reports processing completed");
                }
                catch (Exception ex)
                {
                    WriteLog(
                        "   Maquettes_Et_Rapports error" +
                        " - Error : " + ex.Message);

                    SendTechnicalAlert(
                        nameof(Maquettes_Et_Rapports),
                        ex.Message,
                        "TEMPLATES AND REPORTS");
                }
                finally
                {
                    graphService = null;
                }
            }
        }

        private void ValidateTemplatesReportsConfiguration()
        {
            if (string.IsNullOrWhiteSpace(sql_creation_compte) ||
                string.IsNullOrWhiteSpace(templatesReportsFolder) ||
                string.IsNullOrWhiteSpace(fr_ouverture_graph_send_as) ||
                string.IsNullOrWhiteSpace(fr_credit_administration_graph_send_as))
                throw new InvalidOperationException("Templates/reports configuration is incomplete");
            if (!Directory.Exists(templatesReportsFolder))
                throw new DirectoryNotFoundException("Templates/reports folder not found : " + templatesReportsFolder);
        }

        private void CreateAutomaticOpeningRequests()
        {
            const string firstReminder = @"
INSERT INTO dbo.envoi_mail(id_dossier,nom_mail,email_destinataire,date_demande,top_traite)
SELECT e.id_dossier,
       CASE WHEN c.id_type_customer IN (2,4) THEN '1ere_relance_expert' ELSE '1ere_relance_export' END,
       e.email_destinataire,GETDATE(),'N'
FROM dbo.T_statut s
JOIN dbo.T_customer c ON c.id=s.idDossier
JOIN dbo.envoi_mail e ON e.id_dossier=c.id
JOIN (SELECT idDossier,MAX(Datereception) Datereception FROM dbo.T_Doc_Recu WHERE fichier_valide=0 GROUP BY idDossier) d ON d.idDossier=e.id_dossier
WHERE e.nom_mail='incomplet_expert' AND e.date_trt IS NOT NULL AND e.top_traite='O'
  AND s.idStatut=1 AND c.id_type_customer IN (2,3,4)
  AND DATEDIFF(day,e.date_trt,GETDATE())>=7
  AND NOT EXISTS(SELECT 1 FROM dbo.envoi_mail x WHERE x.id_dossier=e.id_dossier AND x.nom_mail IN('1ere_relance_expert','1ere_relance_export'));";
            ExecuteNonQuery(sql_creation_compte, firstReminder);

            const string refusal = @"
INSERT INTO dbo.envoi_mail(id_dossier,nom_mail,email_destinataire,date_demande,top_traite)
SELECT e.id_dossier,
       CASE WHEN c.id_type_customer IN (2,4) THEN 'refus_expert' ELSE 'refus_export' END,
       e.email_destinataire,GETDATE(),'N'
FROM dbo.T_statut s
JOIN dbo.T_customer c ON c.id=s.idDossier
JOIN dbo.envoi_mail e ON e.id_dossier=c.id
WHERE e.nom_mail IN('1ere_relance_expert','1ere_relance_export')
  AND e.date_trt IS NOT NULL AND e.top_traite='O' AND s.idStatut=3
  AND c.id_type_customer IN (2,3,4) AND DATEDIFF(day,e.date_trt,GETDATE())>=7
  AND NOT EXISTS(SELECT 1 FROM dbo.envoi_mail x WHERE x.id_dossier=e.id_dossier AND x.nom_mail IN('refus_expert','refus_export') AND x.id>e.id);
UPDATE s SET Comment='3'
FROM dbo.T_Statut s
WHERE EXISTS(SELECT 1 FROM dbo.envoi_mail e WHERE e.id_dossier=s.idDossier AND e.top_traite='N' AND e.nom_mail IN('refus_expert','refus_export'));";
            ExecuteNonQuery(sql_creation_compte, refusal);
        }

        private void ProcessPendingOpeningTemplates()
        {
            const string sql = @"
SELECT e.id,
       e.id_dossier,
       e.nom_mail,
       e.email_destinataire,
       l.pieces_jointes,
       l.[sujet email] AS sujet_email,
       l.nom_fic_html,
       s.idStatut
FROM dbo.envoi_mail AS e
INNER JOIN dbo.T_statut AS s
    ON s.idDossier = e.id_dossier
INNER JOIN dbo.lien_email AS l
    ON l.nom_email_court = e.nom_mail
WHERE e.top_traite = 'N'
  AND e.id > 12719481
ORDER BY e.id;";

            DataTable rows = FillDataTable(sql_creation_compte, sql);
            int sent = 0;
            int skipped = 0;
            int errors = 0;

            WriteLog(
                "       Opening account template queue" +
                " - Minimum envoi_mail ID excluded : 12719481" +
                " - Pending mail(s) found : " + rows.Rows.Count);

            foreach (DataRow row in rows.Rows)
            {
                int id = Convert.ToInt32(row["id"]);
                int dossierId = Convert.ToInt32(row["id_dossier"]);
                int currentStatus = Convert.ToInt32(row["idStatut"]);
                string mailName = Convert.ToString(row["nom_mail"]).Trim();
                string recipient = Convert.ToString(row["email_destinataire"]).Trim();

                try
                {
                    if (!CanSendOpeningTemplate(mailName, currentStatus))
                    {
                        skipped++;
                        WriteLog(
                            "       Opening account template skipped" +
                            " - Queue ID : " + id +
                            " - Dossier : " + dossierId +
                            " - Template : " + mailName +
                            " - Current status : " + currentStatus);
                        continue;
                    }

                    string templatePath =
                        SafeResourcePath(Convert.ToString(row["nom_fic_html"]));
                    string html = File.ReadAllText(templatePath, Encoding.UTF8);
                    TemplateCustomerData data =
                        GetTemplateCustomerData(dossierId, recipient);

                    html = ApplyCommonTemplateValues(html, data, dossierId);

                    if (mailName.StartsWith(
                            "incomplet_",
                            StringComparison.OrdinalIgnoreCase) ||
                        mailName.StartsWith(
                            "1ere_relance_",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        html = ApplyDocumentBlocks(html, dossierId);
                    }

                    if (mailName.StartsWith(
                            "refus_",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        html = ApplyRefusalBlocks(html, dossierId);
                    }

                    ValidateNoKnownUnresolvedMarkers(html, templatePath);

                    string subject = BuildOpeningSubject(
                        mailName,
                        Convert.ToString(row["sujet_email"]),
                        dossierId,
                        data);

                    List<string> attachments =
                        ResolveAttachmentPaths(row["pieces_jointes"]);

                    string bcc = "";
                    if (mailName.Equals(
                            "bienvenue",
                            StringComparison.OrdinalIgnoreCase) ||
                        mailName.Equals(
                            "bienvenue_export",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        bcc = "Composant.fr@ingrammicro.com";
                    }
                    else if (mailName.Equals(
                                 "transformation_en_expert",
                                 StringComparison.OrdinalIgnoreCase))
                    {
                        bcc = "Credit.Analystes@IngramMicro.fr";
                    }

                    string sendAs =
                        mailName.Equals(
                            "RIB",
                            StringComparison.OrdinalIgnoreCase)
                            ? fr_credit_administration_graph_send_as
                            : fr_ouverture_graph_send_as;

                    SendGraphMailWithAttachments(
                        sendAs,
                        recipient,
                        "",
                        bcc,
                        subject,
                        html,
                        attachments);

                    SetMailRequestStatus("dbo.envoi_mail", id, "O");

                    if (mailName.StartsWith(
                            "1ere_relance_",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        UpdateOpeningStatus(dossierId, 3);
                    }

                    if (mailName.StartsWith(
                            "refus_",
                            StringComparison.OrdinalIgnoreCase) &&
                        !mailName.Equals(
                            "refus_express",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        UpdateOpeningStatus(dossierId, 4);
                    }

                    sent++;
                    WriteLog(
                        "       Opening account template sent" +
                        " - Queue ID : " + id +
                        " - Dossier : " + dossierId +
                        " - Template : " + mailName +
                        " - Send as : " + sendAs +
                        " - Recipient : " + recipient +
                        " - Attachment(s) : " + attachments.Count +
                        " - Subject : " + subject);
                }
                catch (Exception ex)
                {
                    errors++;
                    SetMailRequestStatus("dbo.envoi_mail", id, "E");

                    WriteLog(
                        "       Opening account template error" +
                        " - Queue ID : " + id +
                        " - Dossier : " + dossierId +
                        " - Template : " + mailName +
                        " - Recipient : " + recipient +
                        " - Error : " + ex.Message);

                    SendTechnicalAlert(
                        nameof(ProcessPendingOpeningTemplates),
                        "Dossier " + dossierId + " - " + ex.Message,
                        "OPENING TEMPLATE");
                }
            }

            WriteLog(
                "       Opening account template queue summary" +
                " - Found : " + rows.Rows.Count +
                " - Sent : " + sent +
                " - Skipped : " + skipped +
                " - Errors : " + errors);
        }

        private sealed class TemplateCustomerData
        {
            public string LastName = "", FirstName = "", CustomerCode = "", CustomerName = "", Rcs = "", Bank = "", Agency = "", Iban = "", Bic = "";
        }

        private TemplateCustomerData GetTemplateCustomerData(int dossierId, string recipient)
        {
            var d = new TemplateCustomerData();
            const string sql = @"
SELECT TOP(1) ISNULL(i.nom,'') nom,ISNULL(i.prenom,'') prenom,
 ISNULL(c.ImpCustNbr,'') code,ISNULL(c.raison_soc,'') raison,
 RIGHT(REPLACE(ISNULL(c.siret,''),'FR',''),9) rcs,
 ISNULL(b.nom_banque,'') banque,ISNULL(b.agence,'') agence,ISNULL(b.IBAN,'') iban,ISNULL(b.Code_BIC,'') bic
FROM dbo.T_customer c
LEFT JOIN dbo.T_interlocuteurs i ON i.id=c.id AND i.email=@EMAIL
LEFT JOIN dbo.T_IBAN b ON b.iddossier=c.id
WHERE c.id=@ID;";
            DataTable t = FillDataTable(sql_creation_compte, sql, new SqlParameter("@EMAIL", SqlDbType.VarChar, 320) { Value = recipient ?? "" }, new SqlParameter("@ID", SqlDbType.Int) { Value = dossierId });
            if (t.Rows.Count == 0) return d;
            DataRow r = t.Rows[0];
            d.LastName = Convert.ToString(r["nom"]); d.FirstName = Convert.ToString(r["prenom"]); d.CustomerCode = Convert.ToString(r["code"]); d.CustomerName = Convert.ToString(r["raison"]); d.Rcs = Convert.ToString(r["rcs"]); d.Bank = Convert.ToString(r["banque"]); d.Agency = Convert.ToString(r["agence"]); d.Iban = Convert.ToString(r["iban"]); d.Bic = Convert.ToString(r["bic"]);
            return d;
        }

        private static string ApplyCommonTemplateValues(string html, TemplateCustomerData d, int dossierId)
        {
            return (html ?? "").Replace("[NOM]", d.LastName).Replace("[PRENOM]", d.FirstName).Replace("[CODE_CLIENT]", d.CustomerCode)
                .Replace("[NOM_CLIENT]", d.CustomerName).Replace("[RCS]", d.Rcs).Replace("[BANQUE]", d.Bank).Replace("[AGENCE]", d.Agency)
                .Replace("[IBAN]", d.Iban).Replace("[BIC]", d.Bic).Replace("[NUM_DEMANDE]", dossierId.ToString(CultureInfo.InvariantCulture));
        }

        private string ApplyDocumentBlocks(string html, int dossierId)
        {
            DataTable valid = FillDataTable(sql_creation_compte, "SELECT d.NomDoc FROM dbo.T_Doc_Recu r JOIN dbo.T_Doc d ON d.idDoc=r.idDoc WHERE r.fichier_valide=1 AND r.idDossier=@ID;", new SqlParameter("@ID", SqlDbType.Int) { Value = dossierId });
            foreach (DataRow r in valid.Rows) html = RemoveDelimitedBlock(html, "[" + Convert.ToString(r[0]) + "]", "[/" + Convert.ToString(r[0]) + "]");
            DataTable all = FillDataTable(sql_creation_compte, "SELECT NomDoc FROM dbo.T_Doc;");
            foreach (DataRow r in all.Rows) { string n = Convert.ToString(r[0]); html = html.Replace("[" + n + "]", "").Replace("[/" + n + "]", ""); }
            return html;
        }

        private string ApplyRefusalBlocks(string html, int dossierId)
        {
            DataTable all = FillDataTable(sql_creation_compte, "SELECT id_motif_refus FROM dbo.T_customer_libelle_refus ORDER BY id_motif_refus;");
            DataTable selected = FillDataTable(sql_creation_compte, "SELECT TRY_CONVERT(int,Comment) id FROM dbo.T_Statut WHERE idDossier=@ID AND TRY_CONVERT(int,Comment) IS NOT NULL;", new SqlParameter("@ID", SqlDbType.Int) { Value = dossierId });
            var keep = new HashSet<int>(selected.AsEnumerable().Select(r => Convert.ToInt32(r[0])));
            foreach (DataRow r in all.Rows) { int n = Convert.ToInt32(r[0]); string a = "[" + n + "]", b = "[/" + n + "]"; html = keep.Contains(n) ? html.Replace(a, "").Replace(b, "") : RemoveDelimitedBlock(html, a, b); }
            return html;
        }

        private static string RemoveDelimitedBlock(string value, string start, string end)
        {
            int a = (value ?? "").IndexOf(start, StringComparison.OrdinalIgnoreCase); if (a < 0) return value;
            int b = value.IndexOf(end, a + start.Length, StringComparison.OrdinalIgnoreCase); if (b < 0) return value;
            return value.Remove(a, b + end.Length - a);
        }

        private static bool CanSendOpeningTemplate(string name, int status)
        {
            if (new[] { "bienvenue", "bienvenue_export", "rgpd", "transformation_en_expert", "RIB" }.Contains(name, StringComparer.OrdinalIgnoreCase)) return status == 6;
            return name.StartsWith("incomplet_", StringComparison.OrdinalIgnoreCase) || name.StartsWith("refus_", StringComparison.OrdinalIgnoreCase) || name.StartsWith("1ere_relance_", StringComparison.OrdinalIgnoreCase) || name.IndexOf("ligne_credit", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string BuildOpeningSubject(string name, string subject, int dossierId, TemplateCustomerData data)
        {
            if (name.Equals("RIB", StringComparison.OrdinalIgnoreCase)) return data.CustomerCode + " - " + data.CustomerName + " - RCS " + data.Rcs + " - Dossier n° " + dossierId;
            return (subject ?? "").Replace(" - version Export", "").Replace(" - Expert", "").Replace(" - Export", "").Replace(" - Express", "") + " - Dossier n° " + dossierId;
        }

        private void ProcessPendingCustomerChangeTemplates()
        {
            const string sql = @"
SELECT q.id,
       q.id_demande,
       q.nom_mail,
       w.email_demandeur,
       t.libelle_typologie,
       ISNULL(w.motif_refus, '') AS motif_refus,
       CONVERT(varchar(10), w.quand, 103) AS date_demande,
       w.branch_customer_nbr,
       CASE
           WHEN w.id_typologie = 1 THEN ISNULL(a.cust_name, '')
           ELSE ISNULL(c.cust_name, '')
       END AS cust_name
FROM dbo.T_Changement_coordonnees_client_via_le_WEB_envoi_emails AS q
INNER JOIN dbo.T_Changement_coordonnees_client_via_le_WEB AS w
    ON w.id_demande = q.id_demande
INNER JOIN dbo.T_Changement_coordonnees_client_via_le_WEB_liste_typologie AS t
    ON t.id_typologie = w.id_typologie
LEFT JOIN DSS_COPIE.dbo.customer AS c
    ON c.branch_customer_nbr = w.branch_customer_nbr
LEFT JOIN dbo.T_Changement_coordonnees_client_via_le_WEB_adresse_facturation AS a
    ON a.id_demande = w.id_demande
WHERE q.top_traite = 'N'
ORDER BY q.id;";

            DataTable rows = FillDataTable(sql_creation_compte, sql);
            int sent = 0;
            int errors = 0;

            WriteLog(
                "       Customer change template queue" +
                " - Pending mail(s) found : " + rows.Rows.Count);

            foreach (DataRow row in rows.Rows)
            {
                int id = Convert.ToInt32(row["id"]);
                int requestId = Convert.ToInt32(row["id_demande"]);
                string templateName = Convert.ToString(row["nom_mail"]);
                string recipient = Convert.ToString(row["email_demandeur"]);

                try
                {
                    string path = SafeResourcePath(templateName + ".html");
                    string html = File.ReadAllText(path, Encoding.UTF8);

                    html = html
                        .Replace("[CODE_CLIENT]", Convert.ToString(row["branch_customer_nbr"]))
                        .Replace("[MOTIF]", Convert.ToString(row["libelle_typologie"]))
                        .Replace("[NOM_CLIENT]", Convert.ToString(row["cust_name"]))
                        .Replace("[DATE_DEMANDE]", Convert.ToString(row["date_demande"]))
                        .Replace("[MOTIF_REFUS]", Convert.ToString(row["motif_refus"]))
                        .Replace(
                            "[NUM_DEMANDE]",
                            requestId.ToString(CultureInfo.InvariantCulture));

                    ValidateNoKnownUnresolvedMarkers(html, path);

                    bool refusal = templateName.EndsWith(
                        "_REFUS",
                        StringComparison.OrdinalIgnoreCase);

                    string subject =
                        (refusal ? "Refus" : "Acceptation") +
                        " de votre demande n° " + requestId +
                        " (" + Convert.ToString(row["libelle_typologie"]) + ")";

                    string cc = templateName.Contains("LIVRAISON")
                        ? ""
                        : admin_ventes;
                    string bcc = templateName.Contains("RIB")
                        ? "analystes.credit@ingrammicro.fr"
                        : "rapports_recouvrement@ingrammicro.com";

                    SendGraphMailWithAttachments(
                        fr_credit_administration_graph_send_as,
                        recipient,
                        cc,
                        bcc,
                        subject,
                        html,
                        new List<string>());

                    SetMailRequestStatus(
                        "dbo.T_Changement_coordonnees_client_via_le_WEB_envoi_emails",
                        id,
                        "O");

                    sent++;
                    WriteLog(
                        "       Customer change template sent" +
                        " - Queue ID : " + id +
                        " - Request : " + requestId +
                        " - Template : " + templateName +
                        " - Send as : " +
                        fr_credit_administration_graph_send_as +
                        " - Recipient : " + recipient +
                        " - Subject : " + subject);
                }
                catch (Exception ex)
                {
                    errors++;
                    SetMailRequestStatus(
                        "dbo.T_Changement_coordonnees_client_via_le_WEB_envoi_emails",
                        id,
                        "E");

                    WriteLog(
                        "       Customer change template error" +
                        " - Queue ID : " + id +
                        " - Request : " + requestId +
                        " - Template : " + templateName +
                        " - Recipient : " + recipient +
                        " - Error : " + ex.Message);

                    SendTechnicalAlert(
                        nameof(ProcessPendingCustomerChangeTemplates),
                        "Demande " + requestId + " - " + ex.Message,
                        "CUSTOMER CHANGE TEMPLATE");
                }
            }

            WriteLog(
                "       Customer change template queue summary" +
                " - Found : " + rows.Rows.Count +
                " - Sent : " + sent +
                " - Errors : " + errors);
        }

        private void ProcessDueOpeningReports()
        {
            DateTime now = DateTime.Now;
            DateTime nextCounter =
                GetReportParameterDate("date_envoi_prochain_rapport_compteur");

            WriteLog(
                "       Opening account report schedule" +
                " - Current date : " + now.ToString("dd/MM/yyyy HH:mm:ss") +
                " - Counter report due : " +
                nextCounter.ToString("dd/MM/yyyy HH:mm:ss"));

            if (now >= nextCounter)
            {
                DateTime last = GetReportParameterDate(
                    "date_envoi_dernier_rapport_compteur");

                WriteLog(
                    "       Generating detailed opening account report" +
                    " - Previous report date : " +
                    last.ToString("dd/MM/yyyy HH:mm:ss"));

                string file = CreateDetailedOpeningReport(last);

                SendGraphMailWithAttachments(
                    fr_ouverture_graph_send_as,
                    email_rapport_compteur,
                    "",
                    "",
                    "Détail des ouvertures de compte (Semaine précédente)",
                    "Ci-joint.",
                    new List<string> { file });

                SetReportParameterDate(
                    "date_envoi_dernier_rapport_compteur",
                    now);
                SetReportParameterDate(
                    "date_envoi_prochain_rapport_compteur",
                    nextCounter.AddDays(7));

                WriteLog(
                    "       Detailed opening account report sent" +
                    " - Send as : " + fr_ouverture_graph_send_as +
                    " - File : " + file +
                    " - Recipients : " + email_rapport_compteur);
            }
            else
            {
                WriteLog(
                    "       Detailed opening account report not due");
            }

            DateTime nextStats = GetReportParameterDate(
                "date_envoi_prochain_rapport_stat_ouverture");

            WriteLog(
                "       Opening account statistics report schedule" +
                " - Due : " + nextStats.ToString("dd/MM/yyyy HH:mm:ss"));

            if (now >= nextStats)
            {
                WriteLog(
                    "       Generating opening account statistics report");

                string file = CreateOpeningStatisticsReport();

                SendGraphMailWithAttachments(
                    fr_ouverture_graph_send_as,
                    email_rapport_stat_ouverture,
                    "",
                    "",
                    "Statistiques ouvertures de compte",
                    "Ci-joint.",
                    new List<string> { file });

                SetReportParameterDate(
                    "date_envoi_prochain_rapport_stat_ouverture",
                    nextStats.AddDays(7));

                WriteLog(
                    "       Opening account statistics report sent" +
                    " - Send as : " + fr_ouverture_graph_send_as +
                    " - File : " + file +
                    " - Recipients : " + email_rapport_stat_ouverture);
            }
            else
            {
                WriteLog(
                    "       Opening account statistics report not due");
            }
        }

        private string CreateOpeningStatisticsReport()
        {
            int reportYear = DateTime.Now.Year;
            DateTime periodStart = new DateTime(reportYear, 1, 1);
            DateTime periodEnd = periodStart.AddYears(1);

            string output = Path.Combine(
                tempFolder,
                DateTime.Now.ToString("yyyyMMdd_HHmmss") +
                "_stat_ouverture_compte_" +
                reportYear +
                ".xlsx");

            DataTable requests = FillDataTable(
                sql_creation_compte,
                @"SELECT CONVERT(char(6), Date_integration_intranet, 112) AS Mois,
                         COUNT(*) AS nb_dde
                  FROM dbo.T_customer
                  WHERE Date_integration_intranet >= @START
                    AND Date_integration_intranet < @END
                  GROUP BY CONVERT(char(6), Date_integration_intranet, 112)
                  ORDER BY Mois;",
                new SqlParameter("@START", SqlDbType.DateTime2) { Value = periodStart },
                new SqlParameter("@END", SqlDbType.DateTime2) { Value = periodEnd });

            DataTable creationsByType = FillDataTable(
                sql_creation_compte,
                @"SELECT CONVERT(char(6), s.Date, 112) AS Mois,
                         ISNULL(NULLIF(LTRIM(RTRIM(t.libel_type_customer)), ''), 'Non renseigné') AS TypeClient,
                         COUNT(*) AS nb_creation
                  FROM dbo.T_Statut AS s
                  INNER JOIN dbo.T_customer AS c
                      ON c.id = s.idDossier
                  LEFT JOIN dbo.T_type_customer AS t
                      ON t.id_type_customer = c.id_type_customer
                  WHERE s.idStatut = 6
                    AND s.Date >= @START
                    AND s.Date < @END
                  GROUP BY CONVERT(char(6), s.Date, 112),
                           ISNULL(NULLIF(LTRIM(RTRIM(t.libel_type_customer)), ''), 'Non renseigné')
                  ORDER BY Mois, TypeClient;",
                new SqlParameter("@START", SqlDbType.DateTime2) { Value = periodStart },
                new SqlParameter("@END", SqlDbType.DateTime2) { Value = periodEnd });

            DataTable creations = FillDataTable(
                sql_creation_compte,
                @"SELECT CONVERT(char(6), s.Date, 112) AS Mois,
                         COUNT(*) AS nb_creation
                  FROM dbo.T_Statut AS s
                  WHERE s.idStatut = 6
                    AND s.Date >= @START
                    AND s.Date < @END
                  GROUP BY CONVERT(char(6), s.Date, 112)
                  ORDER BY Mois;",
                new SqlParameter("@START", SqlDbType.DateTime2) { Value = periodStart },
                new SqlParameter("@END", SqlDbType.DateTime2) { Value = periodEnd });

            DataTable refusals = FillDataTable(
                sql_creation_compte,
                @"SELECT CONVERT(char(6), s.Date, 112) AS Mois,
                         COUNT(*) AS nb_refus
                  FROM dbo.T_Statut AS s
                  WHERE s.idStatut = 4
                    AND s.Date >= @START
                    AND s.Date < @END
                  GROUP BY CONVERT(char(6), s.Date, 112)
                  ORDER BY Mois;",
                new SqlParameter("@START", SqlDbType.DateTime2) { Value = periodStart },
                new SqlParameter("@END", SqlDbType.DateTime2) { Value = periodEnd });

            DataTable refusalDetails = FillDataTable(
                sql_creation_compte,
                @"SELECT s.idDossier,
                         s.Date,
                         CASE
                             WHEN r.libelle_refus IS NULL THEN s.Comment
                             ELSE r.libelle_refus
                         END AS motif_refus,
                         CONVERT(char(6), s.Date, 112) AS Mois,
                         YEAR(s.Date) AS Annee
                  FROM dbo.T_Statut AS s
                  LEFT JOIN dbo.T_customer_libelle_refus AS r
                      ON s.Comment = CONVERT(varchar(20), r.id_motif_refus)
                  WHERE s.idStatut = 4
                    AND s.Date >= @START
                    AND s.Date < @END
                  ORDER BY s.Date DESC, s.idDossier DESC;",
                new SqlParameter("@START", SqlDbType.DateTime2) { Value = periodStart },
                new SqlParameter("@END", SqlDbType.DateTime2) { Value = periodEnd });

            DataTable refusedThenCreated = FillDataTable(
                sql_creation_compte,
                @"SELECT CONVERT(char(6), ref.Date, 112) AS Mois,
                         COUNT(DISTINCT cre.idDossier) AS nb_refuse_puis_cree
                  FROM
                  (
                      SELECT idDossier, MIN(Date) AS Date
                      FROM dbo.T_historique
                      WHERE idStatut = 4
                        AND Date >= @START
                        AND Date < @END
                      GROUP BY idDossier
                  ) AS ref
                  INNER JOIN
                  (
                      SELECT idDossier, MIN(Date) AS Date
                      FROM dbo.T_historique
                      WHERE idStatut = 6
                        AND Date >= @START
                        AND Date < @END
                      GROUP BY idDossier
                  ) AS cre
                      ON cre.idDossier = ref.idDossier
                     AND ref.Date < cre.Date
                  GROUP BY CONVERT(char(6), ref.Date, 112)
                  ORDER BY Mois;",
                new SqlParameter("@START", SqlDbType.DateTime2) { Value = periodStart },
                new SqlParameter("@END", SqlDbType.DateTime2) { Value = periodEnd });

            using (var workbook = new XLWorkbook())
            {
                IXLWorksheet creationSheet = workbook.AddWorksheet("creation");
                WriteCrossTab(
                    creationSheet,
                    "Ouvertures de compte " + reportYear,
                    creationsByType,
                    "TypeClient",
                    "Mois",
                    "nb_creation");

                IXLWorksheet refusalReasonSheet = workbook.AddWorksheet("refus 1");
                WriteRefusalReasonCrossTab(
                    refusalReasonSheet,
                    "Refus par motif " + reportYear,
                    refusalDetails);

                IXLWorksheet refusalSummarySheet = workbook.AddWorksheet("refus 2");
                WriteMonthlySummary(
                    refusalSummarySheet,
                    "Synthèse mensuelle " + reportYear,
                    requests,
                    creations,
                    refusals,
                    refusedThenCreated);

                IXLWorksheet dataSheet = workbook.AddWorksheet("datas");
                int nextRow = 1;
                nextRow = WriteDataSection(dataSheet, nextRow, "Demandes", requests);
                nextRow = WriteDataSection(dataSheet, nextRow, "Créations par type", creationsByType);
                nextRow = WriteDataSection(dataSheet, nextRow, "Créations", creations);
                nextRow = WriteDataSection(dataSheet, nextRow, "Refus", refusals);
                nextRow = WriteDataSection(dataSheet, nextRow, "Détail des refus", refusalDetails);
                WriteDataSection(
                    dataSheet,
                    nextRow,
                    "Dossiers refusés puis créés",
                    refusedThenCreated);

                foreach (IXLWorksheet worksheet in workbook.Worksheets)
                {
                    worksheet.SheetView.FreezeRows(2);
                    worksheet.Columns().AdjustToContents(1, 60);
                    worksheet.Rows().AdjustToContents();
                }

                workbook.SaveAs(output);
            }

            WriteLog(
                "       Opening account statistics workbook generated" +
                " - Year : " + reportYear +
                " - Requests : " + SumColumn(requests, "nb_dde") +
                " - Creations : " + SumColumn(creations, "nb_creation") +
                " - Refusals : " + SumColumn(refusals, "nb_refus") +
                " - File : " + output);

            return output;
        }

        private static int WriteDataSection(
            IXLWorksheet worksheet,
            int startRow,
            string title,
            DataTable data)
        {
            worksheet.Cell(startRow, 1).Value = title;
            worksheet.Cell(startRow, 1).Style.Font.Bold = true;
            worksheet.Cell(startRow, 1).Style.Fill.BackgroundColor = XLColor.DarkBlue;
            worksheet.Cell(startRow, 1).Style.Font.FontColor = XLColor.White;

            int tableRow = startRow + 1;
            if (data.Rows.Count == 0)
            {
                worksheet.Cell(tableRow, 1).Value = "Aucune donnée";
                return tableRow + 2;
            }

            worksheet.Cell(tableRow, 1).InsertTable(data, false);
            return tableRow + data.Rows.Count + 3;
        }

        private static void WriteCrossTab(
            IXLWorksheet worksheet,
            string title,
            DataTable source,
            string rowField,
            string columnField,
            string valueField)
        {
            worksheet.Cell("A1").Value = title;
            worksheet.Cell("A1").Style.Font.Bold = true;
            worksheet.Cell("A1").Style.Font.FontSize = 14;

            List<string> months = source.AsEnumerable()
                .Select(row => Convert.ToString(row[columnField]))
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(value => value)
                .ToList();

            List<string> categories = source.AsEnumerable()
                .Select(row => Convert.ToString(row[rowField]))
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(value => value)
                .ToList();

            worksheet.Cell(2, 1).Value = rowField;
            for (int index = 0; index < months.Count; index++)
                worksheet.Cell(2, index + 2).Value = months[index];
            worksheet.Cell(2, months.Count + 2).Value = "Total";

            int rowNumber = 3;
            foreach (string category in categories)
            {
                worksheet.Cell(rowNumber, 1).Value = category;
                int categoryTotal = 0;
                for (int monthIndex = 0; monthIndex < months.Count; monthIndex++)
                {
                    int value = source.AsEnumerable()
                        .Where(row => string.Equals(
                            Convert.ToString(row[rowField]),
                            category,
                            StringComparison.OrdinalIgnoreCase))
                        .Where(row => string.Equals(
                            Convert.ToString(row[columnField]),
                            months[monthIndex],
                            StringComparison.OrdinalIgnoreCase))
                        .Sum(row => Convert.ToInt32(row[valueField]));
                    worksheet.Cell(rowNumber, monthIndex + 2).Value = value;
                    categoryTotal += value;
                }
                worksheet.Cell(rowNumber, months.Count + 2).Value = categoryTotal;
                rowNumber++;
            }

            worksheet.Cell(rowNumber, 1).Value = "Total général";
            for (int monthIndex = 0; monthIndex < months.Count; monthIndex++)
            {
                int monthTotal = source.AsEnumerable()
                    .Where(row => string.Equals(
                        Convert.ToString(row[columnField]),
                        months[monthIndex],
                        StringComparison.OrdinalIgnoreCase))
                    .Sum(row => Convert.ToInt32(row[valueField]));
                worksheet.Cell(rowNumber, monthIndex + 2).Value = monthTotal;
            }
            worksheet.Cell(rowNumber, months.Count + 2).Value =
                source.AsEnumerable().Sum(row => Convert.ToInt32(row[valueField]));

            StyleSummaryRange(worksheet, rowNumber, months.Count + 2);
        }

        private static void WriteRefusalReasonCrossTab(
            IXLWorksheet worksheet,
            string title,
            DataTable details)
        {
            DataTable summarized = new DataTable();
            summarized.Columns.Add("Motif", typeof(string));
            summarized.Columns.Add("Mois", typeof(string));
            summarized.Columns.Add("Nombre", typeof(int));

            var groups = details.AsEnumerable()
                .GroupBy(row => new
                {
                    Motif = string.IsNullOrWhiteSpace(Convert.ToString(row["motif_refus"]))
                        ? "Non renseigné"
                        : Convert.ToString(row["motif_refus"]),
                    Mois = Convert.ToString(row["Mois"])
                });

            foreach (var group in groups)
                summarized.Rows.Add(group.Key.Motif, group.Key.Mois, group.Count());

            WriteCrossTab(
                worksheet,
                title,
                summarized,
                "Motif",
                "Mois",
                "Nombre");
        }

        private static void WriteMonthlySummary(
            IXLWorksheet worksheet,
            string title,
            DataTable requests,
            DataTable creations,
            DataTable refusals,
            DataTable refusedThenCreated)
        {
            worksheet.Cell("A1").Value = title;
            worksheet.Cell("A1").Style.Font.Bold = true;
            worksheet.Cell("A1").Style.Font.FontSize = 14;

            string[] headers =
            {
                "Mois",
                "Demandes",
                "Créations",
                "Refus",
                "Refusés puis créés"
            };
            for (int index = 0; index < headers.Length; index++)
                worksheet.Cell(2, index + 1).Value = headers[index];

            int year = DateTime.Now.Year;
            for (int month = 1; month <= 12; month++)
            {
                string monthKey = year.ToString(CultureInfo.InvariantCulture) +
                                  month.ToString("00", CultureInfo.InvariantCulture);
                int row = month + 2;
                worksheet.Cell(row, 1).Value = monthKey;
                worksheet.Cell(row, 2).Value = FindMonthlyValue(requests, monthKey, "nb_dde");
                worksheet.Cell(row, 3).Value = FindMonthlyValue(creations, monthKey, "nb_creation");
                worksheet.Cell(row, 4).Value = FindMonthlyValue(refusals, monthKey, "nb_refus");
                worksheet.Cell(row, 5).Value = FindMonthlyValue(
                    refusedThenCreated,
                    monthKey,
                    "nb_refuse_puis_cree");
            }

            worksheet.Cell(15, 1).Value = "Total général";
            worksheet.Cell(15, 2).Value = SumColumn(requests, "nb_dde");
            worksheet.Cell(15, 3).Value = SumColumn(creations, "nb_creation");
            worksheet.Cell(15, 4).Value = SumColumn(refusals, "nb_refus");
            worksheet.Cell(15, 5).Value = SumColumn(
                refusedThenCreated,
                "nb_refuse_puis_cree");

            StyleSummaryRange(worksheet, 15, 5);
        }

        private static int FindMonthlyValue(
            DataTable table,
            string month,
            string valueColumn)
        {
            DataRow row = table.AsEnumerable().FirstOrDefault(item =>
                string.Equals(
                    Convert.ToString(item["Mois"]),
                    month,
                    StringComparison.OrdinalIgnoreCase));
            return row == null ? 0 : Convert.ToInt32(row[valueColumn]);
        }

        private static int SumColumn(DataTable table, string columnName)
        {
            return table.AsEnumerable().Sum(row =>
                row[columnName] == DBNull.Value
                    ? 0
                    : Convert.ToInt32(row[columnName]));
        }

        private static void StyleSummaryRange(
            IXLWorksheet worksheet,
            int lastRow,
            int lastColumn)
        {
            IXLRange header = worksheet.Range(2, 1, 2, lastColumn);
            header.Style.Font.Bold = true;
            header.Style.Fill.BackgroundColor = XLColor.DarkBlue;
            header.Style.Font.FontColor = XLColor.White;

            IXLRange total = worksheet.Range(lastRow, 1, lastRow, lastColumn);
            total.Style.Font.Bold = true;
            total.Style.Fill.BackgroundColor = XLColor.LightBlue;

            worksheet.Range(2, 1, lastRow, lastColumn)
                .Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        }

        private void WriteQuery(IXLWorksheet ws, string start, string clearColumns, string sql)
        { ws.Columns(clearColumns).Clear(XLClearOptions.Contents); DataTable t = FillDataTable(sql_creation_compte, sql); ws.Cell(start).InsertTable(t, false); }

        private string CreateDetailedOpeningReport(DateTime lastReport)
        {
            string source = SafeResourcePath("modele_calendrier_jour_ferie.xlsx"); string output = Path.Combine(tempFolder, DateTime.Now.ToString("yyyyMMdd_HHmmss") + "_detail_ouverture_compte.xlsx"); File.Copy(source, output, true);
            DataTable data = FillDataTable(sql_creation_compte, BuildDetailedReportSql(), new SqlParameter("@LAST_REPORT", SqlDbType.DateTime2) { Value = lastReport });
            InsertDetailedStats(data, lastReport);
            using (var wb = new XLWorkbook(output))
            {
                IXLWorksheet result = wb.Worksheets.FirstOrDefault(x => x.Name.Equals("resultat", StringComparison.OrdinalIgnoreCase)) ?? wb.AddWorksheet("resultat"); result.Clear(); result.Cell(1, 1).InsertTable(data, false);
                IXLWorksheet settings = wb.Worksheet("ne pas modifier"); settings.Cell("E2").Value = DateTime.Now.Year; TimeSpan start = ReadTime(settings, "B2", new TimeSpan(9, 0, 0)), end = ReadTime(settings, "B3", new TimeSpan(18, 0, 0)); HashSet<DateTime> holidays = FrenchHolidays(data.AsEnumerable().SelectMany(r => new[] { AsNullableDate(r, "Date_integration_intranet"), AsNullableDate(r, "Date") }).Where(d => d.HasValue).Select(d => d.Value.Year));
                for (int i = 0; i < data.Rows.Count; i++) { DataRow r = data.Rows[i]; DateTime? begin = AsNullableDate(r, "Date_integration_intranet"), finish = AsNullableDate(r, "Date"), phoneBegin = AsNullableDate(r, "Date de passage par statut 13 A relancer par téléphone"), phoneEnd = AsNullableDate(r, "Date de passage par statut 14 A créer"); double total = BusinessSeconds(begin, finish, start, end, holidays), phone = BusinessSeconds(phoneBegin, phoneEnd, start, end, holidays), net = Math.Max(0, total - phone); result.Cell(i + 2, 15).Value = TimeSpan.FromSeconds(net); result.Cell(i + 2, 16).Value = TimeSpan.FromSeconds(total); UpdateStatsBusinessDelay(Convert.ToInt32(r["idDossier"]), lastReport, Convert.ToInt32(Math.Round(net))); }
                result.Column(15).Style.NumberFormat.Format = "[h]:mm:ss"; result.Column(16).Style.NumberFormat.Format = "[h]:mm:ss"; wb.Save();
            }
            return output;
        }

        private static string BuildDetailedReportSql() { return @"SELECT s.idDossier,s.idStatut,l.LibelleStatut,s.Date,c.id_type_customer,t.libel_type_customer,c.Date_integration_intranet,DATEDIFF(hour,c.Date_integration_intranet,s.Date) delai_traitement_en_heure,CAST(NULL AS datetime) [Date passage par statut incomplet],CAST(NULL AS int) [delai entre l'enregistrement et passage en incomplet],CAST(NULL AS int) [delai entre passage en incomplet et dernier statut connu],h13.d [Date de passage par statut 13 A relancer par téléphone],h14.d [Date de passage par statut 14 A créer],DATEDIFF(hour,h13.d,h14.d) [delai de passage entre le statut  13 a relancer par téléphone et le statut 14 à créer],CAST(NULL AS int) [delai traitement en heure déduction faite du temps de relance tel] FROM dbo.T_Statut s JOIN dbo.T_customer c ON c.id=s.idDossier JOIN dbo.T_type_customer t ON t.id_type_customer=c.id_type_customer JOIN dbo.T_Libelle_Statut l ON l.IdStatut=s.idStatut OUTER APPLY(SELECT MIN(Date) d FROM dbo.T_historique WHERE idDossier=s.idDossier AND idStatut=13) h13 OUTER APPLY(SELECT MIN(Date) d FROM dbo.T_historique WHERE idDossier=s.idDossier AND idStatut=14) h14 WHERE s.Date>@LAST_REPORT ORDER BY s.idDossier;"; }
        private static DateTime? AsNullableDate(DataRow r, string c) { return !r.Table.Columns.Contains(c) || r[c] == DBNull.Value ? (DateTime?)null : Convert.ToDateTime(r[c]); }
        private static TimeSpan ReadTime(IXLWorksheet ws, string cell, TimeSpan fallback) { try { object v = ws.Cell(cell).Value; if (v is TimeSpan) return (TimeSpan)v; if (v is DateTime) return ((DateTime)v).TimeOfDay; if (TimeSpan.TryParse(Convert.ToString(v), out TimeSpan parsed)) return parsed; } catch { } return fallback; }
        private static double BusinessSeconds(DateTime? from, DateTime? to, TimeSpan start, TimeSpan end, HashSet<DateTime> holidays) { if (!from.HasValue || !to.HasValue || to <= from) return 0; double seconds = 0; for (DateTime d = from.Value.Date; d <= to.Value.Date; d = d.AddDays(1)) { if (d.DayOfWeek == DayOfWeek.Saturday || d.DayOfWeek == DayOfWeek.Sunday || holidays.Contains(d)) continue; DateTime a = d + start, b = d + end, left = from.Value > a ? from.Value : a, right = to.Value < b ? to.Value : b; if (right > left) seconds += (right - left).TotalSeconds; } return seconds; }
        private static HashSet<DateTime> FrenchHolidays(IEnumerable<int> years) { var h = new HashSet<DateTime>(); foreach (int y in years.Distinct()) { DateTime easter = EasterSunday(y); foreach (DateTime d in new[] { new DateTime(y, 1, 1), easter.AddDays(1), new DateTime(y, 5, 1), new DateTime(y, 5, 8), easter.AddDays(39), easter.AddDays(50), new DateTime(y, 7, 14), new DateTime(y, 8, 15), new DateTime(y, 11, 1), new DateTime(y, 11, 11), new DateTime(y, 12, 25) }) h.Add(d.Date); } return h; }
        private static DateTime EasterSunday(int year) { int a = year % 19, b = year / 100, c = year % 100, d = b / 4, e = b % 4, f = (b + 8) / 25, g = (b - f + 1) / 3, h = (19 * a + b - d - g + 15) % 30, i = c / 4, k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7, m = (a + 11 * h + 22 * l) / 451, month = (h + l - 7 * m + 114) / 31, day = (h + l - 7 * m + 114) % 31 + 1; return new DateTime(year, month, day); }
        private void InsertDetailedStats(DataTable data, DateTime lastReport) { foreach (DataRow r in data.Rows) { ExecuteNonQuery(sql_creation_compte, @"IF NOT EXISTS(SELECT 1 FROM dbo.T_STATS_OUVERTURE_COMPTE WHERE idDossier=@ID AND annee=DATEPART(year,@D) AND num_sem=DATEPART(week,@D)) INSERT INTO dbo.T_STATS_OUVERTURE_COMPTE(idDossier,idStatut,LibelleStatut,[Date],id_type_customer,libel_type_customer,Date_integration_intranet,delai_traitement_en_heure,annee,num_sem) VALUES(@ID,@S,@L,@DATE,@T,@TL,@I,@H,DATEPART(year,@D),DATEPART(week,@D));", new SqlParameter("@ID", SqlDbType.Int) { Value = r["idDossier"] }, new SqlParameter("@S", SqlDbType.Int) { Value = r["idStatut"] }, new SqlParameter("@L", SqlDbType.NVarChar, 500) { Value = Convert.ToString(r["LibelleStatut"]) }, new SqlParameter("@DATE", SqlDbType.DateTime) { Value = r["Date"] }, new SqlParameter("@T", SqlDbType.Int) { Value = r["id_type_customer"] }, new SqlParameter("@TL", SqlDbType.NVarChar, 500) { Value = Convert.ToString(r["libel_type_customer"]) }, new SqlParameter("@I", SqlDbType.DateTime) { Value = r["Date_integration_intranet"] }, new SqlParameter("@H", SqlDbType.Int) { Value = r["delai_traitement_en_heure"] == DBNull.Value ? (object)DBNull.Value : r["delai_traitement_en_heure"] }, new SqlParameter("@D", SqlDbType.DateTime) { Value = lastReport }); } }
        private void UpdateStatsBusinessDelay(int dossierId, DateTime period, int seconds) { ExecuteNonQuery(sql_creation_compte, "UPDATE dbo.T_STATS_OUVERTURE_COMPTE SET [delai traitement en heure déduction faite du temps de relance tel]=@S WHERE idDossier=@ID AND annee=DATEPART(year,@D) AND num_sem=DATEPART(week,@D);", new SqlParameter("@S", SqlDbType.Int) { Value = seconds }, new SqlParameter("@ID", SqlDbType.Int) { Value = dossierId }, new SqlParameter("@D", SqlDbType.DateTime) { Value = period }); }

        private DateTime GetReportParameterDate(string name) { string v = ExecuteScalarString(sql_creation_compte, "SELECT ISNULL(ParameterValue,'') FROM dbo.T_STATS_OUVERTURE_COMPTEUR_PARAMETERS WHERE ParameterName=@N;", new SqlParameter("@N", SqlDbType.VarChar, 100) { Value = name }); if (!DateTime.TryParse(v, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out DateTime d)) throw new InvalidDataException("Invalid report parameter " + name + " : " + v); return d; }
        private void SetReportParameterDate(string name, DateTime value) { ExecuteNonQuery(sql_creation_compte, @"MERGE dbo.T_STATS_OUVERTURE_COMPTEUR_PARAMETERS t USING(SELECT @N ParameterName)s ON s.ParameterName=t.ParameterName WHEN MATCHED THEN UPDATE SET ParameterValue=@V,UpdatedAtUtc=SYSUTCDATETIME() WHEN NOT MATCHED THEN INSERT(ParameterName,ParameterValue) VALUES(@N,@V);", new SqlParameter("@N", SqlDbType.VarChar, 100) { Value = name }, new SqlParameter("@V", SqlDbType.NVarChar, 2000) { Value = value.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture) }); }
        private void SetMailRequestStatus(string table, int id, string status) { if (table != "dbo.envoi_mail" && table != "dbo.T_Changement_coordonnees_client_via_le_WEB_envoi_emails") throw new ArgumentException("Invalid table"); ExecuteNonQuery(sql_creation_compte, "UPDATE " + table + " SET top_traite=@S,date_trt=GETDATE() WHERE id=@ID;", new SqlParameter("@S", SqlDbType.Char, 1) { Value = status }, new SqlParameter("@ID", SqlDbType.Int) { Value = id }); }
        private void UpdateOpeningStatus(int dossierId, int status) { ExecuteNonQuery(sql_creation_compte, "UPDATE dbo.T_Statut SET idStatut=@S WHERE idDossier=@ID;", new SqlParameter("@S", SqlDbType.Int) { Value = status }, new SqlParameter("@ID", SqlDbType.Int) { Value = dossierId }); }
        private string SafeResourcePath(string file) { string root = Path.GetFullPath(templatesReportsFolder) + Path.DirectorySeparatorChar; string full = Path.GetFullPath(Path.Combine(root, Path.GetFileName(file ?? ""))); if (!full.StartsWith(root, StringComparison.OrdinalIgnoreCase) || !File.Exists(full)) throw new FileNotFoundException("Resource not found", full); return full; }
        private List<string> ResolveAttachmentPaths(object value) { var list = new List<string>(); if (value == null || value == DBNull.Value) return list; foreach (string n in SplitValues(Convert.ToString(value))) { list.Add(SafeResourcePath(n)); } return list; }
        private static void ValidateNoKnownUnresolvedMarkers(string html, string template) { Match m = Regex.Match(html ?? "", @"\[(NOM|PRENOM|CODE_CLIENT|NOM_CLIENT|RCS|BANQUE|AGENCE|IBAN|BIC|MOTIF|DATE_DEMANDE|MOTIF_REFUS|NUM_DEMANDE|GESTIONNAIRE|NUM_POSTE|MAIL_GESTIONNAIRE)\]", RegexOptions.IgnoreCase); if (m.Success) throw new InvalidDataException("Unresolved marker " + m.Value + " in " + template); }
        private static DataTable FillDataTable(string cs, string sql, params SqlParameter[] parameters) { var t = new DataTable(); using (var c = new SqlConnection(cs)) using (var cmd = new SqlCommand(sql, c)) using (var da = new SqlDataAdapter(cmd)) { cmd.CommandTimeout = 300; if (parameters != null && parameters.Length > 0) cmd.Parameters.AddRange(parameters); da.Fill(t); } return t; }

        private void SendGraphMailWithAttachments(string sendAs, string to, string cc, string bcc, string subject, string html, List<string> files)
        {
            var message = new Message { Subject = subject, Body = new ItemBody { ContentType = BodyType.Html, Content = html }, ToRecipients = BuildRecipients(to), CcRecipients = BuildRecipients(cc), BccRecipients = BuildRecipients(bcc), Attachments = new List<Microsoft.Graph.Models.Attachment>() };
            if (message.ToRecipients.Count == 0) throw new InvalidOperationException("Invalid recipients for " + subject);
            foreach (string path in files ?? new List<string>()) { byte[] bytes = File.ReadAllBytes(path); if (bytes.Length > 3 * 1024 * 1024) throw new InvalidDataException("Attachment exceeds 3 MB : " + Path.GetFileName(path)); message.Attachments.Add(new FileAttachment { OdataType = "#microsoft.graph.fileAttachment", Name = Path.GetFileName(path), ContentType = "application/octet-stream", ContentBytes = bytes }); }
            ExecuteGraphWithRetry(() => { graphService.Users[sendAs].SendMail.PostAsync(new Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody { Message = message, SaveToSentItems = true }).GetAwaiter().GetResult(); return true; }, "Send template/report mail : " + subject);
        }

    }
}
