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
using Microsoft.Graph.Models.ODataErrors;
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
        private const string credit_review_mailbox_name = "Credit Review";
        private const string recovery_reports_mailbox_name = "Rapports Recouvrement";
        private const string credit_reports_mailbox_name = "Rapports Crédit";
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
        private string list_credit_mgr_code = "";
        private string path_archives = "";
        private string dss_con_openrowset_parameter_global = "";
        private string dss_con_openrowset = "";
        private string uri_webservice = "";
        private HashSet<string> credit_managers_contentieux_list = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private string logsFolder = "", tempFolder = "", sessionName = "";
        private string maquettes_rapports_pj_credit = "", fr_ouverture_graph_send_as = "", fr_credit_administration_graph_send_as = "", fr_credit_review_graph_send_as = "", fr_lcrna_graph_send_as = "", fr_facturation_graph_send_as = "";
        private string email_rapport_compteur = "", email_rapport_stat_ouverture = "", admin_ventes = "";
        private List<AutomaticStatementCustomer> automaticStatementCustomers =
            new List<AutomaticStatementCustomer>();
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
            public string list_credit_mgr_code { get; set; } = "";

            public string contentieux_recipients { get; set; } = "";
            public string path_archives { get; set; } = "";
            public string dss_con_openrowset_parameter_global { get; set; } = "";
            public string uri_webservice { get; set; } = "";
            public string maquettes_rapports_pj_credit { get; set; } = "";
            public string fr_ouverture_graph_send_as { get; set; } = "";
            public string fr_credit_administration_graph_send_as { get; set; } = "";
            public string fr_credit_review_graph_send_as { get; set; } = "";
            public string fr_lcrna_graph_send_as { get; set; } = "";
            public string fr_facturation_graph_send_as { get; set; } = "";
            public List<AutomaticStatementCustomer> echeancier_auto_clients
            {
                get;
                set;
            } = new List<AutomaticStatementCustomer>();
            public string email_rapport_compteur { get; set; } = "";
            public string email_rapport_stat_ouverture { get; set; } = "";
            public string admin_ventes { get; set; } = "";
        }

        public sealed class AutomaticStatementCustomer
        {
            public string code_client { get; set; } = "";
            public string email_destinataire { get; set; } = "";
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
            list_credit_mgr_code = i.list_credit_mgr_code ?? "";
            contentieux_recipients = i.contentieux_recipients ?? "";
            path_archives = i.path_archives ?? "";
            dss_con_openrowset_parameter_global = i.dss_con_openrowset_parameter_global ?? "";
            uri_webservice = i.uri_webservice ?? "";
            maquettes_rapports_pj_credit = i.maquettes_rapports_pj_credit ?? "";
            fr_ouverture_graph_send_as = i.fr_ouverture_graph_send_as ?? "";
            fr_credit_administration_graph_send_as = i.fr_credit_administration_graph_send_as ?? "";
            fr_credit_review_graph_send_as = i.fr_credit_review_graph_send_as ?? "";
            fr_lcrna_graph_send_as = i.fr_lcrna_graph_send_as ?? "";
            fr_facturation_graph_send_as =
                i.fr_facturation_graph_send_as ?? "";
            automaticStatementCustomers = i.echeancier_auto_clients ??
                new List<AutomaticStatementCustomer>();
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

        private void UpdateMailboxRefreshInformation(
            int currentMailboxId,
            DateTime? latestProcessedMessageDate)
        {
            const string sql = @"
UPDATE dbo.T_SharedMailboxes
SET date_dernier_raf = GETDATE(),
    dt_heure_filtre = CASE
        WHEN @LATEST_MESSAGE_DATE IS NULL THEN dt_heure_filtre
        WHEN dt_heure_filtre IS NULL THEN @LATEST_MESSAGE_DATE
        WHEN @LATEST_MESSAGE_DATE > dt_heure_filtre THEN @LATEST_MESSAGE_DATE
        ELSE dt_heure_filtre
    END
WHERE id_mailboxe = @MAILBOX_ID;";

            using (var connection = new SqlConnection(sql_connexion))
            using (var command = new SqlCommand(sql, connection))
            {
                command.CommandTimeout = 300;
                command.Parameters.Add(
                    "@MAILBOX_ID",
                    SqlDbType.Int).Value = currentMailboxId;
                command.Parameters.Add(
                    "@LATEST_MESSAGE_DATE",
                    SqlDbType.DateTime).Value =
                    latestProcessedMessageDate.HasValue
                        ? (object)latestProcessedMessageDate.Value
                        : DBNull.Value;

                connection.Open();
                int affectedRows = command.ExecuteNonQuery();

                if (affectedRows == 0)
                {
                    throw new InvalidOperationException(
                        "Mailbox refresh information could not be updated. " +
                        "Mailbox ID : " + currentMailboxId);
                }
            }

            lastRefresh = DateTime.Now;

            if (latestProcessedMessageDate.HasValue &&
                latestProcessedMessageDate.Value > mailboxFilterDate)
            {
                mailboxFilterDate = latestProcessedMessageDate.Value;
            }

            if (IsTrue(debug))
            {
                WriteLog(
                    "       Mailbox refresh information updated" +
                    " - Mailbox ID : " + currentMailboxId +
                    " - Last refresh : " +
                    lastRefresh.ToString("dd/MM/yyyy HH:mm:ss") +
                    " - Filter date : " +
                    mailboxFilterDate.ToString("dd/MM/yyyy HH:mm:ss"));
            }
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
                return ReadOpeningAccountMailbox();
            }

            if (mailboxName.Equals(
                    credit_review_mailbox_name,
                    StringComparison.OrdinalIgnoreCase))
            {
                return ReadCreditReviewMailbox();
            }
            if (mailboxName.Equals(
                    recovery_reports_mailbox_name,
                    StringComparison.OrdinalIgnoreCase))
            {
                return ReadRecoveryReportsMailbox();
            }
            if (mailboxName.Equals(
                    credit_reports_mailbox_name,
                    StringComparison.OrdinalIgnoreCase))
            {
                return ReadCreditReportsMailbox();
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

        private DateTime? ReadCreditReviewMailbox()
        {
            MailFolder input = GetRequiredFolder(inputFolderName);
            MailFolder recoveryFolder = GetRequiredFolder(outputFolderName);
            List<Message> messages = GetCandidateMessages(input.Id);
            DateTime? latest = null;
            int recoveryImported = 0;
            int cardImported = 0;
            int ignored = 0;
            int alreadyProcessed = 0;
            int errors = 0;

            WriteLog(
                "       Credit Review email(s) found for processing : " +
                messages.Count);

            foreach (Message summary in messages)
            {
                DateTime received =
                    summary.ReceivedDateTime?.LocalDateTime ?? DateTime.Now;
                if (!latest.HasValue || received > latest.Value)
                    latest = received;

                try
                {
                    if (IsMessageAlreadyProcessed(summary) ||
                        IsProcessingStateSuccessful(summary.Id))
                    {
                        alreadyProcessed++;
                        continue;
                    }

                    Message email = GetCompleteMessage(summary.Id);
                    string subject = (email.Subject ?? "").Trim();
                    WriteLog("       Credit Review subject : " + subject);

                    if (IsRecoveryCreditReviewSubject(subject))
                    {
                        int historyId = ResolveRecoveryHistoryId(email);
                        if (historyId <= 0)
                        {
                            throw new InvalidOperationException(
                                "No pending recovery history found for subject : " +
                                subject);
                        }

                        byte[] mime = GetMimeContent(email.Id);
                        SaveRecoveryHistoryMail(historyId, email, mime);
                        TryMarkMessageAsProcessed(email.Id);
                        MoveMessage(email.Id, recoveryFolder.Id, "Credit Review recovery");
                        UpsertProcessingState(
                            email,
                            "S",
                            null,
                            "Recovery EML inserted for history " + historyId);
                        recoveryImported++;
                        continue;
                    }

                    if (IsCreditCardHistorySubject(subject))
                    {
                        int historyId = GetCreditCardHistoryId(subject);
                        if (historyId <= 0)
                        {
                            throw new InvalidOperationException(
                                "No pending credit card history found for subject : " +
                                subject);
                        }

                        byte[] mime = GetMimeContent(email.Id);
                        SaveCreditCardHistoryMail(historyId, email, mime);
                        TryMarkMessageAsProcessed(email.Id);
                        PermanentlyDeleteGenericMessage(
                            email.Id,
                            "Credit Review card history");
                        UpsertProcessingState(
                            email,
                            "S",
                            null,
                            "Credit card EML inserted for history " + historyId);
                        cardImported++;
                        continue;
                    }

                    UpsertProcessingState(
                        email,
                        "I",
                        null,
                        "Subject outside Credit Review rules");
                    ignored++;
                }
                catch (Exception ex)
                {
                    errors++;
                    UpsertProcessingState(summary, "E", null, ex.Message);
                    WriteLog(
                        "       Credit Review email processing error" +
                        " - Subject : " + (summary.Subject ?? "<no subject>") +
                        " - Error : " + ex.Message);
                    SendTechnicalAlert(
                        nameof(ReadCreditReviewMailbox),
                        mailboxAddress + " - Mail : " +
                        (summary.Subject ?? "<no subject>") + " - " + ex.Message,
                        "CREDIT REVIEW EMAIL PROCESSING");
                }
            }

            WriteLog(
                "       Credit Review processing summary" +
                " - Found : " + messages.Count +
                " - Recovery imported : " + recoveryImported +
                " - Card imported : " + cardImported +
                " - Ignored : " + ignored +
                " - Already processed : " + alreadyProcessed +
                " - Errors : " + errors);
            return latest;
        }

        private DateTime? ReadRecoveryReportsMailbox()
        {
            return ReadStoredReportMailbox(false);
        }

        private DateTime? ReadCreditReportsMailbox()
        {
            return ReadStoredReportMailbox(true);
        }

        private DateTime? ReadStoredReportMailbox(bool creditReport)
        {
            MailFolder input = GetRequiredFolder(inputFolderName);
            List<Message> messages = GetCandidateMessages(input.Id);
            DateTime? latest = null;
            int inserted = 0;
            int requestUpdates = 0;
            int ignored = 0;
            int alreadyProcessed = 0;
            int errors = 0;

            WriteLog(
                "       " + (creditReport ? "Credit" : "Recovery") +
                " report email(s) found for processing : " + messages.Count);

            foreach (Message summary in messages)
            {
                DateTime received =
                    summary.ReceivedDateTime?.LocalDateTime ?? DateTime.Now;
                if (!latest.HasValue || received > latest.Value)
                    latest = received;

                try
                {
                    if (IsMessageAlreadyProcessed(summary) ||
                        IsProcessingStateSuccessful(summary.Id))
                    {
                        alreadyProcessed++;
                        continue;
                    }

                    Message email = GetCompleteMessage(summary.Id);
                    string subject = (email.Subject ?? "").Trim();
                    if (string.IsNullOrWhiteSpace(subject))
                    {
                        UpsertProcessingState(
                            email,
                            "I",
                            null,
                            "Message subject is empty");
                        ignored++;
                        continue;
                    }

                    byte[] mime = GetMimeContent(email.Id);
                    if (!creditReport &&
                        TryExtractCustomerChangeRequestId(subject, out int requestId))
                    {
                        UpdateCustomerChangeSentMail(requestId, mime);
                        requestUpdates++;
                        WriteLog(
                            "       Recovery report linked to customer change request" +
                            " - Request : " + requestId +
                            " - Subject : " + subject);
                    }
                    else
                    {
                        InsertReceivedReportMail(creditReport, email, mime);
                        inserted++;
                        WriteLog(
                            "       " + (creditReport ? "Credit" : "Recovery") +
                            " report inserted in database" +
                            " - Subject : " + subject);
                    }

                    TryMarkMessageAsProcessed(email.Id);
                    PermanentlyDeleteGenericMessage(
                        email.Id,
                        creditReport ? "Credit report" : "Recovery report");
                    UpsertProcessingState(
                        email,
                        "S",
                        null,
                        creditReport
                            ? "Credit report EML inserted"
                            : "Recovery report EML processed");
                }
                catch (Exception ex)
                {
                    errors++;
                    UpsertProcessingState(summary, "E", null, ex.Message);
                    WriteLog(
                        "       Report email processing error" +
                        " - Subject : " + (summary.Subject ?? "<no subject>") +
                        " - Error : " + ex.Message);
                    SendTechnicalAlert(
                        creditReport
                            ? nameof(ReadCreditReportsMailbox)
                            : nameof(ReadRecoveryReportsMailbox),
                        mailboxAddress + " - Mail : " +
                        (summary.Subject ?? "<no subject>") + " - " + ex.Message,
                        creditReport
                            ? "CREDIT REPORT EMAIL PROCESSING"
                            : "RECOVERY REPORT EMAIL PROCESSING");
                }
            }

            WriteLog(
                "       " + (creditReport ? "Credit" : "Recovery") +
                " report processing summary" +
                " - Found : " + messages.Count +
                " - Inserted : " + inserted +
                " - Customer request updates : " + requestUpdates +
                " - Ignored : " + ignored +
                " - Already processed : " + alreadyProcessed +
                " - Errors : " + errors);
            return latest;
        }

        private static bool IsRecoveryCreditReviewSubject(string subject)
        {
            string value = subject ?? "";
            return value.IndexOf("Etat de Compte", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   value.StartsWith("RED ALERT -", StringComparison.OrdinalIgnoreCase) ||
                   value.IndexOf("Vos virements Ingram Micro", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   value.StartsWith("Arrêté de compte Ingram micro", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsCreditCardHistorySubject(string subject)
        {
            string value = subject ?? "";
            return value.IndexOf("Solde débiteur au compte", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   value.IndexOf("Plus d'encours et factures échues", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   value.IndexOf("Attente complément réglement", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   value.IndexOf("Lettrage Trésorie", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   value.IndexOf("Déblocage Ventes", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   value.IndexOf("Multi BL (11/21)", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private int ResolveRecoveryHistoryId(Message email)
        {
            string subject = (email.Subject ?? "").Trim();
            if (subject.StartsWith("RED ALERT -", StringComparison.OrdinalIgnoreCase))
            {
                string customer = Left(
                    subject.Substring("RED ALERT -".Length).Trim(),
                    8);
                return GetRecoveryHistoryId(customer, true, false);
            }

            if (subject.IndexOf(
                    "Vos virements Ingram Micro",
                    StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return GetRecoveryHistoryId(Left(subject, 8), false, true);
            }

            string customerNumber = GetCustomerNumberFromAttachments(email.Id);
            return GetRecoveryHistoryId(customerNumber, false, false);
        }

        private string GetCustomerNumberFromAttachments(string messageId)
        {
            AttachmentCollectionResponse response = ExecuteGraphWithRetry(
                () => graphService.Users[mailboxAddress]
                    .Messages[messageId]
                    .Attachments
                    .GetAsync(config =>
                    {
                        AddImmutableHeader(config.Headers);
                        config.QueryParameters.Top = 999;
                    })
                    .GetAwaiter()
                    .GetResult(),
                "Read Credit Review attachments");

            string fallback = "";
            foreach (Microsoft.Graph.Models.Attachment attachment
                in response?.Value ?? new List<Microsoft.Graph.Models.Attachment>())
            {
                string name = Path.GetFileNameWithoutExtension(
                    attachment?.Name ?? "").Trim();
                if (string.IsNullOrWhiteSpace(name))
                    continue;
                fallback = name;
                const string prefix = "Detail_Compte_";
                if (name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    return name.Substring(prefix.Length).Trim();
            }
            return fallback;
        }

        private int GetRecoveryHistoryId(
            string customerNumber,
            bool isRedAlert,
            bool isAutomatch)
        {
            if (string.IsNullOrWhiteSpace(customerNumber))
                return 0;

            string statusClause;
            if (isRedAlert)
                statusClause = "= 10";
            else if (isAutomatch)
                statusClause = "= 55";
            else
                statusClause = "IN (20,30,40,90)";

            string sql =
                "SELECT TOP (1) id_histo " +
                "FROM dbo.T_Credit_Recouvrement_HISTO " +
                "WHERE fichier IS NULL " +
                "AND id_statut " + statusClause + " " +
                "AND br_cust_nbr=@CUSTOMER;";

            string result = ExecuteScalarString(
                sql_connexion,
                sql,
                new SqlParameter("@CUSTOMER", SqlDbType.VarChar, 50)
                {
                    Value = customerNumber.Trim()
                });
            return int.TryParse(result, out int id) ? id : 0;
        }

        private int GetCreditCardHistoryId(string subject)
        {
            string result = ExecuteScalarString(
                sql_gestion_cdes,
                @"SELECT TOP (1) id_histo
                  FROM dbo.T_deblocage_carte_bleue_histo_mail
                  WHERE fichier IS NULL
                    AND sujet_email=@SUBJECT
                  ORDER BY id_histo DESC;",
                new SqlParameter("@SUBJECT", SqlDbType.NVarChar, 1000)
                {
                    Value = subject ?? ""
                });
            return int.TryParse(result, out int id) ? id : 0;
        }

        private void SaveRecoveryHistoryMail(
            int historyId,
            Message email,
            byte[] mime)
        {
            ExecuteStoredMailProcedure(
                sql_connexion,
                "dbo.USP_ADD_FIC_RECOUVREMENT",
                historyId,
                email,
                mime);
        }

        private void SaveCreditCardHistoryMail(
            int historyId,
            Message email,
            byte[] mime)
        {
            ExecuteStoredMailProcedure(
                sql_gestion_cdes,
                "dbo.USP_ADD_FIC_deblocage_carte_bleue_histo_mail",
                historyId,
                email,
                mime);
        }

        private static void ExecuteStoredMailProcedure(
            string connectionString,
            string procedure,
            int historyId,
            Message email,
            byte[] mime)
        {
            if (mime == null || mime.Length == 0)
                throw new InvalidDataException("MIME content is empty");

            string received =
                (email.ReceivedDateTime?.LocalDateTime ?? DateTime.Now)
                .ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture);
            string fileName = "EMAIL_" + received;

            using (var connection = new SqlConnection(connectionString))
            using (var command = new SqlCommand(procedure, connection))
            {
                command.CommandType = CommandType.StoredProcedure;
                command.CommandTimeout = 300;
                command.Parameters.Add("@id_histo", SqlDbType.NVarChar, 50)
                    .Value = historyId.ToString(CultureInfo.InvariantCulture);
                command.Parameters.Add("@nom_fic", SqlDbType.NVarChar, 500)
                    .Value = fileName;
                command.Parameters.Add("@document", SqlDbType.Image).Value = mime;
                command.Parameters.Add("@extension", SqlDbType.NVarChar, 20)
                    .Value = "eml";
                connection.Open();
                command.ExecuteNonQuery();
            }
        }

        private void InsertReceivedReportMail(
            bool creditReport,
            Message email,
            byte[] mime)
        {
            if (mime == null || mime.Length == 0)
                throw new InvalidDataException("MIME content is empty");

            string table = creditReport
                ? "dbo.T_Credit_Recouvrement_MAIL_RAPPORTS_CREDIT_RECUS"
                : "dbo.T_Credit_Recouvrement_MAIL_RECOUVREMENT_RECUS";
            DateTime received =
                email.ReceivedDateTime?.LocalDateTime ?? DateTime.Now;
            string fileName = "EMAIL_" + received.ToString(
                "yyyyMMdd_HHmmss",
                CultureInfo.InvariantCulture);

            ExecuteNonQuery(
                sql_connexion,
                "INSERT INTO " + table +
                "([date],objet,nom_fichier,extension,fichier) " +
                "VALUES(@DATE,@SUBJECT,@NAME,'eml',@FILE);",
                new SqlParameter("@DATE", SqlDbType.DateTime)
                {
                    Value = received
                },
                new SqlParameter("@SUBJECT", SqlDbType.NVarChar, -1)
                {
                    Value = email.Subject ?? ""
                },
                new SqlParameter("@NAME", SqlDbType.NVarChar, 500)
                {
                    Value = fileName
                },
                new SqlParameter("@FILE", SqlDbType.Image)
                {
                    Value = mime
                });
        }

        private static bool TryExtractCustomerChangeRequestId(
            string subject,
            out int requestId)
        {
            requestId = 0;
            Match match = Regex.Match(
                subject ?? "",
                @"^(?:Acceptation|Refus) de votre demande n°\s*(\d+)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            return match.Success &&
                   int.TryParse(
                       match.Groups[1].Value,
                       NumberStyles.None,
                       CultureInfo.InvariantCulture,
                       out requestId);
        }

        private void UpdateCustomerChangeSentMail(int requestId, byte[] mime)
        {
            if (mime == null || mime.Length == 0)
                throw new InvalidDataException("MIME content is empty");

            int affected;
            using (var connection = new SqlConnection(sql_creation_compte))
            using (var command = new SqlCommand(
                @"UPDATE dbo.T_Changement_coordonnees_client_via_le_WEB
                  SET mail_envoye_au_client=@MAIL
                  WHERE id_demande=@ID;",
                connection))
            {
                command.CommandTimeout = 300;
                command.Parameters.Add("@MAIL", SqlDbType.Image).Value = mime;
                command.Parameters.Add("@ID", SqlDbType.Int).Value = requestId;
                connection.Open();
                affected = command.ExecuteNonQuery();
            }

            if (affected == 0)
            {
                throw new InvalidOperationException(
                    "Customer change request not found : " + requestId);
            }
        }

        private bool IsProcessingStateSuccessful(string graphMessageId)
        {
            string count = ExecuteScalarString(
                sql_connexion,
                @"SELECT COUNT(*)
                  FROM dbo.T_CreditMailProcessing
                  WHERE MailboxId=@MAILBOX
                    AND GraphMessageId=@MESSAGE
                    AND ProcessingStatus='S';",
                new SqlParameter("@MAILBOX", SqlDbType.Int)
                {
                    Value = mailboxId
                },
                new SqlParameter("@MESSAGE", SqlDbType.NVarChar, 500)
                {
                    Value = graphMessageId ?? ""
                });
            return int.TryParse(count, out int value) && value > 0;
        }

        private void MoveMessage(
            string messageId,
            string destinationFolderId,
            string operation)
        {
            ExecuteGraphWithRetry(() =>
            {
                graphService.Users[mailboxAddress]
                    .Messages[messageId]
                    .Move
                    .PostAsync(
                        new Microsoft.Graph.Users.Item.Messages.Item.Move.MovePostRequestBody
                        {
                            DestinationId = destinationFolderId
                        },
                        config => AddImmutableHeader(config.Headers))
                    .GetAwaiter()
                    .GetResult();
                return true;
            }, "Move " + operation + " message");
        }

        private void PermanentlyDeleteGenericMessage(
            string messageId,
            string operation)
        {
            ExecuteGraphWithRetry(() =>
            {
                graphService.Users[mailboxAddress]
                    .Messages[messageId]
                    .PermanentDelete
                    .PostAsync(config => AddImmutableHeader(config.Headers))
                    .GetAwaiter()
                    .GetResult();
                return true;
            }, "Permanently delete " + operation + " message");
        }

        private DateTime? ReadOpeningAccountMailbox()
        {
            if (string.IsNullOrWhiteSpace(sql_creation_compte))
            {
                throw new InvalidOperationException(
                    "sql_creation_compte is empty");
            }

            MailFolder inbox = GetRequiredFolder(inputFolderName);
            List<Message> messages = GetOpeningAccountMessages(inbox.Id);
            DateTime? latestHandled = null;
            int imported = 0;
            int ignored = 0;
            int alreadyProcessed = 0;
            int errors = 0;

            foreach (Message summary in messages)
            {
                DateTime received =
                    summary.ReceivedDateTime?.LocalDateTime ?? DateTime.Now;

                try
                {
                    // Un statut SQL S signifie que l'import metier a reussi.
                    // Si le message est encore present, seule sa suppression est
                    // rejouee, sans nouvel import en base.
                    if (IsProcessingStateSuccessful(summary.Id))
                    {
                        PermanentlyDeleteOpeningMessage(summary.Id);

                        if (!latestHandled.HasValue ||
                            received > latestHandled.Value)
                        {
                            latestHandled = received;
                        }

                        alreadyProcessed++;
                        WriteLog(
                            "       Opening account email already imported " +
                            "and now permanently deleted : " +
                            (summary.Subject ?? "<no subject>"));
                        continue;
                    }

                    // Une propriete etendue sans statut SQL S correspond a un
                    // message hors criteres. Il reste conserve et non lu dans Inbox.
                    if (IsOpeningMessageAlreadyProcessed(summary))
                    {
                        if (!latestHandled.HasValue ||
                            received > latestHandled.Value)
                        {
                            latestHandled = received;
                        }

                        ignored++;
                        WriteLog(
                            "       Opening account email outside robot criteria " +
                            "already tagged; kept unread in Inbox : " +
                            (summary.Subject ?? "<no subject>"));
                        continue;
                    }

                    Message email = GetOpeningAccountMessage(summary.Id);
                    string sender = GetSender(email).Trim();
                    string subject = (email.Subject ?? "").Trim();
                    string dossierIdText = "";
                    bool matched = false;

                    if (sender.Equals(
                            "ouverture@ingrammicro.fr",
                            StringComparison.OrdinalIgnoreCase) &&
                        subject.Equals(
                            "ouverture de compte client ingram micro",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        dossierIdText =
                            GetOpeningDossierIdFromPdfAttachment(email.Id);
                        matched = !string.IsNullOrWhiteSpace(dossierIdText);
                    }
                    else if (sender.Equals(
                                 "analystes.credit@ingrammicro.fr",
                                 StringComparison.OrdinalIgnoreCase))
                    {
                        dossierIdText =
                            GetOpeningDossierIdFromCustomerNumber(subject);
                        matched = !string.IsNullOrWhiteSpace(dossierIdText);
                    }

                    received =
                        email.ReceivedDateTime?.LocalDateTime ??
                        summary.ReceivedDateTime?.LocalDateTime ??
                        DateTime.Now;

                    if (!matched)
                    {
                        // Hors criteres : ajout uniquement de la propriete
                        // etendue IMCA afin de ne pas retraiter le message.
                        // Le message reste non lu et conserve dans Inbox.
                        TryMarkOpeningMessageAsProcessed(
                            email.Id,
                            false);

                        UpsertProcessingState(
                            email,
                            "I",
                            null,
                            "Opening account email outside robot criteria - " +
                            "tagged and kept unread in Inbox");

                        if (!latestHandled.HasValue ||
                            received > latestHandled.Value)
                        {
                            latestHandled = received;
                        }

                        ignored++;
                        WriteLog(
                            "       Opening account email outside robot criteria; " +
                            "extended property added, left unread and kept in Inbox" +
                            " - Sender : " + sender +
                            " - Subject : " + subject);
                        continue;
                    }

                    if (!int.TryParse(
                            dossierIdText,
                            NumberStyles.None,
                            CultureInfo.InvariantCulture,
                            out int dossierId) ||
                        dossierId <= 0)
                    {
                        throw new InvalidOperationException(
                            "Opening account business email matched but no valid " +
                            "dossier ID could be resolved. Sender : " +
                            sender + " - Subject : " + subject);
                    }

                    byte[] mime = GetMimeContent(email.Id);
                    InsertOpeningMailInDatabase(dossierId, email, mime);
                    UpdateOpeningUploadStatus(dossierId);

                    // Le succes metier est enregistre avant le nettoyage
                    // Exchange. En cas d'echec de suppression, le prochain passage
                    // rejouera uniquement la suppression, sans reinserer le MIME.
                    UpsertProcessingState(
                        email,
                        "S",
                        null,
                        "Opening account email imported for dossier " +
                        dossierId);

                    TryMarkOpeningMessageAsProcessed(
                        email.Id,
                        true);

                    PermanentlyDeleteOpeningMessage(email.Id);

                    if (!latestHandled.HasValue ||
                        received > latestHandled.Value)
                    {
                        latestHandled = received;
                    }

                    imported++;
                    WriteLog(
                        "       Opening account email imported directly from Inbox" +
                        " - Dossier : " + dossierId +
                        " - Message permanently deleted after database import" +
                        " - Subject : " +
                        (email.Subject ?? "<no subject>"));
                }
                catch (Exception ex)
                {
                    errors++;
                    UpsertProcessingState(summary, "E", null, ex.ToString());
                    WriteLog(
                        "       Opening account email processing error" +
                        " - Mailbox message left unchanged" +
                        " - Subject : " +
                        (summary.Subject ?? "<no subject>") +
                        " - Error : " + ex);
                    SendTechnicalAlert(
                        nameof(ReadOpeningAccountMailbox),
                        mailboxAddress + " - Mail : " +
                        (summary.Subject ?? "<no subject>") + " - " + ex,
                        "OPENING ACCOUNT EMAIL PROCESSING");
                }
            }

            WriteLog(
                "       Opening account summary" +
                " - Candidates : " + messages.Count +
                " - Imported directly from Inbox : " + imported +
                " - Ignored and left unchanged : " + ignored +
                " - Already processed and left unchanged : " +
                alreadyProcessed +
                " - Processing errors : " + errors);

            return latestHandled;
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
            return ExecuteGraphWithRetry(
                () => graphService.Users[mailboxAddress]
                    .Messages[id]
                    .GetAsync(q =>
                    {
                        AddImmutableHeader(q.Headers);
                        q.QueryParameters.Select = new[]
                        {
                            "id",
                            "subject",
                            "receivedDateTime",
                            "createdDateTime",
                            "from",
                            "sender",
                            "toRecipients",
                            "hasAttachments",
                            "internetMessageId"
                        };
                    })
                    .GetAwaiter()
                    .GetResult(),
                "Get opening account message");
        }

        private static bool IsOpeningMessageAlreadyProcessed(Message m)
        {
            return m?.SingleValueExtendedProperties?.Any(x => string.Equals(x.Id, opening_processed_property_id, StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(x.Value)) == true;
        }

        private void TryMarkOpeningMessageAsProcessed(
            string id,
            bool markAsRead)
        {
            Exception lastException = null;

            for (int attempt = 1; attempt <= 3; attempt++)
            {
                try
                {
                    var update = new Message
                    {
                        IsRead = markAsRead
                            ? (bool?)true
                            : null,

                        SingleValueExtendedProperties =
                            new List<SingleValueLegacyExtendedProperty>
                            {
                        new SingleValueLegacyExtendedProperty
                        {
                            Id = opening_processed_property_id,
                            Value = DateTime.UtcNow.ToString(
                                "yyyy-MM-ddTHH:mm:ss.fffZ",
                                CultureInfo.InvariantCulture)
                        }
                            }
                    };

                    graphService.Users[mailboxAddress]
                        .Messages[id]
                        .PatchAsync(
                            update,
                            configuration =>
                                AddImmutableHeader(
                                    configuration.Headers))
                        .GetAwaiter()
                        .GetResult();

                    return;
                }
                catch (Exception ex)
                    when (IsChangeKeyConflict(ex))
                {
                    lastException = ex;

                    if (attempt >= 3)
                    {
                        break;
                    }

                    WriteLog(
                        "       Opening account change key conflict" +
                        " - Message ID : " + id +
                        " - Attempt : " + attempt + "/3");

                    System.Threading.Thread.Sleep(
                        attempt * 500);
                }
            }

            throw new InvalidOperationException(
                "Unable to mark opening account message " +
                "after change key retries" +
                " - Message ID : " + id +
                " - Error : " +
                GetInnermostExceptionMessage(lastException),
                lastException);
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
                        .PostAsync(configuration =>
                            AddImmutableHeader(configuration.Headers))
                        .GetAwaiter()
                        .GetResult();
                    return true;
                },
                "Permanently delete imported opening account message");
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

            WriteLog(
            "       Listing Graph messages" +
            " - Mailbox : " +
            mailboxAddress +
            " - Folder : " +
            inputFolderName +
            " - Filter date : " +
            date.ToString("dd/MM/yyyy HH:mm:ss") +
            " - Top : " +
            top);

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
            string details =
                GetDetailedExceptionMessage(ex);

            WriteLog(
                "   Error reading mailbox" +
                " - Mailbox : " +
                mailboxAddress +
                " - Folder : " +
                inputFolderName +
                " - Details : " +
                details);

            SendTechnicalAlert(
                nameof(Read_Email_with_Graph),
                mailboxAddress +
                " - Folder : " +
                inputFolderName +
                " - " +
                details,
                "MAILBOX PROCESSING");
        }

        private static string GetDetailedExceptionMessage(
    Exception exception)
        {
            if (exception == null)
            {
                return "Unknown error";
            }

            List<string> details =
                new List<string>();

            Exception current =
                exception;

            while (current != null)
            {
                details.Add(
                    "Type=" +
                    current.GetType().FullName);

                if (!string.IsNullOrWhiteSpace(
                        current.Message))
                {
                    details.Add(
                        "Message=" +
                        current.Message);
                }

                ODataError graphError =
                    current as ODataError;

                if (graphError != null)
                {
                    if (!string.IsNullOrWhiteSpace(
                            graphError.Error?.Code))
                    {
                        details.Add(
                            "GraphCode=" +
                            graphError.Error.Code);
                    }

                    if (!string.IsNullOrWhiteSpace(
                            graphError.Error?.Message))
                    {
                        details.Add(
                            "GraphMessage=" +
                            graphError.Error.Message);
                    }
                }

                if (current is ApiException apiException)
                {
                    details.Add(
                        "HttpStatus=" +
                        apiException.ResponseStatusCode);
                }

                current =
                    current.InnerException;
            }

            return string.Join(
                " | ",
                details.Distinct(
                    StringComparer.OrdinalIgnoreCase));
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
            if (mailboxId <= 0 ||
                string.IsNullOrWhiteSpace(mailboxAddress) ||
                string.IsNullOrWhiteSpace(inputFolderName))
            {
                throw new InvalidOperationException(
                    "Mailbox configuration is incomplete");
            }

            bool outputFolderRequired =
                mailboxName.Equals(
                    credit_card_mailbox_name,
                    StringComparison.OrdinalIgnoreCase) ||
                mailboxName.Equals(
                    score_fraud_mailbox_name,
                    StringComparison.OrdinalIgnoreCase) ||
                mailboxName.Equals(
                    credit_review_mailbox_name,
                    StringComparison.OrdinalIgnoreCase);

            if (outputFolderRequired &&
                string.IsNullOrWhiteSpace(outputFolderName))
            {
                throw new InvalidOperationException(
                    "folder_out is required for mailbox " + mailboxName);
            }

            if (mailboxName.Equals(
                    credit_card_mailbox_name,
                    StringComparison.OrdinalIgnoreCase) &&
                allowedSenders.Count == 0)
            {
                throw new InvalidOperationException(
                    "sender_allowed is empty for the Cartes Bleues mailbox");
            }
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
                    (apiException.ResponseStatusCode == 408 ||
                     apiException.ResponseStatusCode == 429 ||
                     apiException.ResponseStatusCode == 500 ||
                     apiException.ResponseStatusCode == 502 ||
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

        public void Extract_Data_Recouvrement_Ar_Open(
            string sql_con,
            string logs,
            string tmp_folder,
            string session_name)
        {
            const string executionParameter =
                "date_dernier_extract_data_recouvrement_ar_open";

            string root = GetServicePath();
            logsFolder = Path.Combine(root, logs ?? "");
            tempFolder = Path.Combine(root, tmp_folder ?? "");
            sessionName = session_name ?? "";
            Directory.CreateDirectory(logsFolder);
            Directory.CreateDirectory(tempFolder);

            JsonFile configuration = JsonConvert.DeserializeObject<JsonFile>(
                GetImcaParameter(sql_con, global_application_name) ?? "");
            if (configuration?.countries == null || configuration.countries.Count == 0)
                throw new InvalidOperationException(
                    global_application_name + " parameters are empty or invalid");

            foreach (Country item in configuration.countries)
            {
                ApplyCountryConfiguration(item, sql_con);
                if (!IsTrue(active) ||
                    !country.Equals("FR", StringComparison.OrdinalIgnoreCase))
                    continue;

                try
                {
                    ValidateArOpenExtractionConfiguration();

                    DateTime? arOpenUpdatedAt;
                    string sourceDetails;
                    if (!IsArOpenUpdatedToday(out arOpenUpdatedAt, out sourceDetails))
                    {
                        WriteLog(
                            "   AR Open recovery extraction skipped" +
                            " - Source ar_open is not updated today" +
                            " - Details : " + sourceDetails);
                        continue;
                    }

                    DateTime lastExecution =
                        GetOptionalExecutionParameterDate(executionParameter);
                    if (lastExecution.Date >= DateTime.Today)
                    {
                        WriteLog(
                            "   AR Open recovery extraction already completed today" +
                            " - Last successful execution : " +
                            lastExecution.ToString("dd/MM/yyyy HH:mm:ss") +
                            " - Source : " + sourceDetails);
                        continue;
                    }

                    WriteLog(
                        "   Starting AR Open recovery extraction" +
                        " - Source : " + sourceDetails);

                    int importedRows = ProcessArOpenRecoveryExtraction();

                    SetCreditReviewScheduleDate(
                        executionParameter,
                        DateTime.Now);

                    WriteLog(
                        "   AR Open recovery extraction completed" +
                        " - Imported row(s) : " + importedRows +
                        " - Source update : " +
                        arOpenUpdatedAt.Value.ToString("dd/MM/yyyy HH:mm:ss"));
                }
                catch (Exception ex)
                {
                    string details = GetDetailedExceptionMessage(ex);
                    WriteLog(
                        "   AR Open recovery extraction error" +
                        " - Execution parameter not updated" +
                        " - Details : " + details);

                    try
                    {
                        if (graphService == null)
                        {
                            WriteLog(
                                "   Connecting to Microsoft Graph for " +
                                "AR Open technical alert");
                            graphService = ConnectGraph();
                        }

                        SendTechnicalAlert(
                            nameof(Extract_Data_Recouvrement_Ar_Open),
                            details,
                            "AR OPEN RECOVERY EXTRACTION");

                        WriteLog(
                            "   AR Open technical alert sent");
                    }
                    catch (Exception alertException)
                    {
                        WriteLog(
                            "   AR Open technical alert could not be sent" +
                            " - Details : " +
                            GetDetailedExceptionMessage(alertException));
                    }
                }
                finally
                {
                    graphService = null;
                }
            }
        }

        private void ValidateArOpenExtractionConfiguration()
        {
            if (string.IsNullOrWhiteSpace(sql_connexion))
                throw new InvalidOperationException(
                    "sql_connexion is empty for AR Open extraction");
            if (string.IsNullOrWhiteSpace(sql_dss_copie))
                throw new InvalidOperationException(
                    "sql_dss_copie is empty. Configure " +
                    "sql_dss_copie_parameter_global=FR_SQLCON_DSS_COPIE");
            if (string.IsNullOrWhiteSpace(sql_creation_compte))
                throw new InvalidOperationException(
                    "sql_creation_compte is empty for AR Open scheduling");
            if (ParseCreditManagerCodes().Count == 0)
                throw new InvalidOperationException(
                    "list_credit_mgr_code is empty for AR Open extraction");
        }

        private bool IsArOpenUpdatedToday(
            out DateTime? updatedAt,
            out string details)
        {
            updatedAt = null;
            details = "";
            DataTable table = FillDataTable(
                sql_dss_copie,
                @"SELECT TOP (1)
                         id_table,
                         date_maj,
                         id_statut,
                         ISNULL(cmt,'') AS cmt
                  FROM dbo.t_maj_tables
                  WHERE table_name=N'ar_open'
                    AND actif=1
                  ORDER BY date_maj DESC,id_table DESC;");
            if (table.Rows.Count == 0)
            {
                details = "No active t_maj_tables row found for ar_open";
                return false;
            }
            DataRow row = table.Rows[0];
            if (row["date_maj"] == DBNull.Value)
            {
                details = "ar_open date_maj is NULL";
                return false;
            }
            DateTime dateMaj = Convert.ToDateTime(row["date_maj"]);
            updatedAt = dateMaj;
            details =
                "ID : " + Convert.ToString(row["id_table"]) +
                " - Date : " + dateMaj.ToString("dd/MM/yyyy HH:mm:ss") +
                " - Status : " + Convert.ToString(row["id_statut"]) +
                " - Comment : " + Convert.ToString(row["cmt"]);
            return dateMaj.Date == DateTime.Today;
        }

        private DateTime GetOptionalExecutionParameterDate(string parameterName)
        {
            string value = ExecuteScalarString(
                sql_creation_compte,
                @"SELECT ISNULL(ParameterValue,'')
                  FROM dbo.T_STATS_OUVERTURE_COMPTEUR_PARAMETERS
                  WHERE ParameterName=@NAME;",
                new SqlParameter("@NAME", SqlDbType.VarChar, 100)
                {
                    Value = parameterName
                });
            if (string.IsNullOrWhiteSpace(value))
                return new DateTime(1900, 1, 1);
            DateTime parsed;
            if (!DateTime.TryParse(
                    value,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AssumeLocal,
                    out parsed))
                throw new InvalidDataException(
                    "Invalid execution parameter " + parameterName + " : " + value);
            return parsed;
        }

        private List<string> ParseCreditManagerCodes()
        {
            return SplitQuotedValues(list_credit_mgr_code)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private int ProcessArOpenRecoveryExtraction()
        {
            List<string> managerCodes = ParseCreditManagerCodes();
            var parameterNames = new List<string>();
            for (int index = 0; index < managerCodes.Count; index++)
                parameterNames.Add("@MGR" + index.ToString(CultureInfo.InvariantCulture));

            string sourceSql = @"
SELECT DISTINCT
       ar_open.brcustnbr,
       customer.cust_name,
       customer.credit_mgr_code,
       CASE
           WHEN ar_open.PaySeqNbr IS NULL THEN ar_open.BrInvoiceNbr
           ELSE ar_open.BrInvoiceNbr + RIGHT(ar_open.PaySeqNbr,2)
       END AS [N° de Piéce],
       CAST(ar_open.InvoiceDt AS smalldatetime) AS [Date Piéce],
       CAST(ar_open.DueDt AS smalldatetime) AS Date_Echéance,
       ar_open.CustOrderNbr AS [Votre N° de commande],
       CASE
           WHEN ar_open.CurrencyCd='USD' THEN ar_open.ForeignSalesAmt
           ELSE ar_open.TotalSalesAmt
       END AS [Montant TTC],
       ar_open.SourceCd,
       customer.TotalBalanceAmt,
       customer.credit_limit
FROM dbo.ar_open AS ar_open
INNER JOIN dbo.customer AS customer
    ON ar_open.BrCustNbr=customer.branch_customer_nbr
INNER JOIN dbo.customer_location AS customer_location
    ON customer_location.branch_customer_nbr=customer.branch_customer_nbr
WHERE ar_open.CompanyCd='FR'
  AND LEFT(ar_open.BrCustNbr,2) IN ('15','21','24')
  AND customer.credit_mgr_code IN (" +
                string.Join(",", parameterNames) + @")
  AND customer_location.suffix='000';";

            using (var sourceConnection = new SqlConnection(sql_dss_copie))
            using (var sourceCommand = new SqlCommand(sourceSql, sourceConnection))
            using (var targetConnection = new SqlConnection(sql_connexion))
            {
                sourceCommand.CommandTimeout = 0;
                for (int index = 0; index < managerCodes.Count; index++)
                    sourceCommand.Parameters.Add(
                        parameterNames[index],
                        SqlDbType.VarChar,
                        20).Value = managerCodes[index];

                sourceConnection.Open();
                targetConnection.Open();
                using (SqlTransaction transaction = targetConnection.BeginTransaction())
                {
                    try
                    {
                        using (var truncate = new SqlCommand(
                            "TRUNCATE TABLE dbo.T_Credit_Recouvrement_AR_OPEN;",
                            targetConnection,
                            transaction))
                        {
                            truncate.CommandTimeout = 300;
                            truncate.ExecuteNonQuery();
                        }

                        int importedRows = 0;
                        using (SqlDataReader reader = sourceCommand.ExecuteReader())
                        using (var bulkCopy = new SqlBulkCopy(
                            targetConnection,
                            SqlBulkCopyOptions.TableLock,
                            transaction))
                        {
                            bulkCopy.DestinationTableName =
                                "dbo.T_Credit_Recouvrement_AR_OPEN";
                            bulkCopy.BulkCopyTimeout = 0;
                            bulkCopy.ColumnMappings.Add("brcustnbr", "brcustnbr");
                            bulkCopy.ColumnMappings.Add("cust_name", "cust_name");
                            bulkCopy.ColumnMappings.Add("credit_mgr_code", "credit_mgr_code");
                            bulkCopy.ColumnMappings.Add("N° de Piéce", "N° de Piéce");
                            bulkCopy.ColumnMappings.Add("Date Piéce", "Date Piéce");
                            bulkCopy.ColumnMappings.Add("Date_Echéance", "Date_Echéance");
                            bulkCopy.ColumnMappings.Add("Votre N° de commande", "Votre N° de commande");
                            bulkCopy.ColumnMappings.Add("Montant TTC", "Montant TTC");
                            bulkCopy.ColumnMappings.Add("SourceCd", "SourceCd");
                            bulkCopy.ColumnMappings.Add("TotalBalanceAmt", "TotalBalanceAmt");
                            bulkCopy.ColumnMappings.Add("credit_limit", "credit_limit");
                            bulkCopy.WriteToServer(reader);
                        }

                        using (var countCommand = new SqlCommand(
                            "SELECT COUNT(*) FROM dbo.T_Credit_Recouvrement_AR_OPEN;",
                            targetConnection,
                            transaction))
                            importedRows = Convert.ToInt32(countCommand.ExecuteScalar());

                        ExecuteArOpenTargetSql(targetConnection, transaction, @"
INSERT INTO dbo.T_Credit_Recouvrement_CONTACT_FICHE
(
    br_cust_nbr,cust_name,mgr_code,total_balance_amt,credit_limit
)
SELECT DISTINCT
       source.brcustnbr,
       source.cust_name,
       source.credit_mgr_code,
       source.TotalBalanceAmt,
       source.credit_limit
FROM dbo.T_Credit_Recouvrement_AR_OPEN AS source
LEFT JOIN dbo.T_Credit_Recouvrement_CONTACT_FICHE AS target
    ON target.br_cust_nbr=source.brcustnbr
WHERE target.br_cust_nbr IS NULL;");

                        string managerInClause =
                            string.Join(",", parameterNames);
                        using (var insertWithoutInvoice = new SqlCommand(@"
INSERT INTO dbo.T_Credit_Recouvrement_CONTACT_FICHE
(
    br_cust_nbr,cust_name,mgr_code,total_balance_amt,credit_limit
)
SELECT customer.branch_customer_nbr,
       customer.cust_name,
       customer.credit_mgr_code,
       customer.TotalBalanceAmt,
       customer.credit_limit
FROM DSS_COPIE.dbo.customer AS customer
INNER JOIN DSS_COPIE.dbo.customer_location AS customer_location
    ON customer_location.branch_customer_nbr=customer.branch_customer_nbr
LEFT JOIN dbo.T_Credit_Recouvrement_CONTACT_FICHE AS contact
    ON contact.br_cust_nbr=customer.branch_customer_nbr
WHERE customer.CompanyCd='FR'
  AND LEFT(customer.branch_nbr,2) IN ('15','21','24')
  AND customer.credit_mgr_code IN (" + managerInClause + @")
  AND customer_location.suffix='000'
  AND contact.br_cust_nbr IS NULL;", targetConnection, transaction))
                        {
                            insertWithoutInvoice.CommandTimeout = 300;
                            for (int index = 0; index < managerCodes.Count; index++)
                                insertWithoutInvoice.Parameters.Add(
                                    parameterNames[index],
                                    SqlDbType.VarChar,
                                    20).Value = managerCodes[index];
                            insertWithoutInvoice.ExecuteNonQuery();
                        }

                        ExecuteArOpenTargetSql(targetConnection, transaction, @"
UPDATE contact
SET total_balance_amt=customer.TotalBalanceAmt,
    mgr_code=customer.credit_mgr_code,
    credit_limit=customer.credit_limit
FROM dbo.T_Credit_Recouvrement_CONTACT_FICHE AS contact
INNER JOIN DSS_COPIE.dbo.customer AS customer
    ON contact.br_cust_nbr=customer.branch_customer_nbr;");

                        ExecuteArOpenTargetSql(targetConnection, transaction, @"
UPDATE dbo.T_Credit_Recouvrement_CONTACT_FICHE
SET red_alert=0,
    relance_preventive=0,
    relance_active_1=0,
    relance_active_2=0,
    relance_automatch=0,
    relance_preventive_marketplace=0,
    relance_active_1_marketplace=0,
    relance_active_2_marketplace=0;");

                        transaction.Commit();
                        return importedRows;
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw;
                    }
                }
            }
        }

        private static void ExecuteArOpenTargetSql(
            SqlConnection connection,
            SqlTransaction transaction,
            string sql)
        {
            using (var command = new SqlCommand(sql, connection, transaction))
            {
                command.CommandTimeout = 300;
                command.ExecuteNonQuery();
            }
        }

        public void Envoi_Echeancier_Auto_Client(
            string sql_con,
            string logs,
            string tmp_folder,
            string session_name)
        {
            const string executionParameter =
                "date_dernier_envoi_echeancier_auto_client";

            string root = GetServicePath();
            logsFolder = Path.Combine(root, logs ?? "");
            tempFolder = Path.Combine(root, tmp_folder ?? "");
            sessionName = session_name ?? "";
            Directory.CreateDirectory(logsFolder);
            Directory.CreateDirectory(tempFolder);

            JsonFile configuration = JsonConvert.DeserializeObject<JsonFile>(
                GetImcaParameter(sql_con, global_application_name) ?? "");

            if (configuration?.countries == null ||
                configuration.countries.Count == 0)
            {
                throw new InvalidOperationException(
                    global_application_name +
                    " parameters are empty or invalid");
            }

            foreach (Country item in configuration.countries)
            {
                ApplyCountryConfiguration(item, sql_con);

                if (!IsTrue(active) ||
                    !country.Equals("FR", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                try
                {
                    ValidateAutomaticStatementConfiguration();

                    DateTime lastExecution =
                        GetOptionalExecutionParameterDate(executionParameter);

                    WriteLog(
                        "   Automatic customer statement schedule" +
                        " - Last successful execution : " +
                        lastExecution.ToString("dd/MM/yyyy HH:mm:ss") +
                        " - Current date : " +
                        DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));

                    if (lastExecution.Date >= DateTime.Today)
                    {
                        WriteLog(
                            "   Automatic customer statement already sent today");
                        continue;
                    }

                    graphService = ConnectGraph();
                    int sent = 0;
                    int skipped = 0;

                    foreach (AutomaticStatementCustomer customer
                        in automaticStatementCustomers)
                    {
                        string customerCode =
                            (customer?.code_client ?? "").Trim();
                        string recipients =
                            (customer?.email_destinataire ?? "").Trim();

                        if (string.IsNullOrWhiteSpace(customerCode))
                        {
                            throw new InvalidDataException(
                                "An echeancier_auto_clients entry has no code_client");
                        }

                        if (BuildRecipients(recipients).Count == 0)
                        {
                            throw new InvalidDataException(
                                "Invalid automatic statement recipient" +
                                " - Customer : " + customerCode +
                                " - Value : " + recipients);
                        }

                        DataTable details = GetAutomaticStatementRows(customerCode);
                        if (details.Rows.Count == 0)
                        {
                            skipped++;
                            WriteLog(
                                "       Automatic customer statement skipped" +
                                " - No AR Open row" +
                                " - Customer : " + customerCode +
                                " - Recipient(s) : " + recipients);
                            continue;
                        }

                        string filePath = Path.Combine(
                            tempFolder,
                            "Detail_Compte_" +
                            CleanFileName(customerCode) +
                            "_" +
                            DateTime.Now.ToString(
                                "yyyyMMdd_HHmmss",
                                CultureInfo.InvariantCulture) +
                            ".xlsx");

                        try
                        {
                            CreateAutomaticStatementWorkbook(
                                details,
                                filePath);

                            SendGraphMailWithAttachments(
                                fr_facturation_graph_send_as,
                                recipients,
                                "",
                                "",
                                "Etat de compte",
                                "Bonjour,<br/><br/>" +
                                "Veuillez trouver ci-joint votre état de compte." +
                                "<br/><br/>Cordialement,<br/>" +
                                "Ingram Micro - Service Analyse Crédit",
                                new List<string> { filePath });

                            sent++;
                            WriteLog(
                                "       Automatic customer statement sent" +
                                " - Customer : " + customerCode +
                                " - Recipient(s) : " + recipients +
                                " - Row(s) : " + details.Rows.Count +
                                " - File : " + filePath);
                        }
                        finally
                        {
                            try
                            {
                                if (File.Exists(filePath))
                                    File.Delete(filePath);
                            }
                            catch (Exception cleanupException)
                            {
                                WriteLog(
                                    "       Automatic statement temporary file " +
                                    "cleanup error" +
                                    " - File : " + filePath +
                                    " - Error : " +
                                    cleanupException.Message);
                            }
                        }
                    }

                    DateTime completedAt = DateTime.Now;
                    SetCreditReviewScheduleDate(
                        executionParameter,
                        completedAt);

                    WriteLog(
                        "   Automatic customer statement processing completed" +
                        " - Configured customer(s) : " +
                        automaticStatementCustomers.Count +
                        " - Email(s) sent : " + sent +
                        " - Customer(s) without AR Open row : " + skipped +
                        " - Successful execution : " +
                        completedAt.ToString("dd/MM/yyyy HH:mm:ss"));
                }
                catch (Exception ex)
                {
                    string details = GetDetailedExceptionMessage(ex);
                    WriteLog(
                        "   Automatic customer statement error" +
                        " - Execution parameter not updated" +
                        " - Details : " + details);

                    try
                    {
                        if (graphService == null)
                            graphService = ConnectGraph();

                        SendTechnicalAlert(
                            nameof(Envoi_Echeancier_Auto_Client),
                            details,
                            "AUTOMATIC CUSTOMER STATEMENT");
                    }
                    catch (Exception alertException)
                    {
                        WriteLog(
                            "   Automatic customer statement technical alert " +
                            "could not be sent" +
                            " - Details : " +
                            GetDetailedExceptionMessage(alertException));
                    }
                }
                finally
                {
                    graphService = null;
                }
            }
        }

        private void ValidateAutomaticStatementConfiguration()
        {
            if (string.IsNullOrWhiteSpace(sql_connexion))
            {
                throw new InvalidOperationException(
                    "sql_connexion is empty for automatic customer statements");
            }

            if (string.IsNullOrWhiteSpace(sql_creation_compte))
            {
                throw new InvalidOperationException(
                    "sql_creation_compte is empty for automatic customer " +
                    "statement scheduling");
            }

            if (string.IsNullOrWhiteSpace(fr_facturation_graph_send_as))
            {
                throw new InvalidOperationException(
                    "fr_facturation_graph_send_as is empty for automatic " +
                    "customer statements");
            }

            if (automaticStatementCustomers == null ||
                automaticStatementCustomers.Count == 0)
            {
                throw new InvalidOperationException(
                    "echeancier_auto_clients is empty");
            }
        }

        private DataTable GetAutomaticStatementRows(string customerCode)
        {
            return FillDataTable(
                sql_connexion,
                @"SELECT brcustnbr AS [Code Client],
                         cust_name AS [Nom Client],
                         [N° de Piéce],
                         [Votre N° de commande],
                         [Date Piéce],
                         [Montant TTC],
                         Date_Echéance AS [Date Echéance]
                  FROM dbo.T_Credit_Recouvrement_AR_OPEN
                  WHERE brcustnbr=@CUSTOMER
                  ORDER BY Date_Echéance,[Date Piéce];",
                new SqlParameter(
                    "@CUSTOMER",
                    SqlDbType.VarChar,
                    50)
                {
                    Value = customerCode ?? ""
                });
        }

        private static void CreateAutomaticStatementWorkbook(
            DataTable details,
            string outputPath)
        {
            string directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrWhiteSpace(directory))
                Directory.CreateDirectory(directory);

            using (var workbook = new XLWorkbook())
            {
                IXLWorksheet worksheet = workbook.AddWorksheet("echeancier");
                worksheet.Cell(1, 1).InsertTable(
                    details,
                    "Echeancier",
                    true);
                worksheet.SheetView.FreezeRows(1);
                worksheet.Row(1).Style.Font.Bold = true;
                worksheet.Row(1).Style.Fill.BackgroundColor =
                    XLColor.DarkBlue;
                worksheet.Row(1).Style.Font.FontColor = XLColor.White;
                worksheet.Columns().AdjustToContents(1, 80);

                if (details.Columns.Contains("Montant TTC"))
                {
                    int amountColumn =
                        details.Columns["Montant TTC"].Ordinal + 1;
                    worksheet.Column(amountColumn)
                        .Style.NumberFormat.Format = "#,##0.00";
                }

                workbook.SaveAs(outputPath);
            }
        }

        public void Extract_And_Send_Traites(
            string sql_con,
            string logs,
            string tmp_folder,
            string session_name)
        {
            const string extractionParameter =
                "date_dernier_extract_data_traites_ar_open";
            const string sendingParameter =
                "date_dernier_envoi_traites";

            string root = GetServicePath();
            logsFolder = Path.Combine(root, logs ?? "");
            tempFolder = Path.Combine(root, tmp_folder ?? "");
            sessionName = session_name ?? "";
            Directory.CreateDirectory(logsFolder);
            Directory.CreateDirectory(tempFolder);

            JsonFile configuration = JsonConvert.DeserializeObject<JsonFile>(
                GetImcaParameter(sql_con, global_application_name) ?? "");

            if (configuration?.countries == null ||
                configuration.countries.Count == 0)
            {
                throw new InvalidOperationException(
                    global_application_name +
                    " parameters are empty or invalid");
            }

            foreach (Country item in configuration.countries)
            {
                ApplyCountryConfiguration(item, sql_con);

                if (!IsTrue(active) ||
                    !country.Equals("FR", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                try
                {
                    ValidateTraitesConfiguration();

                    DateTime extractionDate =
                        GetOptionalExecutionParameterDate(extractionParameter);

                    if (extractionDate.Date < DateTime.Today)
                    {
                        DateTime? arOpenUpdatedAt;
                        string sourceDetails;

                        if (!IsArOpenUpdatedToday(
                                out arOpenUpdatedAt,
                                out sourceDetails))
                        {
                            WriteLog(
                                "   Traites AR Open extraction skipped" +
                                " - Source ar_open is not updated today" +
                                " - Details : " + sourceDetails);
                            continue;
                        }

                        WriteLog(
                            "   Starting Traites AR Open extraction" +
                            " - Source : " + sourceDetails);

                        int extractedRows;
                        int insertedRows = ProcessTraitesArOpenExtraction(
                            out extractedRows);

                        extractionDate = DateTime.Now;
                        SetCreditReviewScheduleDate(
                            extractionParameter,
                            extractionDate);

                        WriteLog(
                            "   Traites AR Open extraction completed" +
                            " - Temporary row(s) : " + extractedRows +
                            " - New row(s) queued : " + insertedRows +
                            " - Successful execution : " +
                            extractionDate.ToString("dd/MM/yyyy HH:mm:ss"));
                    }
                    else
                    {
                        WriteLog(
                            "   Traites AR Open extraction already completed today" +
                            " - Last successful execution : " +
                            extractionDate.ToString("dd/MM/yyyy HH:mm:ss"));
                    }

                    // Sending is only allowed after a successful extraction today.
                    extractionDate =
                        GetOptionalExecutionParameterDate(extractionParameter);
                    if (extractionDate.Date < DateTime.Today)
                    {
                        WriteLog(
                            "   Traites sending skipped" +
                            " - Today's extraction has not completed successfully");
                        continue;
                    }

                    DateTime sendingDate =
                        GetOptionalExecutionParameterDate(sendingParameter);
                    if (sendingDate.Date >= DateTime.Today)
                    {
                        WriteLog(
                            "   Traites sending already completed today" +
                            " - Last successful execution : " +
                            sendingDate.ToString("dd/MM/yyyy HH:mm:ss"));
                        continue;
                    }

                    WriteLog("   Starting automatic Traites sending");
                    graphService = ConnectGraph();

                    int sentCustomers;
                    int sentRows;
                    int removedWithoutContact;
                    ProcessTraitesSending(
                        out sentCustomers,
                        out sentRows,
                        out removedWithoutContact);

                    sendingDate = DateTime.Now;
                    SetCreditReviewScheduleDate(
                        sendingParameter,
                        sendingDate);

                    WriteLog(
                        "   Automatic Traites sending completed" +
                        " - Customer email(s) sent : " + sentCustomers +
                        " - Row(s) marked sent : " + sentRows +
                        " - Pending row(s) removed without contact : " +
                        removedWithoutContact +
                        " - Successful execution : " +
                        sendingDate.ToString("dd/MM/yyyy HH:mm:ss"));
                }
                catch (Exception ex)
                {
                    string details = GetDetailedExceptionMessage(ex);
                    WriteLog(
                        "   Traites extraction/sending error" +
                        " - Daily parameter of the failing phase not updated" +
                        " - Details : " + details);

                    try
                    {
                        if (graphService == null)
                            graphService = ConnectGraph();

                        SendTechnicalAlert(
                            nameof(Extract_And_Send_Traites),
                            details,
                            "TRAITES AR OPEN");
                    }
                    catch (Exception alertException)
                    {
                        WriteLog(
                            "   Traites technical alert could not be sent" +
                            " - Details : " +
                            GetDetailedExceptionMessage(alertException));
                    }
                }
                finally
                {
                    graphService = null;
                }
            }
        }

        private void ValidateTraitesConfiguration()
        {
            if (string.IsNullOrWhiteSpace(sql_connexion))
                throw new InvalidOperationException(
                    "sql_connexion is empty for Traites processing");
            if (string.IsNullOrWhiteSpace(sql_dss_copie))
                throw new InvalidOperationException(
                    "sql_dss_copie is empty for Traites extraction");
            if (string.IsNullOrWhiteSpace(sql_creation_compte))
                throw new InvalidOperationException(
                    "sql_creation_compte is empty for Traites scheduling");
            if (string.IsNullOrWhiteSpace(fr_lcrna_graph_send_as))
                throw new InvalidOperationException(
                    "fr_lcrna_graph_send_as is empty for Traites sending");
        }

        private int ProcessTraitesArOpenExtraction(out int extractedRows)
        {
            const string sourceSql = @"
SELECT DISTINCT
       ar_open.brcustnbr,
       customer.cust_name,
       customer.credit_mgr_code,
       CASE
           WHEN ar_open.PaySeqNbr IS NULL THEN ar_open.BrInvoiceNbr
           ELSE ar_open.BrInvoiceNbr + RIGHT(ar_open.PaySeqNbr,2)
       END AS [N° de Piéce],
       CAST(ar_open.InvoiceDt AS smalldatetime) AS [Date Piéce],
       CAST(ar_open.DueDt AS smalldatetime) AS Date_Echéance,
       ar_open.CustOrderNbr AS [Votre N° de commande],
       CASE
           WHEN ar_open.CurrencyCd='USD' THEN ar_open.ForeignSalesAmt
           ELSE ar_open.TotalSalesAmt
       END AS [Montant TTC],
       ar_open.SourceCd
FROM dbo.ar_open AS ar_open
INNER JOIN dbo.customer AS customer
    ON ar_open.BrCustNbr=customer.branch_customer_nbr
WHERE ar_open.BrCustNbr IN
(
    SELECT branch_customer_nbr
    FROM dbo.customer
    WHERE branch_nbr='21'
      AND Apply_Cash_Flg='D'
      AND StatusCustFlg<>'D'
)
AND ar_open.CheckApplyCd='T';";

            using (var sourceConnection = new SqlConnection(sql_dss_copie))
            using (var sourceCommand = new SqlCommand(sourceSql, sourceConnection))
            using (var targetConnection = new SqlConnection(sql_connexion))
            {
                sourceCommand.CommandTimeout = 0;
                sourceConnection.Open();
                targetConnection.Open();

                using (SqlTransaction transaction =
                    targetConnection.BeginTransaction())
                {
                    try
                    {
                        using (var truncate = new SqlCommand(
                            "TRUNCATE TABLE dbo.T_Credit_Traites_AR_OPEN_TMP;",
                            targetConnection,
                            transaction))
                        {
                            truncate.CommandTimeout = 300;
                            truncate.ExecuteNonQuery();
                        }

                        using (SqlDataReader reader = sourceCommand.ExecuteReader())
                        using (var bulkCopy = new SqlBulkCopy(
                            targetConnection,
                            SqlBulkCopyOptions.TableLock,
                            transaction))
                        {
                            bulkCopy.DestinationTableName =
                                "dbo.T_Credit_Traites_AR_OPEN_TMP";
                            bulkCopy.BulkCopyTimeout = 0;
                            bulkCopy.ColumnMappings.Add("brcustnbr", "brcustnbr");
                            bulkCopy.ColumnMappings.Add("cust_name", "cust_name");
                            bulkCopy.ColumnMappings.Add(
                                "credit_mgr_code",
                                "credit_mgr_code");
                            bulkCopy.ColumnMappings.Add(
                                "N° de Piéce",
                                "N° de Piéce");
                            bulkCopy.ColumnMappings.Add(
                                "Date Piéce",
                                "Date Piéce");
                            bulkCopy.ColumnMappings.Add(
                                "Date_Echéance",
                                "Date_Echéance");
                            bulkCopy.ColumnMappings.Add(
                                "Votre N° de commande",
                                "Votre N° de commande");
                            bulkCopy.ColumnMappings.Add(
                                "Montant TTC",
                                "Montant TTC");
                            bulkCopy.ColumnMappings.Add("SourceCd", "SourceCd");
                            bulkCopy.WriteToServer(reader);
                        }

                        using (var count = new SqlCommand(
                            "SELECT COUNT(*) FROM " +
                            "dbo.T_Credit_Traites_AR_OPEN_TMP;",
                            targetConnection,
                            transaction))
                        {
                            extractedRows = Convert.ToInt32(
                                count.ExecuteScalar());
                        }

                        int insertedRows;
                        using (var insert = new SqlCommand(@"
INSERT INTO dbo.T_Credit_Traites_AR_OPEN
(
    brcustnbr,
    cust_name,
    credit_mgr_code,
    [N° de Piéce],
    [Date Piéce],
    Date_Echéance,
    [Votre N° de commande],
    [Montant TTC],
    SourceCd,
    top_traite
)
SELECT source.brcustnbr,
       source.cust_name,
       source.credit_mgr_code,
       source.[N° de Piéce],
       source.[Date Piéce],
       source.Date_Echéance,
       source.[Votre N° de commande],
       source.[Montant TTC],
       source.SourceCd,
       'N'
FROM dbo.T_Credit_Traites_AR_OPEN_TMP AS source
LEFT JOIN dbo.T_Credit_Traites_AR_OPEN AS target
    ON target.brcustnbr=source.brcustnbr
   AND target.[N° de Piéce]=source.[N° de Piéce]
   AND target.Date_Echéance=source.Date_Echéance
WHERE target.brcustnbr IS NULL;",
                            targetConnection,
                            transaction))
                        {
                            insert.CommandTimeout = 300;
                            insertedRows = insert.ExecuteNonQuery();
                        }

                        transaction.Commit();
                        return insertedRows;
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw;
                    }
                }
            }
        }

        private void ProcessTraitesSending(
            out int sentCustomers,
            out int sentRows,
            out int removedWithoutContact)
        {
            sentCustomers = 0;
            sentRows = 0;
            removedWithoutContact = 0;

            DataTable customers = FillDataTable(
                sql_connexion,
                @"SELECT DISTINCT brcustnbr
                  FROM dbo.T_Credit_Traites_AR_OPEN
                  WHERE top_traite='N'
                  ORDER BY brcustnbr;");

            const string body =
                "Bonjour Madame, Monsieur,<br/><br/>" +
                "Vous avez opté pour le paiement par lcr automatique.<br/>" +
                "Afin de vous aider au mieux dans la gestion de votre " +
                "trésorerie, nous vous adressons ci-joint le détail des " +
                "prélèvements à venir.<br/>" +
                "Si il vous manque des factures ou avoirs, vous pouvez vous " +
                "connecter sur votre compte Ingram micro Rubrique " +
                "<a href='https://fr-new.ingrammicro.com/Site/InvoiceList'>" +
                "Mon compte/ Mes factures</a> ou nous contacter sur " +
                "l’adresse mail <a href='mailto:duplicatas-France@ingrammicro.com'>" +
                "duplicatas-France@ingrammicro.com</a><br/>" +
                "Si vous souhaitez l’arrêt de ce service n’hésitez pas à " +
                "contacter votre gestionnaire de crédit.<br/><br/>" +
                "Nous restons bien évidemment à votre disposition,<br/>" +
                "Cordialement<br/>L’équipe Finance Ingram Micro";

            foreach (DataRow customerRow in customers.Rows)
            {
                string customerCode =
                    Convert.ToString(customerRow["brcustnbr"]).Trim();

                string recipients = GetTraitesRecipients(customerCode);
                if (BuildRecipients(recipients).Count == 0)
                {
                    WriteLog(
                        "       Traites customer has no recovery contact" +
                        " - Customer : " + customerCode);
                    continue;
                }

                DataTable details = FillDataTable(
                    sql_connexion,
                    @"SELECT brcustnbr AS [Code Client],
                             cust_name AS [Nom Client],
                             [N° de Piéce],
                             [Votre N° de commande],
                             [Date Piéce],
                             [Montant TTC],
                             Date_Echéance AS [Date Echéance]
                      FROM dbo.T_Credit_Traites_AR_OPEN
                      WHERE brcustnbr=@CUSTOMER
                        AND top_traite='N'
                      ORDER BY Date_Echéance,[Date Piéce];",
                    new SqlParameter(
                        "@CUSTOMER",
                        SqlDbType.VarChar,
                        50)
                    {
                        Value = customerCode
                    });

                if (details.Rows.Count == 0)
                    continue;

                string filePath = Path.Combine(
                    tempFolder,
                    "Traites_Client_" +
                    CleanFileName(customerCode) +
                    "_" +
                    DateTime.Now.ToString(
                        "yyyyMMdd_HHmmss",
                        CultureInfo.InvariantCulture) +
                    ".xlsx");

                try
                {
                    CreateTraitesWorkbook(details, filePath);

                    SendGraphMailWithAttachments(
                        fr_lcrna_graph_send_as,
                        recipients,
                        "",
                        "",
                        "Le détail de votre prélèvement automatique (" +
                        customerCode + ")",
                        body,
                        new List<string> { filePath });

                    int affected;
                    using (var connection = new SqlConnection(sql_connexion))
                    using (var command = new SqlCommand(@"
UPDATE dbo.T_Credit_Traites_AR_OPEN
SET top_traite='O',
    envoye_le=GETDATE()
WHERE brcustnbr=@CUSTOMER
  AND top_traite='N';", connection))
                    {
                        command.CommandTimeout = 300;
                        command.Parameters.Add(
                            "@CUSTOMER",
                            SqlDbType.VarChar,
                            50).Value = customerCode;
                        connection.Open();
                        affected = command.ExecuteNonQuery();
                    }

                    sentCustomers++;
                    sentRows += affected;
                    WriteLog(
                        "       Traites customer email sent" +
                        " - Customer : " + customerCode +
                        " - Recipient(s) : " + recipients +
                        " - Row(s) : " + affected);
                }
                finally
                {
                    try
                    {
                        if (File.Exists(filePath))
                            File.Delete(filePath);
                    }
                    catch (Exception cleanupException)
                    {
                        WriteLog(
                            "       Traites temporary workbook cleanup error" +
                            " - File : " + filePath +
                            " - Error : " + cleanupException.Message);
                    }
                }
            }

            // Historical behavior: rows still pending after the loop correspond
            // to customers without a configured recovery contact and are removed.
            using (var connection = new SqlConnection(sql_connexion))
            using (var command = new SqlCommand(
                "DELETE FROM dbo.T_Credit_Traites_AR_OPEN " +
                "WHERE top_traite='N';",
                connection))
            {
                command.CommandTimeout = 300;
                connection.Open();
                removedWithoutContact = command.ExecuteNonQuery();
            }
        }

        private string GetTraitesRecipients(string customerCode)
        {
            DataTable contacts = FillDataTable(
                sql_connexion,
                @"SELECT DISTINCT Email
                  FROM dbo.T_Credit_Review_CONTACT_FICHE
                  WHERE code_client=@CUSTOMER
                    AND id_fonction IN (6,7)
                    AND NULLIF(LTRIM(RTRIM(Email)),'') IS NOT NULL;",
                new SqlParameter(
                    "@CUSTOMER",
                    SqlDbType.VarChar,
                    50)
                {
                    Value = customerCode ?? ""
                });

            return string.Join(
                ";",
                contacts.AsEnumerable()
                    .Select(row => Convert.ToString(row["Email"]).Trim())
                    .Where(IsEmail)
                    .Distinct(StringComparer.OrdinalIgnoreCase));
        }

        private static void CreateTraitesWorkbook(
            DataTable details,
            string outputPath)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

            using (var workbook = new XLWorkbook())
            {
                IXLWorksheet worksheet = workbook.AddWorksheet("traites");
                worksheet.Cell(1, 1).InsertTable(details, "Traites", true);
                worksheet.SheetView.FreezeRows(1);
                worksheet.Row(1).Style.Font.Bold = true;
                worksheet.Row(1).Style.Fill.BackgroundColor = XLColor.DarkBlue;
                worksheet.Row(1).Style.Font.FontColor = XLColor.White;
                worksheet.Columns().AdjustToContents(1, 80);

                if (details.Columns.Contains("Montant TTC"))
                {
                    int amountColumn =
                        details.Columns["Montant TTC"].Ordinal + 1;
                    worksheet.Column(amountColumn)
                        .Style.NumberFormat.Format = "#,##0.00";
                }

                workbook.SaveAs(outputPath);
            }
        }

        public void Planification_Auto_Calcul_Encours(
            string sql_con,
            string logs,
            string tmp_folder,
            string session_name)
        {
            const string executionParameter =
                "date_planification_auto_calcul_encours";

            string root = GetServicePath();
            logsFolder = Path.Combine(root, logs ?? "");
            tempFolder = Path.Combine(root, tmp_folder ?? "");
            sessionName = session_name ?? "";
            Directory.CreateDirectory(logsFolder);
            Directory.CreateDirectory(tempFolder);

            JsonFile configuration = JsonConvert.DeserializeObject<JsonFile>(
                GetImcaParameter(sql_con, global_application_name) ?? "");

            if (configuration?.countries == null ||
                configuration.countries.Count == 0)
            {
                throw new InvalidOperationException(
                    global_application_name +
                    " parameters are empty or invalid");
            }

            foreach (Country item in configuration.countries)
            {
                ApplyCountryConfiguration(item, sql_con);

                if (!IsTrue(active) ||
                    !country.Equals(
                        "FR",
                        StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                try
                {
                    ValidateAutomaticOutstandingCalculationConfiguration();

                    DateTime lastExecution =
                        GetOptionalExecutionParameterDate(executionParameter);

                    WriteLog(
                        "   Automatic outstanding calculation schedule" +
                        " - Last successful execution : " +
                        lastExecution.ToString("dd/MM/yyyy HH:mm:ss") +
                        " - Current date : " +
                        DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));

                    if (lastExecution.Date >= DateTime.Today)
                    {
                        WriteLog(
                            "   Automatic outstanding calculation " +
                            "already completed today");
                        continue;
                    }

                    WriteLog(
                        "   Starting automatic outstanding calculation planning");

                    int copiedClients;
                    int createdRequests =
                        ProcessAutomaticOutstandingCalculationPlanning(
                            out copiedClients);

                    DateTime completedAt = DateTime.Now;
                    SetCreditReviewScheduleDate(
                        executionParameter,
                        completedAt);

                    WriteLog(
                        "   Automatic outstanding calculation planning completed" +
                        " - New request(s) : " + createdRequests +
                        " - Client row(s) copied : " + copiedClients +
                        " - Successful execution : " +
                        completedAt.ToString("dd/MM/yyyy HH:mm:ss"));
                }
                catch (Exception ex)
                {
                    string details = GetDetailedExceptionMessage(ex);

                    WriteLog(
                        "   Automatic outstanding calculation planning error" +
                        " - Execution parameter not updated" +
                        " - Details : " + details);

                    try
                    {
                        if (graphService == null)
                        {
                            WriteLog(
                                "   Connecting to Microsoft Graph for automatic " +
                                "outstanding calculation technical alert");
                            graphService = ConnectGraph();
                        }

                        SendTechnicalAlert(
                            nameof(Planification_Auto_Calcul_Encours),
                            details,
                            "AUTOMATIC OUTSTANDING CALCULATION PLANNING");

                        WriteLog(
                            "   Automatic outstanding calculation " +
                            "technical alert sent");
                    }
                    catch (Exception alertException)
                    {
                        WriteLog(
                            "   Automatic outstanding calculation technical " +
                            "alert could not be sent" +
                            " - Details : " +
                            GetDetailedExceptionMessage(alertException));
                    }
                }
                finally
                {
                    graphService = null;
                }
            }
        }

        private void ValidateAutomaticOutstandingCalculationConfiguration()
        {
            if (string.IsNullOrWhiteSpace(sql_connexion))
            {
                throw new InvalidOperationException(
                    "sql_connexion is empty for automatic outstanding " +
                    "calculation planning");
            }

            if (string.IsNullOrWhiteSpace(sql_creation_compte))
            {
                throw new InvalidOperationException(
                    "sql_creation_compte is empty for automatic outstanding " +
                    "calculation scheduling");
            }
        }

        private int ProcessAutomaticOutstandingCalculationPlanning(
            out int copiedClients)
        {
            copiedClients = 0;
            int createdRequests = 0;

            using (var connection = new SqlConnection(sql_connexion))
            {
                connection.Open();

                using (SqlTransaction transaction =
                    connection.BeginTransaction(IsolationLevel.Serializable))
                {
                    try
                    {
                        // The application lock prevents two IMCA executions from
                        // duplicating the same automatic requests concurrently.
                        using (var lockCommand = new SqlCommand(
                            @"DECLARE @result int;
                              EXEC @result = sys.sp_getapplock
                                  @Resource=N'CREDIT_PLANIFICATION_AUTO_CALCUL_ENCOURS',
                                  @LockMode=N'Exclusive',
                                  @LockOwner=N'Transaction',
                                  @LockTimeout=0;
                              SELECT @result;",
                            connection,
                            transaction))
                        {
                            lockCommand.CommandTimeout = 30;
                            int lockResult =
                                Convert.ToInt32(lockCommand.ExecuteScalar());

                            if (lockResult < 0)
                            {
                                throw new InvalidOperationException(
                                    "Automatic outstanding calculation planning " +
                                    "is already running" +
                                    " - SQL application lock result : " +
                                    lockResult);
                            }
                        }

                        var sourceRequestIds = new List<int>();

                        using (var selectCommand = new SqlCommand(
                            @"SELECT id_demande
                              FROM dbo.T_Credit_Calcul_encours_demandes
                              WHERE automatique=1
                              ORDER BY id_demande;",
                            connection,
                            transaction))
                        {
                            selectCommand.CommandTimeout = 300;

                            using (SqlDataReader reader =
                                selectCommand.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    sourceRequestIds.Add(
                                        Convert.ToInt32(reader["id_demande"]));
                                }
                            }
                        }

                        foreach (int sourceRequestId in sourceRequestIds)
                        {
                            int newRequestId;

                            using (var insertRequest = new SqlCommand(
                                @"INSERT INTO dbo.T_Credit_Calcul_encours_demandes
                                  (
                                      qui,
                                      quand,
                                      id_statut,
                                      commentaires,
                                      pourcentage_risk,
                                      open_order,
                                      cto_valides,
                                      reglement_en_cours,
                                      commandes_a_valider,
                                      commandes_non_valorisees,
                                      financement,
                                      automatique,
                                      factor
                                  )
                                  SELECT qui,
                                         GETDATE(),
                                         0,
                                         commentaires,
                                         pourcentage_risk,
                                         0,
                                         0,
                                         0,
                                         0,
                                         0,
                                         financement,
                                         1,
                                         factor
                                  FROM dbo.T_Credit_Calcul_encours_demandes
                                  WHERE id_demande=@SOURCE_ID
                                    AND automatique=1;

                                  SELECT CAST(SCOPE_IDENTITY() AS int);",
                                connection,
                                transaction))
                            {
                                insertRequest.CommandTimeout = 300;
                                insertRequest.Parameters.Add(
                                    "@SOURCE_ID",
                                    SqlDbType.Int).Value = sourceRequestId;

                                object result = insertRequest.ExecuteScalar();
                                if (result == null || result == DBNull.Value)
                                {
                                    throw new InvalidOperationException(
                                        "Automatic outstanding calculation source " +
                                        "request is no longer active" +
                                        " - Source request ID : " +
                                        sourceRequestId);
                                }

                                newRequestId = Convert.ToInt32(result);
                            }

                            using (var copyClients = new SqlCommand(
                                @"INSERT INTO dbo.T_Credit_Calcul_encours_clients
                                  (
                                      id_demande,
                                      code_client,
                                      nom_client,
                                      id_statut,
                                      terms,
                                      terms_descr,
                                      risk_class,
                                      tax_exempt_nbr,
                                      taxe_code,
                                      credit_limit_impulse,
                                      devise,
                                      taux_dollar
                                  )
                                  SELECT @NEW_REQUEST_ID,
                                         code_client,
                                         nom_client,
                                         id_statut,
                                         terms,
                                         terms_descr,
                                         risk_class,
                                         tax_exempt_nbr,
                                         taxe_code,
                                         credit_limit_impulse,
                                         devise,
                                         taux_dollar
                                  FROM dbo.T_Credit_Calcul_encours_clients
                                  WHERE id_demande=@SOURCE_ID;",
                                connection,
                                transaction))
                            {
                                copyClients.CommandTimeout = 300;
                                copyClients.Parameters.Add(
                                    "@NEW_REQUEST_ID",
                                    SqlDbType.Int).Value = newRequestId;
                                copyClients.Parameters.Add(
                                    "@SOURCE_ID",
                                    SqlDbType.Int).Value = sourceRequestId;
                                copiedClients += copyClients.ExecuteNonQuery();
                            }

                            using (var finalizeRequests = new SqlCommand(
                                @"UPDATE dbo.T_Credit_Calcul_encours_demandes
                                  SET automatique=0
                                  WHERE id_demande=@SOURCE_ID;

                                  UPDATE dbo.T_Credit_Calcul_encours_demandes
                                  SET id_statut=1
                                  WHERE id_demande=@NEW_REQUEST_ID;",
                                connection,
                                transaction))
                            {
                                finalizeRequests.CommandTimeout = 300;
                                finalizeRequests.Parameters.Add(
                                    "@SOURCE_ID",
                                    SqlDbType.Int).Value = sourceRequestId;
                                finalizeRequests.Parameters.Add(
                                    "@NEW_REQUEST_ID",
                                    SqlDbType.Int).Value = newRequestId;
                                finalizeRequests.ExecuteNonQuery();
                            }

                            createdRequests++;

                            WriteLog(
                                "       Automatic outstanding calculation " +
                                "request duplicated" +
                                " - Source request ID : " + sourceRequestId +
                                " - New request ID : " + newRequestId);
                        }

                        transaction.Commit();
                        return createdRequests;
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw;
                    }
                }
            }
        }

        public void Credit_Review_Mails(
            string sql_con,
            string logs,
            string tmp_folder,
            string session_name)
        {
            const string scheduleParameter =
                "date_derniere_recherche_credit_review";

            string root = GetServicePath();
            logsFolder = Path.Combine(root, logs ?? "");
            tempFolder = Path.Combine(root, tmp_folder ?? "");
            sessionName = session_name ?? "";
            Directory.CreateDirectory(logsFolder);
            Directory.CreateDirectory(tempFolder);

            JsonFile configuration =
                JsonConvert.DeserializeObject<JsonFile>(
                    GetImcaParameter(
                        sql_con,
                        global_application_name) ?? "");

            if (configuration?.countries == null ||
                configuration.countries.Count == 0)
            {
                throw new InvalidOperationException(
                    global_application_name +
                    " parameters are empty or invalid");
            }

            foreach (Country item in configuration.countries)
            {
                ApplyCountryConfiguration(item, sql_con);

                if (!IsTrue(active))
                    continue;

                try
                {
                    ValidateCreditReviewMailConfiguration();

                    DateTime lastSuccessfulExecution =
                        GetCreditReviewScheduleDate(scheduleParameter);

                    DateTime currentDateTime = DateTime.Now;

                    TimeSpan configuredExecutionTime =
                        lastSuccessfulExecution.TimeOfDay;
                    DateTime nextExecutionTime =
                        currentDateTime.Date.Add(configuredExecutionTime);

                    WriteLog(
                        "   Credit Review daily schedule" +
                        " - Last successful execution : " +
                        lastSuccessfulExecution.ToString("dd/MM/yyyy HH:mm:ss") +
                        " - Current date : " +
                        currentDateTime.ToString("dd/MM/yyyy HH:mm:ss") +
                        " - Next execution time : " +
                        nextExecutionTime.ToString("dd/MM/yyyy HH:mm:ss"));

                    if (lastSuccessfulExecution.Date >= currentDateTime.Date)
                    {
                        WriteLog(
                            "   Credit Review daily processing already completed today");

                        continue;
                    }

                    if (currentDateTime < nextExecutionTime)
                    {
                        WriteLog(
                            "   Credit Review daily processing not yet allowed." +
                            " Next execution time : " +
                            nextExecutionTime.ToString("dd/MM/yyyy HH:mm:ss"));

                        continue;
                    }

                    graphService = ConnectGraph();

                    InitializeCreditReviewReminderDates();
                    CreateCreditReviewMailRequests();
                    SendPendingCreditReviewClientMails();
                    SendPendingCreditReviewImperativeMails();

                    DateTime completedAt = DateTime.Now;
                    DateTime scheduleReference =
                        completedAt.Date.Add(configuredExecutionTime);
                    SetCreditReviewScheduleDate(
                        scheduleParameter,
                        scheduleReference);

                    WriteLog(
                        "   Credit Review daily processing completed" +
                        " - Successful execution : " +
                        completedAt.ToString("dd/MM/yyyy HH:mm:ss"));
                }
                catch (Exception ex)
                {
                    string details = GetDetailedExceptionMessage(ex);

                    WriteLog(
                        "   Credit Review daily processing error" +
                        " - Schedule parameter not updated" +
                        " - Details : " + details);

                    SendTechnicalAlert(
                        nameof(Credit_Review_Mails),
                        details,
                        "CREDIT REVIEW DAILY PROCESSING");
                }
                finally
                {
                    graphService = null;
                }
            }
        }

        private sealed class CreditReviewReminderRow
        {
            public int DossierId { get; set; }
            public string ProcessingType { get; set; } = "";
            public DateTime RequestDate { get; set; }
            public DateTime ReminderJ7 { get; set; }
            public DateTime ReminderJ14 { get; set; }
            public DateTime ReminderJ21 { get; set; }
            public bool SentInitial { get; set; }
            public bool SentJ7 { get; set; }
            public bool SentJ14 { get; set; }
            public bool SentJ21 { get; set; }
        }

        private void ValidateCreditReviewMailConfiguration()
        {
            if (string.IsNullOrWhiteSpace(sql_connexion))
                throw new InvalidOperationException(
                    "sql_connexion is empty for Credit Review");

            if (string.IsNullOrWhiteSpace(sql_creation_compte))
                throw new InvalidOperationException(
                    "sql_creation_compte is empty for Credit Review schedule");

            if (string.IsNullOrWhiteSpace(templatesReportsFolder) ||
                !Directory.Exists(templatesReportsFolder))
                throw new DirectoryNotFoundException(
                    "Credit Review templates folder not found : " +
                    templatesReportsFolder);

            if (string.IsNullOrWhiteSpace(fr_credit_review_graph_send_as))
                throw new InvalidOperationException(
                    "fr_credit_review_graph_send_as is empty for Credit Review");
        }

        private DateTime GetCreditReviewScheduleDate(
            string parameterName)
        {
            string value = ExecuteScalarString(
                sql_creation_compte,
                @"SELECT ISNULL(ParameterValue, '')
                  FROM dbo.T_STATS_OUVERTURE_COMPTEUR_PARAMETERS
                  WHERE ParameterName=@NAME;",
                new SqlParameter("@NAME", SqlDbType.VarChar, 100)
                {
                    Value = parameterName
                });

            if (string.IsNullOrWhiteSpace(value))
                return new DateTime(1900, 1, 1);

            if (DateTime.TryParseExact(
                    value,
                    "yyyy-MM-ddTHH:mm:ss",
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AssumeLocal,
                    out DateTime exact))
                return exact;

            if (DateTime.TryParse(
                    value,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AssumeLocal,
                    out DateTime parsed))
                return parsed;

            throw new InvalidDataException(
                "Invalid Credit Review schedule parameter" +
                " - Parameter : " + parameterName +
                " - Value : " + value);
        }

        private void SetCreditReviewScheduleDate(
            string parameterName,
            DateTime executionDate)
        {
            ExecuteNonQuery(
                sql_creation_compte,
                @"MERGE dbo.T_STATS_OUVERTURE_COMPTEUR_PARAMETERS AS target
                  USING (SELECT @NAME AS ParameterName) AS source
                     ON source.ParameterName=target.ParameterName
                  WHEN MATCHED THEN
                      UPDATE SET ParameterValue=@VALUE,
                                 UpdatedAtUtc=SYSUTCDATETIME()
                  WHEN NOT MATCHED THEN
                      INSERT(ParameterName,ParameterValue,UpdatedAtUtc)
                      VALUES(@NAME,@VALUE,SYSUTCDATETIME());",
                new SqlParameter("@NAME", SqlDbType.VarChar, 100)
                {
                    Value = parameterName
                },
                new SqlParameter("@VALUE", SqlDbType.NVarChar, 2000)
                {
                    Value = executionDate.ToString(
                        "yyyy-MM-ddTHH:mm:ss",
                        CultureInfo.InvariantCulture)
                });
        }

        private void InitializeCreditReviewReminderDates()
        {
            const string sql = @"
UPDATE affectation
SET date_dde_bilan=GETDATE(),
    dt_relance_j7=DATEADD(day,7,GETDATE()),
    dt_relance_j14=DATEADD(day,14,GETDATE()),
    dt_relance_j21=DATEADD(day,21,GETDATE())
FROM dbo.T_Credit_Review_AFFECTATION AS affectation
INNER JOIN dbo.T_Credit_Review_ORT AS ort
    ON affectation.id_indice_ort=ort.indice_ORT
INNER JOIN dbo.T_Credit_Review_CONTACT_FICHE AS contact
    ON contact.code_client=ort.branche_nbr+ort.customer_nbr
INNER JOIN dbo.T_Credit_Review_CONTACT_FONCTION AS fonction
    ON fonction.id=contact.id_fonction
WHERE affectation.top_relance=1
  AND contact.top_relance=1
  AND fonction.is_recouvrement=0
  AND affectation.date_dde_bilan IS NULL
  AND affectation.traitement IS NOT NULL;";

            ExecuteNonQuery(sql_connexion, sql);
        }

        private void CreateCreditReviewMailRequests()
        {
            const string sql = @"
SELECT DISTINCT
       affectation.id_indice_ort,
       affectation.traitement,
       affectation.date_dde_bilan,
       affectation.dt_relance_j7,
       affectation.dt_relance_j14,
       affectation.dt_relance_j21,
       ISNULL(affectation.top_relance_j,0) AS top_relance_j,
       ISNULL(affectation.top_relance_j7,0) AS top_relance_j7,
       ISNULL(affectation.top_relance_j14,0) AS top_relance_j14,
       ISNULL(affectation.top_relance_j21,0) AS top_relance_j21
FROM dbo.T_Credit_Review_AFFECTATION AS affectation
INNER JOIN dbo.T_Credit_Review_ORT AS ort
    ON affectation.id_indice_ort=ort.indice_ORT
INNER JOIN dbo.T_Credit_Review_CONTACT_FICHE AS contact
    ON contact.code_client=ort.branche_nbr+ort.customer_nbr
INNER JOIN dbo.T_Credit_Review_CONTACT_FONCTION AS fonction
    ON fonction.id=contact.id_fonction
WHERE affectation.top_relance=1
  AND contact.top_relance=1
  AND fonction.is_recouvrement=0
  AND affectation.date_dde_bilan IS NOT NULL
  AND affectation.traitement IS NOT NULL
ORDER BY affectation.id_indice_ort;";

            DataTable rows = FillDataTable(sql_connexion, sql);
            int created = 0;

            foreach (DataRow row in rows.Rows)
            {
                var reminder = new CreditReviewReminderRow
                {
                    DossierId = Convert.ToInt32(row["id_indice_ort"]),
                    ProcessingType =
                        Convert.ToString(row["traitement"]).Trim(),
                    RequestDate = Convert.ToDateTime(row["date_dde_bilan"]),
                    ReminderJ7 = Convert.ToDateTime(row["dt_relance_j7"]),
                    ReminderJ14 = Convert.ToDateTime(row["dt_relance_j14"]),
                    ReminderJ21 = Convert.ToDateTime(row["dt_relance_j21"]),
                    SentInitial = Convert.ToBoolean(row["top_relance_j"]),
                    SentJ7 = Convert.ToBoolean(row["top_relance_j7"]),
                    SentJ14 = Convert.ToBoolean(row["top_relance_j14"]),
                    SentJ21 = Convert.ToBoolean(row["top_relance_j21"])
                };

                string templateName = "";
                string recipients = "";
                string updateSql = "";

                if (reminder.ReminderJ21 <= DateTime.Now &&
                    !reminder.SentJ21)
                {
                    templateName = "traitement_imperatif";
                    recipients = GetCreditReviewAnalystRecipients(
                        reminder.DossierId);
                    updateSql = @"
UPDATE dbo.T_Credit_Review_AFFECTATION
SET top_relance=0,
    top_relance_j21=1,
    top_relance_j14=1,
    top_relance_j7=1,
    top_relance_j=1
WHERE id_indice_ort=@DOSSIER;";
                }
                else if (reminder.ReminderJ14 <= DateTime.Now &&
                         !reminder.SentJ14)
                {
                    templateName = GetCreditReviewReminderTemplate(
                        reminder.ProcessingType,
                        false);
                    recipients = GetCreditReviewCustomerRecipients(
                        reminder.DossierId);
                    updateSql = @"
UPDATE dbo.T_Credit_Review_AFFECTATION
SET top_relance_j14=1,
    top_relance_j7=1,
    top_relance_j=1
WHERE id_indice_ort=@DOSSIER;";
                }
                else if (reminder.ReminderJ7 <= DateTime.Now &&
                         !reminder.SentJ7)
                {
                    templateName = GetCreditReviewReminderTemplate(
                        reminder.ProcessingType,
                        false);
                    recipients = GetCreditReviewCustomerRecipients(
                        reminder.DossierId);
                    updateSql = @"
UPDATE dbo.T_Credit_Review_AFFECTATION
SET top_relance_j7=1,
    top_relance_j=1
WHERE id_indice_ort=@DOSSIER;";
                }
                else if (reminder.RequestDate <= DateTime.Now &&
                         !reminder.SentInitial)
                {
                    templateName = GetCreditReviewReminderTemplate(
                        reminder.ProcessingType,
                        true);
                    recipients = GetCreditReviewCustomerRecipients(
                        reminder.DossierId);
                    updateSql = @"
UPDATE dbo.T_Credit_Review_AFFECTATION
SET top_relance_j=1
WHERE id_indice_ort=@DOSSIER;";
                }

                if (string.IsNullOrWhiteSpace(templateName))
                    continue;

                if (BuildRecipients(recipients).Count == 0)
                {
                    throw new InvalidOperationException(
                        "No valid Credit Review recipient" +
                        " - Dossier : " + reminder.DossierId +
                        " - Template : " + templateName);
                }

                using (var connection = new SqlConnection(sql_connexion))
                {
                    connection.Open();
                    using (SqlTransaction transaction =
                        connection.BeginTransaction())
                    {
                        try
                        {
                            using (var insert = new SqlCommand(@"
IF NOT EXISTS
(
    SELECT 1
    FROM dbo.T_Credit_Review_ENVOI_MAIL
    WHERE id_dossier=@DOSSIER
      AND nom_mail=@TEMPLATE
      AND top_traite IN ('N','A','O')
      AND CONVERT(date,date_demande)=CONVERT(date,GETDATE())
)
INSERT INTO dbo.T_Credit_Review_ENVOI_MAIL
(
    id_dossier,
    nom_mail,
    email_destinataire,
    date_demande,
    top_traite
)
VALUES
(
    @DOSSIER,
    @TEMPLATE,
    @RECIPIENTS,
    GETDATE(),
    'N'
);", connection, transaction))
                            {
                                insert.Parameters.Add(
                                    "@DOSSIER",
                                    SqlDbType.Int).Value =
                                    reminder.DossierId;
                                insert.Parameters.Add(
                                    "@TEMPLATE",
                                    SqlDbType.VarChar,
                                    100).Value = templateName;
                                insert.Parameters.Add(
                                    "@RECIPIENTS",
                                    SqlDbType.NVarChar,
                                    -1).Value = recipients;
                                insert.ExecuteNonQuery();
                            }

                            using (var update = new SqlCommand(
                                updateSql,
                                connection,
                                transaction))
                            {
                                update.Parameters.Add(
                                    "@DOSSIER",
                                    SqlDbType.Int).Value =
                                    reminder.DossierId;
                                update.ExecuteNonQuery();
                            }

                            transaction.Commit();
                            created++;
                        }
                        catch
                        {
                            transaction.Rollback();
                            throw;
                        }
                    }
                }
            }

            WriteLog(
                "       Credit Review request generation completed" +
                " - Candidate(s) : " + rows.Rows.Count +
                " - Request(s) created or confirmed : " + created);
        }

        private static string GetCreditReviewReminderTemplate(
            string processingType,
            bool initialRequest)
        {
            if (string.Equals(
                    processingType,
                    "Situation",
                    StringComparison.OrdinalIgnoreCase))
            {
                return initialRequest
                    ? "demande_bilan_intermediaire"
                    : "relance_intermediaire";
            }

            if (string.Equals(
                    processingType,
                    "Bilan",
                    StringComparison.OrdinalIgnoreCase))
            {
                return initialRequest
                    ? "demande_bilan"
                    : "relance";
            }

            return "";
        }

        private string GetCreditReviewCustomerRecipients(int dossierId)
        {
            DataTable table = FillDataTable(
                sql_connexion,
                @"SELECT DISTINCT contact.email
                  FROM dbo.T_Credit_Review_CONTACT_FICHE AS contact
                  INNER JOIN dbo.T_Credit_Review_ORT AS ort
                      ON contact.code_client=ort.branche_nbr+ort.customer_nbr
                  INNER JOIN dbo.T_Credit_Review_CONTACT_FONCTION AS fonction
                      ON fonction.id=contact.id_fonction
                  WHERE ort.indice_ort=@DOSSIER
                    AND contact.top_relance=1
                    AND fonction.is_recouvrement=0
                    AND NULLIF(LTRIM(RTRIM(contact.email)),'') IS NOT NULL;",
                new SqlParameter("@DOSSIER", SqlDbType.Int)
                {
                    Value = dossierId
                });

            return string.Join(
                ";",
                table.AsEnumerable()
                    .Select(row => Convert.ToString(row["email"]).Trim())
                    .Where(IsEmail)
                    .Distinct(StringComparer.OrdinalIgnoreCase));
        }

        private string GetCreditReviewAnalystRecipients(int dossierId)
        {
            DataTable table = FillDataTable(
                sql_connexion,
                @"SELECT DISTINCT analyste.email
                  FROM dbo.T_Credit_Review_ANALYSTE AS analyste
                  INNER JOIN dbo.T_Credit_Review_AFFECTATION AS affectation
                      ON analyste.id=affectation.id_analyste
                  WHERE affectation.id_indice_ort=@DOSSIER
                    AND NULLIF(LTRIM(RTRIM(analyste.email)),'') IS NOT NULL;",
                new SqlParameter("@DOSSIER", SqlDbType.Int)
                {
                    Value = dossierId
                });

            return string.Join(
                ";",
                table.AsEnumerable()
                    .Select(row => Convert.ToString(row["email"]).Trim())
                    .Where(IsEmail)
                    .Distinct(StringComparer.OrdinalIgnoreCase));
        }

        private void SendPendingCreditReviewClientMails()
        {
            DataTable queue = FillDataTable(
                sql_connexion,
                @"SELECT sending.id,
                         sending.id_dossier,
                         sending.nom_mail,
                         sending.email_destinataire,
                         sending.top_traite,
                         sending.date_demande,
                         link.[sujet email] AS subject,
                         link.nom_fic_html
                  FROM dbo.T_Credit_Review_ENVOI_MAIL AS sending
                  INNER JOIN dbo.T_Credit_Review_LIEN_EMAIL AS link
                      ON link.nom_email_court=sending.nom_mail
                  WHERE sending.top_traite IN ('N','A')
                    AND sending.nom_mail<>'traitement_imperatif'
                  ORDER BY sending.id;");

            int sent = 0;
            int errors = 0;

            foreach (DataRow row in queue.Rows)
            {
                int queueId = Convert.ToInt32(row["id"]);
                int dossierId = Convert.ToInt32(row["id_dossier"]);
                string templateName =
                    Convert.ToString(row["nom_mail"]).Trim();
                string recipients =
                    Convert.ToString(row["email_destinataire"]).Trim();
                string queueStatus =
                    Convert.ToString(row["top_traite"]).Trim();
                DateTime requestDate =
                    row["date_demande"] == DBNull.Value
                        ? DateTime.Today
                        : Convert.ToDateTime(row["date_demande"]);

                try
                {
                    string templatePath = SafeResourcePath(
                        Convert.ToString(row["nom_fic_html"]));
                    string html = File.ReadAllText(
                        templatePath,
                        Encoding.UTF8);

                    DataTable dossier = FillDataTable(
                        sql_connexion,
                        @"SELECT TOP(1)
                                 customer_nbr AS code_client,
                                 ISNULL(nom_client,'') AS nom_client,
                                 ISNULL(date_mois,'') AS date_mois
                          FROM dbo.T_Credit_Review_ORT
                          WHERE indice_ORT=@DOSSIER;",
                        new SqlParameter("@DOSSIER", SqlDbType.Int)
                        {
                            Value = dossierId
                        });

                    if (dossier.Rows.Count == 0)
                        throw new InvalidOperationException(
                            "Credit Review dossier not found : " + dossierId);

                    string customerCode =
                        Convert.ToString(dossier.Rows[0]["code_client"]).Trim();
                    string customerName =
                        Convert.ToString(dossier.Rows[0]["nom_client"]).Trim();
                    string subject =
                        Convert.ToString(row["subject"])
                            .Replace("[CODE_CLIENT]", customerCode)
                            .Replace("[NOM_CLIENT]", customerName);

                    html = html.Replace(
                        "[CODE_CLIENT]",
                        customerCode);

                    string creditReviewPeriod =
                        Convert.ToString(
                            dossier.Rows[0]["date_mois"]).Trim();

                    if (queueStatus.Equals(
                            "A",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        ArchiveExistingCreditReviewSentMail(
                            fr_credit_review_graph_send_as,
                            recipients,
                            subject,
                            creditReviewPeriod,
                            requestDate);
                    }
                    else
                    {
                        SendCreditReviewMailAndArchive(
                            queueId,
                            fr_credit_review_graph_send_as,
                            recipients,
                            subject,
                            html,
                            creditReviewPeriod);
                    }

                    SetCreditReviewQueueStatus(queueId, "O");
                    sent++;

                    WriteLog(
                        "       Credit Review client email sent" +
                        " - Queue ID : " + queueId +
                        " - Dossier : " + dossierId +
                        " - Template : " + templateName +
                        " - Recipient(s) : " + recipients);
                }
                catch (Exception ex)
                {
                    errors++;

                    string currentStatus =
                        GetCreditReviewQueueStatus(queueId);

                    if (!currentStatus.Equals(
                            "A",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        SetCreditReviewQueueStatus(queueId, "E");
                    }

                    WriteLog(
                        "       Credit Review client email error" +
                        " - Current status : " + currentStatus +
                        " - Queue ID : " + queueId +
                        " - Dossier : " + dossierId +
                        " - Error : " +
                        GetDetailedExceptionMessage(ex));
                }
            }

            WriteLog(
                "       Credit Review client queue summary" +
                " - Found : " + queue.Rows.Count +
                " - Sent : " + sent +
                " - Errors : " + errors);

            if (errors > 0)
                throw new InvalidOperationException(
                    errors +
                    " Credit Review client email(s) could not be sent");
        }

        private void SendPendingCreditReviewImperativeMails()
        {
            DataTable analysts = FillDataTable(
                sql_connexion,
                @"SELECT DISTINCT email_destinataire
                  FROM dbo.T_Credit_Review_ENVOI_MAIL
                  WHERE top_traite='N'
                    AND nom_mail='traitement_imperatif'
                    AND NULLIF(LTRIM(RTRIM(email_destinataire)),'') IS NOT NULL
                  ORDER BY email_destinataire;");

            int sent = 0;

            foreach (DataRow analystRow in analysts.Rows)
            {
                string recipient =
                    Convert.ToString(
                        analystRow["email_destinataire"]).Trim();

                DataTable dossiers = FillDataTable(
                    sql_connexion,
                    @"SELECT sending.id,
                             ort.date_mois,
                             ort.branche_nbr+ort.customer_nbr AS code_client
                      FROM dbo.T_Credit_Review_ENVOI_MAIL AS sending
                      INNER JOIN dbo.T_Credit_Review_ORT AS ort
                          ON sending.id_dossier=ort.indice_ort
                      WHERE sending.email_destinataire=@EMAIL
                        AND sending.top_traite='N'
                        AND sending.nom_mail='traitement_imperatif'
                      ORDER BY ort.date_mois,sending.id;",
                    new SqlParameter("@EMAIL", SqlDbType.NVarChar, -1)
                    {
                        Value = recipient
                    });

                if (dossiers.Rows.Count == 0)
                    continue;

                var body = new StringBuilder();
                body.Append(
                    "Merci de traiter ces dossiers impérativement !" +
                    "<br/><br/>");

                foreach (DataRow dossier in dossiers.Rows)
                {
                    body.Append(WebUtility.HtmlEncode(
                        Convert.ToString(dossier["date_mois"])));
                    body.Append(" - ");
                    body.Append(WebUtility.HtmlEncode(
                        Convert.ToString(dossier["code_client"])));
                    body.Append("<br/>");
                }

                SendGraphMailWithAttachments(
                    fr_credit_review_graph_send_as,
                    recipient,
                    "",
                    "",
                    "Credit Review - Traitement Impératif",
                    body.ToString(),
                    new List<string>());

                foreach (DataRow dossier in dossiers.Rows)
                {
                    SetCreditReviewQueueStatus(
                        Convert.ToInt32(dossier["id"]),
                        "O");
                }

                sent++;

                WriteLog(
                    "       Credit Review imperative email sent" +
                    " - Analyst : " + recipient +
                    " - Dossier(s) : " + dossiers.Rows.Count);
            }

            WriteLog(
                "       Credit Review imperative queue summary" +
                " - Analyst email(s) sent : " + sent);
        }

        private void SetCreditReviewQueueStatus(
            int queueId,
            string status)
        {
            ExecuteNonQuery(
                sql_connexion,
                @"UPDATE dbo.T_Credit_Review_ENVOI_MAIL
                  SET top_traite=@STATUS,
                      date_trt=GETDATE()
                  WHERE id=@ID;",
                new SqlParameter("@STATUS", SqlDbType.Char, 1)
                {
                    Value = status
                },
                new SqlParameter("@ID", SqlDbType.Int)
                {
                    Value = queueId
                });
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
WHERE NULLIF(LTRIM(RTRIM(s.Comment)), '') IS NULL
  AND EXISTS(SELECT 1 FROM dbo.envoi_mail e WHERE e.id_dossier=s.idDossier AND e.top_traite='N' AND e.nom_mail IN('refus_expert','refus_export'));";
            ExecuteNonQuery(sql_creation_compte, refusal);
        }

        private void ProcessPendingOpeningTemplates()
        {
            const string sql = @"
SELECT e.id,
       e.id_dossier,
       e.nom_mail,
       e.email_destinataire,
       e.top_traite,
       e.date_demande,
       l.pieces_jointes,
       l.[sujet email] AS sujet_email,
       l.nom_fic_html,
       s.idStatut
FROM dbo.envoi_mail AS e
INNER JOIN dbo.T_statut AS s
    ON s.idDossier = e.id_dossier
INNER JOIN dbo.lien_email AS l
    ON l.nom_email_court = e.nom_mail
WHERE e.top_traite IN ('N','A')
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
                string queueStatus = Convert.ToString(row["top_traite"]).Trim();
                DateTime requestDate = row["date_demande"] == DBNull.Value
                    ? DateTime.Today
                    : Convert.ToDateTime(row["date_demande"]);
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

                    if (mailName.Equals(
                            "rgpd",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        int divisionCount;
                        html = ApplyRgpdDivisionRows(html, dossierId, out divisionCount);
                        WriteLog(
                            "       RGPD divisions inserted" +
                            " - Dossier : " + dossierId +
                            " - Division(s) : " + divisionCount);
                    }

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

                    if (queueStatus.Equals(
                            "A",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        ArchiveExistingSentOpeningMail(
                            id,
                            dossierId,
                            sendAs,
                            recipient,
                            subject,
                            requestDate);
                    }
                    else
                    {
                        SendOpeningMailAndArchiveInDatabase(
                            id,
                            dossierId,
                            sendAs,
                            recipient,
                            "",
                            bcc,
                            subject,
                            html,
                            attachments);
                    }

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
                    string currentQueueStatus = GetOpeningQueueStatus(id);
                    if (!currentQueueStatus.Equals(
                            "A",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        SetMailRequestStatus("dbo.envoi_mail", id, "E");
                    }
                    WriteLog(
                        "       Opening account template error" +
                        " - Current status : " + currentQueueStatus +
                        " - Queue ID : " + id +
                        " - Dossier : " + dossierId +
                        " - Template : " + mailName +
                        " - Recipient : " + recipient +
                        " - Error : " + ex.Message);

                    bool isMissingRgpdDivision =
                        mailName.Equals(
                            "rgpd",
                            StringComparison.OrdinalIgnoreCase) &&
                        ex.Message.IndexOf(
                            "No active RGPD subscription division found",
                            StringComparison.OrdinalIgnoreCase) >= 0;

                    if (isMissingRgpdDivision)
                    {
                        WriteLog(
                            "       RGPD technical alert suppressed" +
                            " - Queue ID : " + id +
                            " - Dossier : " + dossierId +
                            " - Reason : no active RGPD subscription division");
                    }
                    else
                    {
                        SendTechnicalAlert(
                            nameof(ProcessPendingOpeningTemplates),
                            "Dossier " + dossierId + " - " + ex.Message,
                            "OPENING TEMPLATE");
                    }
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

        private string ApplyRgpdDivisionRows(
            string html,
            int dossierId,
            out int divisionCount)
        {
            List<string> divisions = GetRgpdDivisions(dossierId);
            divisionCount = divisions.Count;
            if (divisions.Count == 0)
                throw new InvalidOperationException(
                    "No active RGPD subscription division found - Dossier : " + dossierId);

            string sourceHtml = html ?? "";
            Match rowMatch = Regex.Match(
                sourceHtml,
                "<tr\\b[^>]*\\bid\\s*=\\s*['\"]ligne_tab['\"][^>]*>.*?</tr>",
                RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.CultureInvariant);
            if (!rowMatch.Success)
                throw new InvalidDataException(
                    "The RGPD template does not contain the historical table row id=\"ligne_tab\" - Dossier : " + dossierId);
            if (rowMatch.Value.IndexOf("[DIVISION]", StringComparison.OrdinalIgnoreCase) < 0)
                throw new InvalidDataException(
                    "The RGPD template row id=\"ligne_tab\" does not contain [DIVISION] - Dossier : " + dossierId);

            var generatedRows = new StringBuilder();
            foreach (string division in divisions)
                generatedRows.Append(Regex.Replace(
                    rowMatch.Value,
                    "\\[DIVISION\\]",
                    WebUtility.HtmlEncode(division),
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant));

            return sourceHtml.Substring(0, rowMatch.Index) +
                generatedRows +
                sourceHtml.Substring(rowMatch.Index + rowMatch.Length);
        }

        private List<string> GetRgpdDivisions(int dossierId)
        {
            DataTable divisions = FillDataTable(
                sql_creation_compte,
                @"SELECT DISTINCT LTRIM(RTRIM(category.lib_cat)) AS lib_cat
                  FROM dbo.T_customer_choix_abonnement AS subscription
                  INNER JOIN dbo.T_liste_categorie_MKT AS category
                      ON category.id_cat=subscription.id_cat
                  WHERE subscription.id=@DOSSIER_ID
                    AND NULLIF(LTRIM(RTRIM(category.lib_cat)),'') IS NOT NULL
                  ORDER BY LTRIM(RTRIM(category.lib_cat));",
                new SqlParameter("@DOSSIER_ID", SqlDbType.Int) { Value = dossierId });
            return divisions.AsEnumerable()
                .Select(row => Convert.ToString(row["lib_cat"]).Trim())
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
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
        private static void ValidateNoKnownUnresolvedMarkers(string html, string template) { Match m = Regex.Match(html ?? "", @"\[(NOM|PRENOM|CODE_CLIENT|NOM_CLIENT|RCS|BANQUE|AGENCE|IBAN|BIC|MOTIF|DATE_DEMANDE|MOTIF_REFUS|NUM_DEMANDE|DIVISION|GESTIONNAIRE|NUM_POSTE|MAIL_GESTIONNAIRE)\]", RegexOptions.IgnoreCase); if (m.Success) throw new InvalidDataException("Unresolved marker " + m.Value + " in " + template); }
        private static DataTable FillDataTable(string cs, string sql, params SqlParameter[] parameters) { var t = new DataTable(); using (var c = new SqlConnection(cs)) using (var cmd = new SqlCommand(sql, c)) using (var da = new SqlDataAdapter(cmd)) { cmd.CommandTimeout = 300; if (parameters != null && parameters.Length > 0) cmd.Parameters.AddRange(parameters); da.Fill(t); } return t; }

        private string GetCreditReviewQueueStatus(int queueId)
        {
            return ExecuteScalarString(
                sql_connexion,
                @"SELECT ISNULL(top_traite,'')
                  FROM dbo.T_Credit_Review_ENVOI_MAIL
                  WHERE id=@ID;",
                new SqlParameter("@ID", SqlDbType.Int)
                {
                    Value = queueId
                });
        }

        private void ArchiveExistingCreditReviewSentMail(
            string mailbox,
            string recipients,
            string subject,
            string creditReviewPeriod,
            DateTime requestDate)
        {
            string periodFolderName =
                string.IsNullOrWhiteSpace(creditReviewPeriod)
                    ? requestDate.ToString("yyyyMM", CultureInfo.InvariantCulture)
                    : creditReviewPeriod.Trim();

            MailFolder sentItems = ExecuteGraphWithRetry(
                () => graphService.Users[mailbox]
                    .MailFolders["sentitems"]
                    .GetAsync(config => AddImmutableHeader(config.Headers))
                    .GetAwaiter()
                    .GetResult(),
                "Read Credit Review Sent Items folder for archive retry");

            MailFolder archiveRoot = GetOrCreateCreditReviewFolder(
                mailbox,
                sentItems.Id,
                "Archives_mails");

            MailFolder periodFolder = GetOrCreateCreditReviewFolder(
                mailbox,
                archiveRoot.Id,
                periodFolderName);

            string filter =
                "subject eq '" + EscapeODataString(subject) + "'" +
                " and sentDateTime ge " +
                requestDate.Date.ToUniversalTime().ToString(
                    "yyyy-MM-ddTHH:mm:ssZ",
                    CultureInfo.InvariantCulture);

            MessageCollectionResponse response = ExecuteGraphWithRetry(
                () => graphService.Users[mailbox]
                    .MailFolders[sentItems.Id]
                    .Messages
                    .GetAsync(config =>
                    {
                        AddImmutableHeader(config.Headers);
                        config.QueryParameters.Top = 25;
                        config.QueryParameters.Filter = filter;
                        config.QueryParameters.Select = new[]
                        {
                            "id",
                            "subject",
                            "sentDateTime",
                            "toRecipients",
                            "parentFolderId",
                            "isDraft"
                        };
                    })
                    .GetAwaiter()
                    .GetResult(),
                "Find sent Credit Review email for archive retry");

            HashSet<string> expectedRecipients =
                BuildRecipients(recipients)
                    .Select(recipient =>
                        recipient.EmailAddress?.Address ?? "")
                    .Where(address =>
                        !string.IsNullOrWhiteSpace(address))
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

            Message sentMessage =
                (response?.Value ?? new List<Message>())
                    .Where(message => message.IsDraft != true)
                    .OrderByDescending(message => message.SentDateTime)
                    .FirstOrDefault(message =>
                    {
                        HashSet<string> actualRecipients =
                            (message.ToRecipients ?? new List<Recipient>())
                                .Select(recipient =>
                                    recipient.EmailAddress?.Address ?? "")
                                .Where(address =>
                                    !string.IsNullOrWhiteSpace(address))
                                .ToHashSet(StringComparer.OrdinalIgnoreCase);

                        return expectedRecipients.SetEquals(actualRecipients);
                    });

            if (sentMessage == null)
            {
                throw new InvalidOperationException(
                    "Sent Credit Review email not found for archive retry" +
                    " - Subject : " + subject +
                    " - Recipient(s) : " + recipients +
                    " - Request date : " +
                    requestDate.ToString("dd/MM/yyyy HH:mm:ss"));
            }

            MoveSentCreditReviewMessage(
                mailbox,
                sentMessage.Id,
                periodFolder.Id,
                subject);

            WriteLog(
                "       Credit Review archive retry completed" +
                " - Subject : " + subject +
                " - Recipient(s) : " + recipients +
                " - Folder : Archives_mails/" + periodFolderName);
        }

        private void SendCreditReviewMailAndArchive(
            int queueId,
            string sendAs,
            string recipients,
            string subject,
            string html,
            string creditReviewPeriod)
        {
            List<Recipient> toRecipients = BuildRecipients(recipients);

            if (toRecipients.Count == 0)
            {
                throw new InvalidOperationException(
                    "Invalid Credit Review recipients for " + subject);
            }

            string periodFolderName =
                string.IsNullOrWhiteSpace(creditReviewPeriod)
                    ? DateTime.Now.ToString("yyyyMM", CultureInfo.InvariantCulture)
                    : creditReviewPeriod.Trim();

            MailFolder sentItems = ExecuteGraphWithRetry(
                () => graphService.Users[sendAs]
                    .MailFolders["sentitems"]
                    .GetAsync(config => AddImmutableHeader(config.Headers))
                    .GetAwaiter()
                    .GetResult(),
                "Read Credit Review Sent Items folder");

            MailFolder archiveRoot = GetOrCreateCreditReviewFolder(
                sendAs,
                sentItems.Id,
                "Archives_mails");

            MailFolder periodFolder = GetOrCreateCreditReviewFolder(
                sendAs,
                archiveRoot.Id,
                periodFolderName);

            var draftMessage = new Message
            {
                Subject = subject,
                From = new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = sendAs,
                        Name = "Ingram Micro - Service Analyse Crédit"
                    }
                },
                Sender = new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = sendAs,
                        Name = "Ingram Micro - Service Analyse Crédit"
                    }
                },
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = html
                },
                ToRecipients = toRecipients
            };

            Message draft = ExecuteGraphWithRetry(
                () => graphService.Users[sendAs]
                    .Messages
                    .PostAsync(
                        draftMessage,
                        config => AddImmutableHeader(config.Headers))
                    .GetAwaiter()
                    .GetResult(),
                "Create Credit Review draft : " + subject);

            if (draft == null || string.IsNullOrWhiteSpace(draft.Id))
            {
                throw new InvalidOperationException(
                    "Credit Review draft creation returned no message ID");
            }

            ExecuteGraphWithRetry(
                () =>
                {
                    graphService.Users[sendAs]
                        .Messages[draft.Id]
                        .Send
                        .PostAsync(config => AddImmutableHeader(config.Headers))
                        .GetAwaiter()
                        .GetResult();
                    return true;
                },
                "Send Credit Review draft : " + subject);

            // L'envoi est accepté par Graph. Le statut A interdit tout
            // nouvel envoi si l'archivage échoue ensuite.
            SetCreditReviewQueueStatus(queueId, "A");

            MoveSentCreditReviewMessage(
                sendAs,
                draft.Id,
                periodFolder.Id,
                subject);
        }

        private MailFolder GetOrCreateCreditReviewFolder(
            string mailbox,
            string parentFolderId,
            string folderName)
        {
            string safeFolderName = EscapeODataString(folderName);

            MailFolderCollectionResponse response = ExecuteGraphWithRetry(
                () => graphService.Users[mailbox]
                    .MailFolders[parentFolderId]
                    .ChildFolders
                    .GetAsync(config =>
                    {
                        AddImmutableHeader(config.Headers);
                        config.QueryParameters.Top = 100;
                        config.QueryParameters.Filter =
                            "displayName eq '" + safeFolderName + "'";
                    })
                    .GetAwaiter()
                    .GetResult(),
                "Find Credit Review folder " + folderName);

            MailFolder existingFolder = response?.Value?.FirstOrDefault(
                folder => string.Equals(
                    folder.DisplayName,
                    folderName,
                    StringComparison.OrdinalIgnoreCase));

            if (existingFolder != null)
            {
                return existingFolder;
            }

            try
            {
                return ExecuteGraphWithRetry(
                    () => graphService.Users[mailbox]
                        .MailFolders[parentFolderId]
                        .ChildFolders
                        .PostAsync(
                            new MailFolder { DisplayName = folderName },
                            config => AddImmutableHeader(config.Headers))
                        .GetAwaiter()
                        .GetResult(),
                    "Create Credit Review folder " + folderName);
            }
            catch (Exception)
            {
                // Une autre exécution peut avoir créé le dossier entre la
                // recherche et la création. Une seconde lecture tranche.
                response = ExecuteGraphWithRetry(
                    () => graphService.Users[mailbox]
                        .MailFolders[parentFolderId]
                        .ChildFolders
                        .GetAsync(config =>
                        {
                            AddImmutableHeader(config.Headers);
                            config.QueryParameters.Top = 100;
                            config.QueryParameters.Filter =
                                "displayName eq '" + safeFolderName + "'";
                        })
                        .GetAwaiter()
                        .GetResult(),
                    "Reload Credit Review folder " + folderName);

                existingFolder = response?.Value?.FirstOrDefault(
                    folder => string.Equals(
                        folder.DisplayName,
                        folderName,
                        StringComparison.OrdinalIgnoreCase));

                if (existingFolder != null)
                {
                    return existingFolder;
                }

                throw;
            }
        }

        private void MoveSentCreditReviewMessage(
            string mailbox,
            string immutableMessageId,
            string destinationFolderId,
            string subject)
        {
            Exception lastException = null;

            for (int attempt = 1; attempt <= 10; attempt++)
            {
                try
                {
                    Message sentMessage = graphService.Users[mailbox]
                        .Messages[immutableMessageId]
                        .GetAsync(config =>
                        {
                            AddImmutableHeader(config.Headers);
                            config.QueryParameters.Select = new[]
                            {
                                "id",
                                "parentFolderId",
                                "subject",
                                "isDraft"
                            };
                        })
                        .GetAwaiter()
                        .GetResult();

                    if (sentMessage != null &&
                        sentMessage.IsDraft != true)
                    {
                        if (string.Equals(
                                sentMessage.ParentFolderId,
                                destinationFolderId,
                                StringComparison.Ordinal))
                        {
                            return;
                        }

                        var requestBody =
                            new Microsoft.Graph.Users.Item.Messages.Item.Move
                                .MovePostRequestBody
                            {
                                DestinationId = destinationFolderId
                            };

                        graphService.Users[mailbox]
                            .Messages[immutableMessageId]
                            .Move
                            .PostAsync(
                                requestBody,
                                config => AddImmutableHeader(config.Headers))
                            .GetAwaiter()
                            .GetResult();

                        WriteLog(
                            "       Credit Review sent email archived" +
                            " - Folder ID : " + destinationFolderId +
                            " - Subject : " + subject);
                        return;
                    }
                }
                catch (Exception ex)
                {
                    lastException = ex;

                    if (!IsGraphObjectNotFound(ex) &&
                        !IsTransientGraphError(ex))
                    {
                        throw;
                    }
                }

                System.Threading.Thread.Sleep(attempt * 500);
            }

            throw new InvalidOperationException(
                "The Credit Review email was sent but its Sent Items copy " +
                "could not be archived" +
                " - Subject : " + subject +
                " - Details : " +
                GetDetailedExceptionMessage(lastException),
                lastException);
        }

        private string GetOpeningQueueStatus(int queueId)
        {
            return ExecuteScalarString(
                sql_creation_compte,
                @"SELECT ISNULL(top_traite,'')
                  FROM dbo.envoi_mail
                  WHERE id=@ID;",
                new SqlParameter("@ID", SqlDbType.Int)
                {
                    Value = queueId
                });
        }

        private void SendOpeningMailAndArchiveInDatabase(
            int queueId,
            int dossierId,
            string sendAs,
            string to,
            string cc,
            string bcc,
            string subject,
            string html,
            List<string> files)
        {
            Message draftMessage = BuildOpeningOutgoingMessage(
                sendAs,
                to,
                cc,
                bcc,
                subject,
                html,
                files);

            Message draft = ExecuteGraphWithRetry(
                () => graphService.Users[sendAs]
                    .Messages
                    .PostAsync(
                        draftMessage,
                        config => AddImmutableHeader(config.Headers))
                    .GetAwaiter()
                    .GetResult(),
                "Create opening account draft : " + subject);

            if (draft == null || string.IsNullOrWhiteSpace(draft.Id))
            {
                throw new InvalidOperationException(
                    "Opening account draft creation returned no message ID");
            }

            ExecuteGraphWithRetry(
                () =>
                {
                    graphService.Users[sendAs]
                        .Messages[draft.Id]
                        .Send
                        .PostAsync(config => AddImmutableHeader(config.Headers))
                        .GetAwaiter()
                        .GetResult();
                    return true;
                },
                "Send opening account draft : " + subject);

            // Le message a été accepté par Graph. Le statut A empêche tout
            // nouvel envoi si l'archivage MIME échoue ensuite.
            SetMailRequestStatus("dbo.envoi_mail", queueId, "A");

            Message sentMessage = WaitForSentOpeningMessage(
                sendAs,
                draft.Id,
                subject);

            ArchiveSentOpeningMessageInDatabase(
                queueId,
                dossierId,
                sendAs,
                sentMessage,
                to);
        }

        private Message BuildOpeningOutgoingMessage(
            string sendAs,
            string to,
            string cc,
            string bcc,
            string subject,
            string html,
            List<string> files)
        {
            List<Recipient> toRecipients = BuildRecipients(to);
            if (toRecipients.Count == 0)
            {
                throw new InvalidOperationException(
                    "Invalid opening account recipients for " + subject);
            }

            string displayName = string.Equals(
                    sendAs,
                    fr_credit_administration_graph_send_as,
                    StringComparison.OrdinalIgnoreCase)
                ? "Ingram Micro - Changement de domiciliation bancaire"
                : "Ingram Micro - Service Nouveaux clients";

            var sender = new Recipient
            {
                EmailAddress = new EmailAddress
                {
                    Address = sendAs,
                    Name = displayName
                }
            };

            var message = new Message
            {
                Subject = subject,
                From = sender,
                Sender = sender,
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = html
                },
                ToRecipients = toRecipients,
                CcRecipients = BuildRecipients(cc),
                BccRecipients = BuildRecipients(bcc),
                Attachments = new List<Microsoft.Graph.Models.Attachment>()
            };

            foreach (string path in files ?? new List<string>())
            {
                byte[] bytes = File.ReadAllBytes(path);
                if (bytes.Length > 3 * 1024 * 1024)
                {
                    throw new InvalidDataException(
                        "Attachment exceeds 3 MB : " + Path.GetFileName(path));
                }

                message.Attachments.Add(new FileAttachment
                {
                    OdataType = "#microsoft.graph.fileAttachment",
                    Name = Path.GetFileName(path),
                    ContentType = "application/octet-stream",
                    ContentBytes = bytes
                });
            }

            return message;
        }

        private Message WaitForSentOpeningMessage(
            string mailbox,
            string immutableMessageId,
            string subject)
        {
            Exception lastException = null;

            for (int attempt = 1; attempt <= 12; attempt++)
            {
                try
                {
                    Message message = graphService.Users[mailbox]
                        .Messages[immutableMessageId]
                        .GetAsync(config =>
                        {
                            AddImmutableHeader(config.Headers);
                            config.QueryParameters.Select = new[]
                            {
                                "id",
                                "subject",
                                "sentDateTime",
                                "createdDateTime",
                                "toRecipients",
                                "isDraft"
                            };
                        })
                        .GetAwaiter()
                        .GetResult();

                    if (message != null && message.IsDraft != true)
                    {
                        return message;
                    }
                }
                catch (Exception ex)
                {
                    lastException = ex;
                    if (!IsGraphObjectNotFound(ex) &&
                        !IsTransientGraphError(ex))
                    {
                        throw;
                    }
                }

                System.Threading.Thread.Sleep(attempt * 500);
            }

            throw new InvalidOperationException(
                "The opening account email was sent but its Sent Items copy " +
                "could not be loaded" +
                " - Subject : " + subject +
                " - Details : " +
                GetDetailedExceptionMessage(lastException),
                lastException);
        }

        private void ArchiveExistingSentOpeningMail(
            int queueId,
            int dossierId,
            string mailbox,
            string recipients,
            string subject,
            DateTime requestDate)
        {
            MailFolder sentItems = ExecuteGraphWithRetry(
                () => graphService.Users[mailbox]
                    .MailFolders["sentitems"]
                    .GetAsync(config => AddImmutableHeader(config.Headers))
                    .GetAwaiter()
                    .GetResult(),
                "Read opening account Sent Items folder for archive retry");

            string filter =
                "subject eq '" + EscapeODataString(subject) + "'" +
                " and sentDateTime ge " +
                requestDate.Date.ToUniversalTime().ToString(
                    "yyyy-MM-ddTHH:mm:ssZ",
                    CultureInfo.InvariantCulture);

            MessageCollectionResponse response = ExecuteGraphWithRetry(
                () => graphService.Users[mailbox]
                    .MailFolders[sentItems.Id]
                    .Messages
                    .GetAsync(config =>
                    {
                        AddImmutableHeader(config.Headers);
                        config.QueryParameters.Top = 25;
                        config.QueryParameters.Filter = filter;
                        config.QueryParameters.Select = new[]
                        {
                            "id",
                            "subject",
                            "sentDateTime",
                            "createdDateTime",
                            "toRecipients",
                            "isDraft"
                        };
                    })
                    .GetAwaiter()
                    .GetResult(),
                "Find sent opening account email for database archive retry");

            HashSet<string> expectedRecipients = BuildRecipients(recipients)
                .Select(recipient => recipient.EmailAddress?.Address ?? "")
                .Where(address => !string.IsNullOrWhiteSpace(address))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            Message sentMessage = (response?.Value ?? new List<Message>())
                .Where(message => message.IsDraft != true)
                .OrderByDescending(message => message.SentDateTime)
                .FirstOrDefault(message =>
                {
                    HashSet<string> actualRecipients =
                        (message.ToRecipients ?? new List<Recipient>())
                            .Select(recipient =>
                                recipient.EmailAddress?.Address ?? "")
                            .Where(address =>
                                !string.IsNullOrWhiteSpace(address))
                            .ToHashSet(StringComparer.OrdinalIgnoreCase);
                    return expectedRecipients.SetEquals(actualRecipients);
                });

            if (sentMessage == null)
            {
                throw new InvalidOperationException(
                    "Sent opening account email not found for database archive retry" +
                    " - Queue ID : " + queueId +
                    " - Dossier : " + dossierId +
                    " - Subject : " + subject);
            }

            ArchiveSentOpeningMessageInDatabase(
                queueId,
                dossierId,
                mailbox,
                sentMessage,
                recipients);
        }

        private void ArchiveSentOpeningMessageInDatabase(
            int queueId,
            int dossierId,
            string mailbox,
            Message sentMessage,
            string recipients)
        {
            byte[] mime = GetMailboxMessageMimeContent(
                mailbox,
                sentMessage.Id,
                "Get sent opening account MIME");

            DateTime sentDate =
                sentMessage.SentDateTime?.LocalDateTime ??
                sentMessage.CreatedDateTime?.LocalDateTime ??
                DateTime.Now;

            string actualRecipients = string.Join(
                ";",
                (sentMessage.ToRecipients ?? new List<Recipient>())
                    .Select(recipient =>
                        recipient?.EmailAddress?.Address)
                    .Where(address =>
                        !string.IsNullOrWhiteSpace(address))
                    .Distinct(StringComparer.OrdinalIgnoreCase));

            if (string.IsNullOrWhiteSpace(actualRecipients))
            {
                throw new InvalidDataException(
                    "The sent opening account message has no Graph recipient" +
                    " - Queue ID : " + queueId +
                    " - Dossier : " + dossierId);
            }

            InsertSentOpeningMailInDatabase(
                dossierId,
                sentDate,
                sentMessage.Subject ?? "",
                actualRecipients,
                mime);

            UpdateOpeningUploadStatus(dossierId);

            WriteLog(
                "       Sent opening account email archived in database" +
                " - Queue ID : " + queueId +
                " - Dossier : " + dossierId +
                " - Subject : " + (sentMessage.Subject ?? "") +
                " - Recipient(s) : " + actualRecipients);

            // La copie reste volontairement dans les Elements envoyes.
            // La base contient le MIME pour l'interface web et Exchange
            // conserve la copie utilisateur du message envoye.
            WriteLog(
                "       Sent opening account email kept in Sent Items" +
                " - Queue ID : " + queueId +
                " - Dossier : " + dossierId +
                " - Message ID : " + sentMessage.Id);
        }

        private byte[] GetMailboxMessageMimeContent(
            string mailbox,
            string messageId,
            string operation)
        {
            using (Stream input = ExecuteGraphWithRetry(
                () => graphService.Users[mailbox]
                    .Messages[messageId]
                    .Content
                    .GetAsync(config => AddImmutableHeader(config.Headers))
                    .GetAwaiter()
                    .GetResult(),
                operation))
            using (var output = new MemoryStream())
            {
                if (input == null)
                {
                    throw new InvalidDataException(
                        "Graph MIME stream is empty" +
                        " - Mailbox : " + mailbox +
                        " - Message ID : " + messageId);
                }

                input.CopyTo(output);
                return output.ToArray();
            }
        }

        private void InsertSentOpeningMailInDatabase(
            int dossierId,
            DateTime sentDate,
            string subject,
            string recipients,
            byte[] mime)
        {
            if (mime == null || mime.Length == 0)
            {
                throw new InvalidDataException(
                    "Sent opening account MIME content is empty");
            }

            string fileName =
                dossierId +
                "_EMAIL_" +
                sentDate.ToString(
                    "yyyyMMdd_HHmmss",
                    CultureInfo.InvariantCulture);

            using (var connection = new SqlConnection(sql_creation_compte))
            using (var command = new SqlCommand(
                "dbo.USP_Insert_mail_boite_ouverture",
                connection))
            {
                command.CommandType = CommandType.StoredProcedure;
                command.CommandTimeout = 300;
                command.Parameters.Add("@id_dossier", SqlDbType.Int)
                    .Value = dossierId;
                command.Parameters.Add("@fichier", SqlDbType.Image)
                    .Value = mime;
                command.Parameters.Add("@nom_fichier", SqlDbType.VarChar, 500)
                    .Value = Truncate(fileName, 500);
                command.Parameters.Add("@extension", SqlDbType.NChar, 10)
                    .Value = "eml";
                command.Parameters.Add("@date_mail", SqlDbType.Date)
                    .Value = sentDate.Date;
                command.Parameters.Add("@sujet_mail", SqlDbType.VarChar, -1)
                    .Value = subject ?? "";
                command.Parameters.Add("@destinataire_mail", SqlDbType.VarChar, -1)
                    .Value = recipients ?? "";
                command.Parameters.Add("@date_importation", SqlDbType.Date)
                    .Value = DateTime.Today;
                connection.Open();
                command.ExecuteNonQuery();
            }
        }

        private void SendGraphMailWithAttachments(string sendAs, string to, string cc, string bcc, string subject, string html, List<string> files)
        {
            bool isCreditReviewSender = string.Equals(
                sendAs,
                fr_credit_review_graph_send_as,
                StringComparison.OrdinalIgnoreCase);
            bool isLcrnaSender = string.Equals(
                sendAs,
                fr_lcrna_graph_send_as,
                StringComparison.OrdinalIgnoreCase);
            bool isFacturationSender = string.Equals(
                sendAs,
                fr_facturation_graph_send_as,
                StringComparison.OrdinalIgnoreCase);

            string senderDisplayName = isCreditReviewSender
                ? "Ingram Micro - Service Analyse Crédit"
                : isLcrnaSender
                    ? "Ingram Micro - Service Finances"
                    : isFacturationSender
                        ? "Ingram Micro - Service Facturation"
                        : "";

            var sender = string.IsNullOrWhiteSpace(senderDisplayName)
                ? null
                : new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = sendAs,
                        Name = senderDisplayName
                    }
                };

            var message = new Message { Subject = subject, From = sender, Sender = sender, Body = new ItemBody { ContentType = BodyType.Html, Content = html }, ToRecipients = BuildRecipients(to), CcRecipients = BuildRecipients(cc), BccRecipients = BuildRecipients(bcc), Attachments = new List<Microsoft.Graph.Models.Attachment>() };
            if (message.ToRecipients.Count == 0) throw new InvalidOperationException("Invalid recipients for " + subject);
            foreach (string path in files ?? new List<string>()) { byte[] bytes = File.ReadAllBytes(path); if (bytes.Length > 3 * 1024 * 1024) throw new InvalidDataException("Attachment exceeds 3 MB : " + Path.GetFileName(path)); message.Attachments.Add(new FileAttachment { OdataType = "#microsoft.graph.fileAttachment", Name = Path.GetFileName(path), ContentType = "application/octet-stream", ContentBytes = bytes }); }
            ExecuteGraphWithRetry(() => { graphService.Users[sendAs].SendMail.PostAsync(new Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody { Message = message, SaveToSentItems = true }).GetAwaiter().GetResult(); return true; }, "Send template/report mail : " + subject);
        }

    }
}
