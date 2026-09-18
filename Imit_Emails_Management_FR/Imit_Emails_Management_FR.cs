using ClosedXML.Excel;
using ExcelDataReader;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using MimeKit;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;

namespace IMIT_EMAILS_MANAGEMENT_FR
{
    public class IMIT_EMAILS_MANAGEMENT_FR
    {
        private const string global_application_name = "IMIT_EMAILS_MANAGEMENT_FR";
        private const string imit_mailbox_fr = "IMIT_FR_4@ingrammicro.com";
        private const string imit_mailbox_es = "IMIT_ES@ingrammicro.com";
        private const string processedPropertyId = "String {B4A15BD8-83F0-4212-A932-4D9A14DF2889} Name IMCAProcessed";
        private string country = "";
        private string name = "";
        private string active = "";
        private string debug = "";
        private string start_date_scan = "";
        private string number_Of_Mails = "10";
        private string sharedmailbox_folder_in = "Inbox";
        private string sharedmailbox_folder_out = "Archives";
        private string sharedmailbox_folder_error = "Erreur";

        private string sql_connexion_parameter_global = "";
        private string sql_connexion = "";
        private string sql_con_transporteur_parameter_global = "";
        private string sql_con_transporteur = "";
        private string email_in_case_of_technical_issue_parameter_global = "";
        private string email_in_case_of_technical_issue = "";
        private string fedex_Sender = "";
        private string working_Folder = "";
        private string computacenter_Folder = "";
        private string output_Files_Folder = "";
        private string archive_Bcc_Mailbox = "";
        private string logsFolder = "";
        private string tempFolder = "";
        private string sessionName = "";
        private int mailboxId;
        private string mailboxDisplayName = "";
        private string mailboxAddress = "";
        private DateTime mailboxFilterDate = new DateTime(1900, 1, 1);
        private int refreshMinutes;
        private DateTime lastRefresh = new DateTime(1900, 1, 1);
        private bool automaticIntegration;
        private bool currentMailboxIsImit;
        private List<MailboxDomainRule> currentDomainRules = new List<MailboxDomainRule>();
        private GraphServiceClient graphService;

        static IMIT_EMAILS_MANAGEMENT_FR()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        public sealed class JsonFile
        {
            public List<Country> countries { get; set; } = new List<Country>();
        }

        public class Country
        {
            public string country { get; set; } = "";
            public string sk_valid { get; set; } = "";
            public string name { get; set; } = "";
            public string active { get; set; } = "TRUE";
            public string debug { get; set; } = "FALSE";
            public string start_date_scan { get; set; } = "";
            public string number_of_mails { get; set; } = "100";
            public string sharedmailbox_folder_in { get; set; } = "Inbox";
            public string sharedmailbox_folder_out { get; set; } = "Archives";
            public string sql_connexion_parameter_global { get; set; } = "";
            public string sql_con_transporteur_parameter_global { get; set; } = "";
            public string email_in_case_of_technical_issue_parameter_global { get; set; } = "";
            public string mail_fedex_service_client { get; set; } = "";
            public string path_macro_computacenter { get; set; } = "";
            public string path_fichier { get; set; } = "";
        }

        private sealed class MailboxConfiguration
        {
            public int Id { get; set; }
            public string Name { get; set; } = "";
            public string Address { get; set; } = "";
            public DateTime FilterDate { get; set; }
            public int RefreshMinutes { get; set; }
            public DateTime LastRefresh { get; set; }
            public bool AutomaticIntegration { get; set; }
            public List<MailboxDomainRule> DomainRules { get; set; } = new List<MailboxDomainRule>();
        }

        private sealed class MailboxDomainRule
        {
            public int SourceMailboxId { get; set; }
            public int DestinationMailboxId { get; set; }
            public string Domains { get; set; } = "";
        }

        private sealed class ForwardRequest
        {
            public int Id { get; set; }
            public string MessageId { get; set; } = "";
            public string SourceMailbox { get; set; } = "";
            public string To { get; set; } = "";
            public string Cc { get; set; } = "";
            public string From { get; set; } = "";
            public string FromName { get; set; } = "";
            public string Comment { get; set; } = "";
            public string Subject { get; set; } = "";
            public string StoredBody { get; set; } = "";
            public byte[] MimeContent { get; set; } = Array.Empty<byte>();
        }


        private sealed class TechnicalAlertAlreadySentException : Exception
        {
            public TechnicalAlertAlreadySentException(string message, Exception inner) : base(message, inner) { }
        }

        private sealed class FedexInvoice
        {
            public string Customer { get; set; } = "";
            public string CustomerName { get; set; } = "";
            public string Invoice { get; set; } = "";
            public DateTime InvoiceDate { get; set; }
            public string Suffix { get; set; } = "";
            public string CustomerPo { get; set; } = "";
            public string Order { get; set; } = "";
            public DateTime ShipDate { get; set; }
            public string CarrierCode { get; set; } = "";
            public string CarrierName { get; set; } = "";
            public string EntryMethod { get; set; } = "";
            public string Commercial { get; set; } = "";
        }

        public void Read_Email_with_Graph(string sql_con, string logs, string tmp_folder, string session_name)
        {
            string servicePath = GetServicePath();
            logsFolder = Path.Combine(servicePath, logs ?? "");
            tempFolder = Path.Combine(servicePath, tmp_folder ?? "");
            sessionName = session_name ?? "";

            try
            {
                string json = GetImcaParameter(sql_con, global_application_name);
                if (string.IsNullOrWhiteSpace(json))
                    throw new InvalidOperationException("No parameters found for " + global_application_name);

                JsonFile configuration = JsonConvert.DeserializeObject<JsonFile>(json);
                if (configuration?.countries == null || configuration.countries.Count == 0)
                    throw new InvalidOperationException(global_application_name + " parameters are empty or invalid");

                foreach (Country item in configuration.countries)
                {
                    ApplyCountryConfiguration(item, sql_con);
                    if (!IsTrue(active)) continue;

                    var countryErrors = new List<Exception>();
                    try
                    {
                        WriteLog(name.ToUpperInvariant() + "(" + country.ToUpperInvariant() + ") at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                        WriteLog("   Debug Parameter is set to " + debug.ToUpperInvariant());
                        ValidateCountryConfiguration();
                        Directory.CreateDirectory(tempFolder);
                        working_Folder = tempFolder;
                        WriteLog("   Working folder uses IMCA temp folder : " + working_Folder);
                        graphService = ConnectGraph();

                        List<MailboxConfiguration> mailboxes = GetActiveMailboxes();
                        if (mailboxes.Count == 0)
                        {
                            WriteLog("   No active shared mailbox found in T_SharedMailboxes");
                            continue;
                        }

                        MailboxConfiguration imitMailbox = GetCountryImitMailbox(mailboxes);
                        archive_Bcc_Mailbox = imitMailbox.Address;
                        WriteLog("   IMIT mailbox loaded from T_SharedMailboxes : ID " + imitMailbox.Id + " - " + archive_Bcc_Mailbox);
                        foreach (MailboxConfiguration current in mailboxes)
                        {
                            SetCurrentMailbox(current);
                            if (!IsRefreshDue(lastRefresh, refreshMinutes))
                            {
                                if (IsTrue(debug)) WriteLog("   Mailbox skipped, refresh is not due : " + mailboxAddress);
                                continue;
                            }

                            try
                            {
                                WriteLog("   Processing mailbox " + mailboxDisplayName + " : " + mailboxAddress);
                                ReadCurrentMailbox();
                                UpdateMailboxRefreshInformation(mailboxId);
                            }
                            catch (Exception ex)
                            {
                                WriteLog("   Error reading mailbox " + mailboxAddress + " : " + ex.Message);
                                countryErrors.Add(new Exception("Mailbox " + mailboxAddress + " : " + ex.Message, ex));
                                if (!(ex is TechnicalAlertAlreadySentException))
                                    SendTechnicalAlert(nameof(Read_Email_with_Graph), mailboxAddress + " - " + ex.Message, "MAILBOX PROCESSING");
                            }
                        }

                        try
                        {
                            ProcessPendingForwards();
                        }
                        catch (Exception ex)
                        {
                            countryErrors.Add(ex);
                        }

                        try { SendMonthlyBusinessUsageReportIfRequired(); }
                        catch (Exception ex)
                        {
                            WriteLog("   Monthly business usage report error : " + ex.Message);
                            countryErrors.Add(ex);
                        }

                        if (countryErrors.Count > 0)
                            throw new AggregateException(countryErrors.Count + " IMIT email processing error(s).", countryErrors);
                    }
                    finally
                    {
                        graphService = null;
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog("Global error Read_Email_with_Graph : " + ex.Message);
                throw;
            }
            finally
            {
                graphService = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private GraphServiceClient ConnectGraph()
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            return new class_dev_tools.Ews_Modern_Auth().Get_Graph_Service();
        }

        private void ApplyCountryConfiguration(Country item, string imcaConnection)
        {
            country = item.country ?? "";
            name = item.name ?? "";
            active = item.active ?? "";
            debug = item.debug ?? "";
            start_date_scan = item.start_date_scan ?? "";
            number_Of_Mails = string.IsNullOrWhiteSpace(item.number_of_mails) ? "100" : item.number_of_mails;
            sharedmailbox_folder_in = string.IsNullOrWhiteSpace(item.sharedmailbox_folder_in) ? "Inbox" : item.sharedmailbox_folder_in;
            sharedmailbox_folder_out = string.IsNullOrWhiteSpace(item.sharedmailbox_folder_out) ? "Archives" : item.sharedmailbox_folder_out;
            sql_connexion_parameter_global = item.sql_connexion_parameter_global ?? "";
            sql_con_transporteur_parameter_global = item.sql_con_transporteur_parameter_global ?? "";
            email_in_case_of_technical_issue_parameter_global = item.email_in_case_of_technical_issue_parameter_global ?? "";
            fedex_Sender = item.mail_fedex_service_client ?? "";
            working_Folder = tempFolder;
            computacenter_Folder = item.path_macro_computacenter ?? "";
            output_Files_Folder = item.path_fichier ?? "";
            sql_connexion = GetImcaParameter(imcaConnection, sql_connexion_parameter_global);
            email_in_case_of_technical_issue = GetImcaParameter(imcaConnection, email_in_case_of_technical_issue_parameter_global);
            sql_con_transporteur = string.IsNullOrWhiteSpace(sql_con_transporteur_parameter_global)
                ? ""
                : GetImcaParameter(imcaConnection, sql_con_transporteur_parameter_global);
        }

        private List<MailboxConfiguration> GetActiveMailboxes()
        {
            const string sql = @"
SELECT  m.id_mailboxe,
        m.nom_mailboxe,
        m.mailboxe,
        ISNULL(m.dt_heure_filtre,'19000101') AS dt_heure_filtre,
        ISNULL(m.raffraichissement_min,0) AS raffraichissement_min,
        ISNULL(m.date_dernier_raf,'19000101') AS date_dernier_raf,
        ISNULL(m.is_integration_auto,0) AS is_integration_auto,
        ISNULL(d.id_mailboxe_source,0) AS id_mailboxe_source,
        ISNULL(d.id_mailboxe_dest,0) AS id_mailboxe_dest,
        ISNULL(d.nom_domaine,'') AS nom_domaine
FROM dbo.T_SharedMailboxes AS m
LEFT JOIN dbo.T_SharedMailboxes_Domain AS d
    ON d.id_mailboxe_source=m.id_mailboxe
WHERE m.actif=1
ORDER BY m.raffraichissement_min, m.id_mailboxe;";

            var byId = new Dictionary<int, MailboxConfiguration>();
            using (var connection = new SqlConnection(sql_connexion))
            using (var command = new SqlCommand(sql, connection))
            {
                command.CommandTimeout = 300;
                connection.Open();
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        int id = Convert.ToInt32(reader["id_mailboxe"]);
                        if (!byId.TryGetValue(id, out MailboxConfiguration mailbox))
                        {
                            mailbox = new MailboxConfiguration
                            {
                                Id = id,
                                Name = Convert.ToString(reader["nom_mailboxe"]).Trim(),
                                Address = Convert.ToString(reader["mailboxe"]).Trim(),
                                FilterDate = Convert.ToDateTime(reader["dt_heure_filtre"]),
                                RefreshMinutes = Convert.ToInt32(reader["raffraichissement_min"]),
                                LastRefresh = Convert.ToDateTime(reader["date_dernier_raf"]),
                                AutomaticIntegration = Convert.ToBoolean(reader["is_integration_auto"])
                            };
                            byId.Add(id, mailbox);
                        }

                        int source = Convert.ToInt32(reader["id_mailboxe_source"]);
                        int destination = Convert.ToInt32(reader["id_mailboxe_dest"]);
                        string domains = Convert.ToString(reader["nom_domaine"]);
                        if (source > 0 && destination > 0 && !string.IsNullOrWhiteSpace(domains))
                        {
                            mailbox.DomainRules.Add(new MailboxDomainRule
                            {
                                SourceMailboxId = source,
                                DestinationMailboxId = destination,
                                Domains = domains
                            });
                        }
                    }
                }
            }
            return byId.Values.OrderBy(x => x.RefreshMinutes).ThenBy(x => x.Id).ToList();
        }

        private MailboxConfiguration GetCountryImitMailbox(IEnumerable<MailboxConfiguration> mailboxes)
        {
            string expectedMailbox = GetCountryImitMailboxAddress();
            List<MailboxConfiguration> matches = mailboxes
                .Where(x => string.Equals(x.Address?.Trim(), expectedMailbox, StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (matches.Count == 0)
                throw new InvalidOperationException("The active IMIT mailbox was not found in T_SharedMailboxes for country " + country + " : " + expectedMailbox);
            if (matches.Count > 1)
                throw new InvalidOperationException("The IMIT mailbox is duplicated in T_SharedMailboxes for country " + country + " : " + expectedMailbox);
            return matches[0];
        }

        private string GetCountryImitMailboxAddress()
        {
            if (country.Equals("FR", StringComparison.OrdinalIgnoreCase)) return imit_mailbox_fr;
            if (country.Equals("ES", StringComparison.OrdinalIgnoreCase)) return imit_mailbox_es;
            throw new InvalidOperationException("No IMIT mailbox is configured in code for country " + country);
        }

        private bool IsCurrentCountryImitMailbox()
        {
            return string.Equals(mailboxAddress?.Trim(), GetCountryImitMailboxAddress(), StringComparison.OrdinalIgnoreCase);
        }

        private void SetCurrentMailbox(MailboxConfiguration mailbox)
        {
            mailboxId = mailbox.Id;
            mailboxDisplayName = mailbox.Name;
            mailboxAddress = mailbox.Address;
            mailboxFilterDate = mailbox.FilterDate;
            refreshMinutes = mailbox.RefreshMinutes;
            lastRefresh = mailbox.LastRefresh;
            automaticIntegration = mailbox.AutomaticIntegration;
            currentMailboxIsImit = IsCurrentCountryImitMailbox();
            currentDomainRules = mailbox.DomainRules ?? new List<MailboxDomainRule>();
        }

        private void ReadCurrentMailbox()
        {
            ValidateMailboxConfiguration();
            int inserted = 0;
            int existing = 0;
            int routed = 0;
            MailFolder inputFolder = GetInputFolder();
            MailFolder errorFolder = null;
            bool errorFolderLookupAttempted = false;
            MessageCollectionResponse messages = GetMessages(inputFolder.Id);

            if (messages?.Value == null || messages.Value.Count == 0)
            {
                if (IsTrue(debug)) WriteLog("       No emails found");
                return;
            }

            foreach (Message summary in messages.Value)
            {
                try
                {
                    if (IsMessageAlreadyProcessed(summary))
                    {
                        continue;
                    }
                    if (IsSpanishErmaMailbox() && (summary.Subject ?? "").IndexOf("Notificación de RMA", StringComparison.OrdinalIgnoreCase) < 0)
                        continue;

                    Message email = GetCompleteMessage(summary.Id);
                    WriteLog("       Subject : " + (email.Subject ?? "<no subject>"));
                    byte[] mime = GetMimeContent(email.Id);
                    int targetMailboxId = ResolveDestinationMailbox(email);
                    if (targetMailboxId != mailboxId) routed++;

                    int emailDatabaseId = InsertEmailIfNew(email, mime, targetMailboxId, false);
                    if (emailDatabaseId == 0)
                    {
                        existing++;
                        WriteLog("       Email already present in T_contenu_email");
                        MarkMessageAsProcessed(email.Id);
                        if (IsCurrentCountryImitMailbox()) MarkAsRead(email.Id);
                        continue;
                    }

                    inserted++;
                    WriteLog("       Email inserted in T_contenu_email. ID : " + emailDatabaseId);
                    ProcessBusinessRules(email, emailDatabaseId);
                    MarkMessageAsProcessed(email.Id);
                    if (IsCurrentCountryImitMailbox())
                    {
                        MarkAsRead(email.Id);
                        WriteLog("       IMIT mailbox message marked as read");
                    }
                }
                catch (Exception ex)
                {
                    WriteLog("       Error processing email : " + ex.Message);
                    if (!(ex is TechnicalAlertAlreadySentException))
                        SendTechnicalAlert(nameof(ReadCurrentMailbox), mailboxAddress + " - Mail : " + (summary.Subject ?? "<no subject>") + " - " + ex.Message, "EMAIL PROCESSING");
                    if (!errorFolderLookupAttempted)
                    {
                        errorFolderLookupAttempted = true;
                        errorFolder = TryGetErrorFolderAfterMessageFailure();
                    }

                    if (errorFolder != null)
                    {
                        try
                        {
                            MarkAsRead(summary.Id);
                            MoveMessage(summary.Id, errorFolder.Id);
                            WriteLog("       Email moved to Erreur folder : " + errorFolder.DisplayName);
                        }
                        catch (Exception moveException)
                        {
                            string moveError = mailboxAddress + " - Mail : " +
                                (summary.Subject ?? "<no subject>") +
                                " - Unable to move the failed message to folder " + sharedmailbox_folder_error +
                                " : " + GetInnermostExceptionMessage(moveException);
                            WriteLog("       " + moveError);
                            SendTechnicalAlert(nameof(ReadCurrentMailbox), moveError, "ERROR FOLDER MOVE");
                        }
                    }
                    else
                    {
                        WriteLog("       Failed message was not moved because folder " +
                            sharedmailbox_folder_error + " is unavailable. Processing continues.");
                    }
                }
            }

            if (IsTrue(debug))
                WriteLog("       Inserted : " + inserted + " - Already present : " + existing + " - Routed : " + routed);
        }

        private MessageCollectionResponse GetMessages(string folderId)
        {
            int top;
            if (!int.TryParse(number_Of_Mails, out top) || top <= 0) top = 100;
            if (IsCurrentCountryImitMailbox() && top > 20) top = 20;
            DateTime date = mailboxFilterDate > new DateTime(1900, 1, 1) ? mailboxFilterDate : ParseStartDate();
            return graphService.Users[mailboxAddress].MailFolders[folderId].Messages.GetAsync(config =>
            {
                config.QueryParameters.Top = top;
                config.QueryParameters.Orderby = new[] { "receivedDateTime asc" };
                string filter = "receivedDateTime gt " + date.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ", CultureInfo.InvariantCulture);
                if (IsCurrentCountryImitMailbox()) filter += " and isRead eq false";
                config.QueryParameters.Filter = filter;
                config.QueryParameters.Select = new[] { "id", "subject", "receivedDateTime", "createdDateTime", "hasAttachments" };
                config.QueryParameters.Expand = new[]
                {
                    "singleValueExtendedProperties($filter=id eq '" + EscapeODataString(processedPropertyId) + "')"
                };
            }).GetAwaiter().GetResult();
        }

        private Message GetCompleteMessage(string messageId)
        {
            return graphService.Users[mailboxAddress].Messages[messageId].GetAsync(config =>
            {
                config.Headers.Add("Prefer", "outlook.body-content-type=\"text\"");
                config.QueryParameters.Select = new[]
                {
                    "id", "subject", "body", "bodyPreview", "from", "sender", "toRecipients", "ccRecipients",
                    "receivedDateTime", "createdDateTime", "conversationId", "internetMessageId", "hasAttachments", "webLink"
                };
            }).GetAwaiter().GetResult();
        }

        private byte[] GetMimeContent(string messageId)
        {
            using (Stream input = graphService.Users[mailboxAddress].Messages[messageId].Content.GetAsync().GetAwaiter().GetResult())
            using (var output = new MemoryStream())
            {
                input.CopyTo(output);
                return output.ToArray();
            }
        }

        private int InsertEmailIfNew(Message email, byte[] mime, int targetMailboxId, bool sentElement)
        {
            if (EmailExists(email.Id, targetMailboxId)) return 0;
            using (var connection = new SqlConnection(sql_connexion))
            using (var command = new SqlCommand("dbo.USP_ADD_EMAIL_IN_DB", connection))
            {
                command.CommandType = CommandType.StoredProcedure;
                command.CommandTimeout = 300;
                Add(command, "@id_comment", SqlDbType.Int, 0);
                Add(command, "@EwsID", SqlDbType.NVarChar, email.Id ?? "", 500);
                Add(command, "@id_boite_mail", SqlDbType.Int, targetMailboxId);
                Add(command, "@dt_time_received", SqlDbType.DateTime, email.ReceivedDateTime?.LocalDateTime ?? email.CreatedDateTime?.LocalDateTime ?? DateTime.Now);
                Add(command, "@body", SqlDbType.NVarChar, GetTextBody(email), -1);
                Add(command, "@from", SqlDbType.NVarChar, GetSender(email), 500);
                Add(command, "@sujet", SqlDbType.NVarChar, Truncate(email.Subject, 500), 500);
                Add(command, "@to", SqlDbType.NVarChar, Truncate(GetRecipients(email.ToRecipients), 2000), 2000);
                Add(command, "@cc", SqlDbType.NVarChar, Truncate(GetRecipients(email.CcRecipients), 2000), 2000);
                Add(command, "@new_to", SqlDbType.NVarChar, "", 2000);
                Add(command, "@new_cc", SqlDbType.NVarChar, "", 2000);
                Add(command, "@new_from", SqlDbType.NVarChar, "", 500);
                Add(command, "@new_from_nom", SqlDbType.NVarChar, "", 500);
                Add(command, "@new_extra_body", SqlDbType.NVarChar, "", -1);
                Add(command, "@top_a_envoyer", SqlDbType.NVarChar, "N", 1);
                Add(command, "@nom_fic", SqlDbType.NVarChar, Truncate(CleanFileName(string.IsNullOrWhiteSpace(email.Subject) ? "email" : email.Subject).Replace(",", ""), 500), 500);
                Add(command, "@document", SqlDbType.Image, mime ?? Array.Empty<byte>());
                Add(command, "@extension", SqlDbType.NVarChar, "eml", 5);
                Add(command, "@affecte_a", SqlDbType.NVarChar, DBNull.Value, 30);
                Add(command, "@sent_element", SqlDbType.Bit, sentElement);
                SqlParameter output = command.Parameters.Add("@@id", SqlDbType.Int);
                output.Direction = ParameterDirection.Output;
                connection.Open();
                command.ExecuteNonQuery();
                int insertedId = output.Value == DBNull.Value ? 0 : Convert.ToInt32(output.Value);
                if (insertedId > 0) UpdateInsertedEmailMailbox(insertedId, mailboxAddress);
                return insertedId;
            }
        }

        private void UpdateInsertedEmailMailbox(int emailId, string sourceMailbox)
        {
            const string sql = @"UPDATE dbo.T_contenu_email SET boite_mail=@MAILBOX WHERE id=@ID;";
            ExecuteNonQuery(sql,
                new SqlParameter("@MAILBOX", SqlDbType.VarChar, 500) { Value = sourceMailbox ?? "" },
                new SqlParameter("@ID", SqlDbType.Int) { Value = emailId });
        }

        private bool EmailExists(string graphId, int targetMailboxId)
        {
            const string sql = @"SELECT TOP (1) id FROM dbo.T_contenu_email
WHERE EwsID COLLATE Latin1_General_CS_AS=@EWSID AND id_mailboxe=@MAILBOX;";
            using (var connection = new SqlConnection(sql_connexion))
            using (var command = new SqlCommand(sql, connection))
            {
                Add(command, "@EWSID", SqlDbType.VarChar, graphId ?? "", 500);
                Add(command, "@MAILBOX", SqlDbType.Int, targetMailboxId);
                connection.Open();
                return command.ExecuteScalar() != null;
            }
        }

        private int ResolveDestinationMailbox(Message email)
        {
            var addresses = new List<string> { GetSender(email) };
            addresses.AddRange(GetRecipientAddresses(email.ToRecipients));
            addresses.AddRange(GetRecipientAddresses(email.CcRecipients));
            foreach (MailboxDomainRule rule in currentDomainRules)
            {
                if (rule.SourceMailboxId != mailboxId || rule.DestinationMailboxId <= 0) continue;
                var domains = rule.Domains.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(NormalizeDomain).Where(x => x.Length > 0).ToHashSet(StringComparer.OrdinalIgnoreCase);
                if (addresses.Select(GetDomain).Any(domains.Contains)) return rule.DestinationMailboxId;
            }
            return mailboxId;
        }

        private void ProcessBusinessRules(Message email, int emailDatabaseId)
        {
            string subject = (email.Subject ?? "").Trim();
            string sender = GetSender(email);

            // Historical generic linking applies only to the IMIT mailbox of the current country.
            if (IsCurrentCountryImitMailbox())
                ProcessImitCountryMailboxIssueLink(email, emailDatabaseId);

            if (country.Equals("FR", StringComparison.OrdinalIgnoreCase) && mailboxId == 4)
            {
                if (subject.IndexOf("dossier(s) imit en blocage ", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    subject.IndexOf("tr:", StringComparison.OrdinalIgnoreCase) < 0 &&
                    subject.IndexOf("fw:", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    string file = DownloadAttachmentByPrefix(email.Id, "Dossiers_IMIT_blocage_interne", working_Folder);
                    if (!string.IsNullOrWhiteSpace(file))
                    {
                        TrackBusinessMethod("ProcessInternalBlockage", subject, () => ProcessInternalBlockage(file, emailDatabaseId, subject));
                        ArchiveEmailInDatabase(emailDatabaseId);
                    }
                }

                if (subject.IndexOf("discounts of the day", StringComparison.OrdinalIgnoreCase) >= 0 && sender.Equals("pline_imit@ingrammicro.fr", StringComparison.OrdinalIgnoreCase))
                {
                    string file = DownloadAttachmentByPrefix(email.Id, "Escompte", working_Folder);
                    if (!string.IsNullOrWhiteSpace(file)) TrackBusinessMethod("ProcessCashDiscount", subject, () => ProcessCashDiscount(file, emailDatabaseId));
                    ArchiveEmailInDatabase(emailDatabaseId);
                }

                if (!string.IsNullOrWhiteSpace(fedex_Sender) && sender.Equals(fedex_Sender, StringComparison.OrdinalIgnoreCase))
                {
                    bool processed = TrackBusinessMethod("ProcessFedex", subject, () => ProcessFedex(email, emailDatabaseId));
                    UpdateConversationInformation(email.InternetMessageId ?? "", GetFedexKey(subject), emailDatabaseId);
                    if (processed) ArchiveEmailInDatabase(emailDatabaseId);
                }
            }

            if (country.Equals("FR", StringComparison.OrdinalIgnoreCase) && mailboxId == 1)
                ProcessTransportRules(email, emailDatabaseId);

            if (country.Equals("FR", StringComparison.OrdinalIgnoreCase) && mailboxId == 5 &&
                subject.StartsWith("CUS001_HPSTORE_FR_Temp vom", StringComparison.OrdinalIgnoreCase) &&
                sender.Equals("JOBDISP@ingrammicro.com", StringComparison.OrdinalIgnoreCase))
            {
                string file = DownloadAttachmentByPrefix(email.Id, "CUS001_HPSTORE_FR_Temp", working_Folder);
                if (!string.IsNullOrWhiteSpace(file)) TrackBusinessMethod("ProcessClaimId", subject, () => ProcessClaimId(PrepareClaimIdFile(file), subject));
            }

            if (automaticIntegration &&
                (country.Equals("FR", StringComparison.OrdinalIgnoreCase) ||
                 (country.Equals("ES", StringComparison.OrdinalIgnoreCase) && subject.StartsWith("Notificación de RMA", StringComparison.OrdinalIgnoreCase))))
            {
                TrackBusinessMethod("ProcessAutomaticErma_" + country.ToUpperInvariant(), subject,
                    () => ProcessAutomaticErma(email, emailDatabaseId));
            }
        }

        private bool ProcessImitCountryMailboxIssueLink(Message email, int emailDatabaseId)
        {
            int issueId = ExtractImitIssueId(email?.Subject);
            if (issueId <= 0) return false;

            if (!ImitIssueExists(issueId))
            {
                WriteLog("       IMIT issue ID extracted from subject but not found in T_issue : " + issueId);
                return false;
            }

            if (EmailCommentLinkExists(emailDatabaseId, issueId))
            {
                WriteLog("       Email database ID " + emailDatabaseId + " is already linked to IMIT issue " + issueId);
                return true;
            }

            int commentId = InsertComment(
                issueId,
                string.IsNullOrWhiteSpace(email?.Subject) ? "Echange de mail" : email.Subject,
                "robot_mail",
                emailDatabaseId,
                email?.WebLink);

            WriteLog("       Email linked to IMIT issue " + issueId + " with comment ID " + commentId);
            return true;
        }

        private static int ExtractImitIssueId(string subject)
        {
            if (string.IsNullOrWhiteSpace(subject)) return 0;

            Match match = Regex.Match(
                subject,
                @"N[°ºo]?\s*de\s*dossier\s*INGRAM\s*:\s*(\d+)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

            if (!match.Success) return 0;
            return int.TryParse(match.Groups[1].Value, NumberStyles.Integer,
                CultureInfo.InvariantCulture, out int issueId) ? issueId : 0;
        }

        private bool ImitIssueExists(int issueId)
        {
            const string sql = @"SELECT TOP (1) id_issue FROM dbo.T_issue WHERE id_issue=@ISSUE_ID;";
            using (var connection = new SqlConnection(sql_connexion))
            using (var command = new SqlCommand(sql, connection))
            {
                command.CommandTimeout = 300;
                command.Parameters.Add("@ISSUE_ID", SqlDbType.Int).Value = issueId;
                connection.Open();
                return command.ExecuteScalar() != null;
            }
        }

        private bool EmailCommentLinkExists(int emailDatabaseId, int issueId)
        {
            const string sql = @"SELECT TOP (1) id_comment
FROM dbo.T_comment
WHERE id_issue=@ISSUE_ID
  AND id_contenu_email=@EMAIL_ID;";
            using (var connection = new SqlConnection(sql_connexion))
            using (var command = new SqlCommand(sql, connection))
            {
                command.CommandTimeout = 300;
                command.Parameters.Add("@ISSUE_ID", SqlDbType.Int).Value = issueId;
                command.Parameters.Add("@EMAIL_ID", SqlDbType.Int).Value = emailDatabaseId;
                connection.Open();
                return command.ExecuteScalar() != null;
            }
        }

        private void ProcessTransportRules(Message email, int emailDatabaseId)
        {
            string subject = (email.Subject ?? "").Trim();
            string sender = GetSender(email);
            if (subject.Equals("suivi computa center", StringComparison.OrdinalIgnoreCase))
            {
                string file = DownloadAttachmentByPrefix(email.Id, "suivi computacenter", computacenter_Folder);
                if (!string.IsNullOrWhiteSpace(file)) TrackBusinessMethod("ProcessComputacenter", subject, () => ProcessComputacenter(file, emailDatabaseId));
                return;
            }
            if (subject.Equals("suivi computacenter du", StringComparison.OrdinalIgnoreCase))
            {
                ArchiveEmailInDatabase(emailDatabaseId);
                return;
            }
            if (subject.Equals("dhl intraship - shipment history", StringComparison.OrdinalIgnoreCase) && sender.Equals("intraship@dhl.com", StringComparison.OrdinalIgnoreCase))
            {
                string file = DownloadAttachmentByPrefix(email.Id, "shipments", working_Folder);
                if (!string.IsNullOrWhiteSpace(file) && TrackBusinessMethod("ImportTrackingFile_DHL", subject, () => ImportTrackingFile(1, "T_Tracking_Auto_DHL", file))) ArchiveEmailInDatabase(emailDatabaseId);
                return;
            }
            if (subject.Equals("tracking export calberson", StringComparison.OrdinalIgnoreCase) && sender.Equals("transport@ingrammicro.fr", StringComparison.OrdinalIgnoreCase))
            {
                string file = DownloadAttachmentByPrefix(email.Id, "DownloadSuiviExpeditions", working_Folder);
                if (!string.IsNullOrWhiteSpace(file) && TrackBusinessMethod("ImportTrackingFile_CALBERSON", subject, () => ImportTrackingFile(2, "T_Tracking_Auto_CALBERSON", file))) ArchiveEmailInDatabase(emailDatabaseId);
                return;
            }
            if (subject.Equals("tracking export chronopost", StringComparison.OrdinalIgnoreCase) && sender.Equals("transport@ingrammicro.fr", StringComparison.OrdinalIgnoreCase))
            {
                string file = DownloadAttachmentByPrefix(email.Id, "bordereau", working_Folder);
                if (!string.IsNullOrWhiteSpace(file) && TrackBusinessMethod("ImportTrackingFile_CHRONOPOST", subject, () => ImportTrackingFile(3, "T_Tracking_Auto_CHRONOPOST", file))) ArchiveEmailInDatabase(emailDatabaseId);
                return;
            }

            string[] automaticArchiveSenders =
            {
                "robot.it@ingrammicro.com", "jobdisp@ingrammicro.com", "retour.information@fedex.com",
                "te.prod@tatexpress.fr", "chrexp@ediserv.chronopost.fr"
            };
            if (automaticArchiveSenders.Contains(sender, StringComparer.OrdinalIgnoreCase) ||
                (subject.IndexOf("ticket", StringComparison.OrdinalIgnoreCase) >= 0 && subject.IndexOf("en cours de traitement", StringComparison.OrdinalIgnoreCase) >= 0 && sender.Equals("no-reply@chronopost.fr", StringComparison.OrdinalIgnoreCase)) ||
                (subject.IndexOf("suivi recto verso", StringComparison.OrdinalIgnoreCase) >= 0 && sender.Equals("frd_scr_lille@fedex.com", StringComparison.OrdinalIgnoreCase)))
                ArchiveEmailInDatabase(emailDatabaseId);
        }

        private void ProcessAutomaticErma(Message email, int emailDatabaseId)
        {
            string body = email.Body?.Content ?? "";
            DateTime received = email.ReceivedDateTime?.LocalDateTime ?? DateTime.Now;
            int issueId = country.Equals("ES", StringComparison.OrdinalIgnoreCase)
                ? FindSpanishErmaIssue(body, received)
                : FindFrenchErmaIssue(body, received);
            if (issueId <= 0) return;
            ArchiveEmailInDatabase(emailDatabaseId);
            InsertComment(issueId, "Erma Confirmation", "robot_mail", emailDatabaseId);
        }

        private int FindFrenchErmaIssue(string body, DateTime received)
        {
            string customer = FindBodyValue(body, "compte client").Replace("-", "");
            string reference = FindBodyValue(body, "numéro de dossier revendeur", "numero de dossier revendeur").Replace("-", "");
            string status = FindBodyValue(body, "statut du rma");
            string rma = FindBodyValue(body, "numéro de rma", "numero de rma").Replace("-", "");
            return FindErmaIssue(customer, reference, status, rma, received, false);
        }

        private int FindSpanishErmaIssue(string body, DateTime received)
        {
            string customer = FindBodyValue(body, "código de cliente", "codigo de cliente").Replace("-", "");
            string reference = FindBodyValue(body, "referencia del cliente").Replace("-", "");
            string status = FindBodyValue(body, "estado del rma");
            string rma = FindBodyValue(body, "número de rma", "numero de rma").Replace("-", "");
            return FindErmaIssue(customer, reference, status, rma, received, true);
        }

        private int FindErmaIssue(string customer, string reference, string status, string rma, DateTime received, bool spanish)
        {
            if (string.IsNullOrWhiteSpace(customer) || string.IsNullOrWhiteSpace(reference) || string.IsNullOrWhiteSpace(status)) return 0;
            bool accepted = status.Equals(spanish ? "ACEPTADO" : "ACCEPTE", StringComparison.OrdinalIgnoreCase);
            bool rejected = status.Equals(spanish ? "NO ACEPTADO" : "REJETE", StringComparison.OrdinalIgnoreCase);
            if (!accepted && !rejected) return 0;
            string history = spanish ? "Generation ERMA request" : "Demande de génération du ERMA";
            string baseSql = @"
SELECT DISTINCT i.id_issue
FROM dbo.T_issue i
INNER JOIN dbo.T_RMA r ON r.id_issue=i.id_issue
INNER JOIN dbo.T_histo h ON h.id_issue=i.id_issue
LEFT JOIN (SELECT DISTINCT id_issue FROM dbo.T_comment WHERE comment_name='ERMA') c ON c.id_issue=i.id_issue
WHERE " + (spanish ? "i.erma=1" : "(i.erma=1 OR c.id_issue IS NOT NULL)") + @"
AND i.br_cut_nbr=@CUSTOMER
AND i.refuser=@REFUSED
AND h.comment=@HISTORY
AND CONVERT(varchar(8),h.histo_date,112)=@RECEIVED";

            List<int> first = ExecuteErmaSearch(baseSql + (accepted ? " AND r.rma_nbr=@RMA" : " AND REPLACE(REPLACE(i.num_litige_client,'-',''),'/','')=@REFERENCE"), customer, reference, rma, history, received, rejected);
            if (first.Count == 1) return first[0];
            if (first.Count > 1 && accepted)
            {
                List<int> qualified = ExecuteErmaSearch(baseSql + " AND r.rma_nbr=@RMA AND REPLACE(REPLACE(i.num_litige_client,'-',''),'/','')=@REFERENCE", customer, reference, rma, history, received, rejected);
                return qualified.Count == 1 ? qualified[0] : 0;
            }
            return 0;
        }

        private List<int> ExecuteErmaSearch(string sql, string customer, string reference, string rma, string history, DateTime received, bool rejected)
        {
            var ids = new List<int>();
            using (var connection = new SqlConnection(sql_connexion))
            using (var command = new SqlCommand(sql, connection))
            {
                command.CommandTimeout = 300;
                Add(command, "@CUSTOMER", SqlDbType.VarChar, customer, 50);
                Add(command, "@REFUSED", SqlDbType.Bit, rejected);
                Add(command, "@RMA", SqlDbType.VarChar, rma ?? "", 50);
                Add(command, "@REFERENCE", SqlDbType.VarChar, (reference ?? "").Replace("/", ""), 500);
                Add(command, "@HISTORY", SqlDbType.VarChar, history, 255);
                Add(command, "@RECEIVED", SqlDbType.Char, received.ToString("yyyyMMdd"), 8);
                connection.Open();
                using (SqlDataReader reader = command.ExecuteReader()) while (reader.Read()) ids.Add(Convert.ToInt32(reader[0]));
            }
            return ids;
        }

        private bool ProcessFedex(Message email, int emailDatabaseId)
        {
            string subject = email.Subject ?? "";
            string body = email.Body?.Content ?? "";
            string[] subjectParts = subject.Split('/');
            if (subjectParts.Length < 5 || subject.IndexOf("0085537", StringComparison.OrdinalIgnoreCase) < 0) return false;
            string parcel = subjectParts[1].Trim();
            string sheet = subjectParts[4].Trim();

            int existingIssue = FindExistingFedexIssue(sheet + "#" + parcel, emailDatabaseId);
            if (existingIssue > 0)
            {
                InsertComment(existingIssue, "Echange de mail", "robot_mail", emailDatabaseId);
                return true;
            }

            if (!body.TrimStart().StartsWith("cher client,", StringComparison.OrdinalIgnoreCase) ||
                body.IndexOf("APPEL NON MOTIVE", StringComparison.OrdinalIgnoreCase) < 0 ||
                body.IndexOf("=================================================================", StringComparison.Ordinal) < 0)
                return false;

            string postalCode = FindSecondValue(body, "cp :", 5);
            string customerPo = FindSectionValue(body, "ref3 :", "DEMANDE").Replace("=", "").Trim();
            string deliveryOrder = FindFixedValue(body, "ref1 :", 12).Replace("-", "").Trim();
            string customerShort = FindFixedValue(body, "ref2 :", 6).Trim();
            FedexInvoice invoice = FindFedexInvoice(parcel, postalCode, customerPo, deliveryOrder, customerShort);
            if (invoice == null || string.IsNullOrWhiteSpace(invoice.Invoice)) return false;

            int issueId = InsertIssue();
            UpdateFedexIssue(issueId, invoice);
            InsertFedexSkus(issueId, invoice);
            InsertComment(issueId, "demande commercial", "robot_mail", emailDatabaseId);
            ExecuteNonQuery("INSERT INTO dbo.T_DMS(id_issue,job_status) VALUES(@ID,1); INSERT INTO dbo.T_DMS(id_issue,job_status) VALUES(@ID,9);", new SqlParameter("@ID", SqlDbType.Int) { Value = issueId });
            ExecuteNonQuery("INSERT INTO dbo.T_RMA(id_issue,rma_nbr,rma,rma_amount,cn,a_faire) VALUES(@ID,'','',0,'',0);", new SqlParameter("@ID", SqlDbType.Int) { Value = issueId });
            InsertHistory(issueId, 10, "", "macro_imit");
            InsertHistory(issueId, 20, "Le dossier est passé en statut \"En attente de traitement\"", "macro_imit");
            InsertHistory(issueId, 30, "Le dossier est passé en statut \"En attente de réponse\"", "macro_imit");
            InsertHistory(issueId, 80, "Le dossier est passé en statut \"Archivé\" via le bouton [Refuser]", "macro_imit");
            return true;
        }

        private FedexInvoice FindFedexInvoice(string parcel, string postalCode, string customerPo, string deliveryOrder, string customerShort)
        {
            string customer = "";
            string invoice = "";
            if (!string.IsNullOrWhiteSpace(sql_con_transporteur))
            {
                const string trackingSql = @"SELECT TOP (1) p.br+p.cust_account customer, s.invoice_nbr
FROM dbo.T_IM_SO s INNER JOIN dbo.T_CUST_PO p ON s.id_po=p.id_po
INNER JOIN dbo.T_LAST_STATE l ON s.id_so=l.id_so
WHERE l.tracking LIKE @TRACKING AND (@PO='' OR p.cust_po=@PO);";
                using (var connection = new SqlConnection(sql_con_transporteur))
                using (var command = new SqlCommand(trackingSql, connection))
                {
                    Add(command, "@TRACKING", SqlDbType.VarChar, "%" + parcel + postalCode + "%", 255);
                    Add(command, "@PO", SqlDbType.VarChar, customerPo ?? "", 255);
                    connection.Open();
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            customer = Convert.ToString(reader["customer"]);
                            invoice = Convert.ToString(reader["invoice_nbr"]);
                        }
                    }
                }
            }

            const string invoiceSql = @"SELECT TOP (1) sh.invoice_nbr,sh.ship_to_suffix,sh.ShipmentDt,sh.carrier_code_at_ship,sh.Ship_Via,
sh.entry_method,LOWER(e.employee_name) employee_name,sh.branch_customer_nbr,c.cust_name,
sh.order_nbr+sh.DistCd+sh.ShipCd bl,sh.CustPoNbr,sh.invoice_date
FROM dbo.shipment_header sh INNER JOIN dbo.customer c ON sh.branch_customer_nbr=c.branch_customer_nbr
INNER JOIN dbo.employee e ON e.employee=c.assigned_is
WHERE ((@INVOICE<>'' AND sh.branch_customer_nbr=@CUSTOMER AND sh.invoice_nbr=@INVOICE)
OR (@INVOICE='' AND RIGHT(sh.branch_customer_nbr,6)=@SHORT AND sh.CustPoNbr=@PO AND sh.order_nbr+sh.DistCd+sh.ShipCd=@ORDER))
ORDER BY sh.invoice_date DESC;";
            using (var connection = new SqlConnection(sql_connexion))
            using (var command = new SqlCommand(invoiceSql, connection))
            {
                Add(command, "@INVOICE", SqlDbType.VarChar, invoice ?? "", 20);
                Add(command, "@CUSTOMER", SqlDbType.VarChar, customer ?? "", 20);
                Add(command, "@SHORT", SqlDbType.VarChar, customerShort ?? "", 10);
                Add(command, "@PO", SqlDbType.VarChar, customerPo ?? "", 50);
                Add(command, "@ORDER", SqlDbType.VarChar, deliveryOrder ?? "", 20);
                connection.Open();
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    if (!reader.Read()) return null;
                    return new FedexInvoice
                    {
                        Customer = Convert.ToString(reader["branch_customer_nbr"]),
                        CustomerName = Convert.ToString(reader["cust_name"]),
                        Invoice = Convert.ToString(reader["invoice_nbr"]),
                        InvoiceDate = Convert.ToDateTime(reader["invoice_date"]),
                        Suffix = Convert.ToString(reader["ship_to_suffix"]),
                        CustomerPo = Convert.ToString(reader["CustPoNbr"]),
                        Order = Convert.ToString(reader["bl"]),
                        ShipDate = Convert.ToDateTime(reader["ShipmentDt"]),
                        CarrierCode = Convert.ToString(reader["carrier_code_at_ship"]),
                        CarrierName = "FEDEX",
                        EntryMethod = Convert.ToString(reader["entry_method"]),
                        Commercial = Convert.ToString(reader["employee_name"])
                    };
                }
            }
        }

        private int InsertIssue()
        {
            using (var connection = new SqlConnection(sql_connexion))
            using (var command = new SqlCommand("dbo.USP_Insert_Issue", connection))
            {
                command.CommandType = CommandType.StoredProcedure;
                SqlParameter output = command.Parameters.Add("@@num", SqlDbType.Int);
                output.Direction = ParameterDirection.Output;
                connection.Open(); command.ExecuteNonQuery();
                return Convert.ToInt32(output.Value);
            }
        }

        private void UpdateFedexIssue(int issueId, FedexInvoice invoice)
        {
            using (var connection = new SqlConnection(sql_connexion))
            using (var command = new SqlCommand("dbo.USP_Update_Issue_9", connection))
            {
                command.CommandType = CommandType.StoredProcedure; command.CommandTimeout = 300;
                Add(command, "@c_id_issue", SqlDbType.Int, issueId);
                Add(command, "@c_br_cut_nbr", SqlDbType.Char, invoice.Customer, 10);
                Add(command, "@c_ship_to_suffix", SqlDbType.Char, invoice.Suffix, 3);
                Add(command, "@c_cust_name", SqlDbType.VarChar, invoice.CustomerName, 35);
                Add(command, "@c_order_nbr", SqlDbType.Char, invoice.Order, 9);
                Add(command, "@c_invoice_nbr", SqlDbType.Char, invoice.Invoice, 9);
                Add(command, "@c_invoice_date", SqlDbType.SmallDateTime, invoice.InvoiceDate);
                Add(command, "@c_ship_date", SqlDbType.SmallDateTime, invoice.ShipDate);
                Add(command, "@c_carrier_code", SqlDbType.Char, invoice.CarrierCode, 2);
                Add(command, "@c_carrier_name", SqlDbType.VarChar, "FEDEX", 35);
                Add(command, "@c_cust_po", SqlDbType.VarChar, invoice.CustomerPo, 18);
                Add(command, "@c_id_status", SqlDbType.Int, 80);
                Add(command, "@c_reason_code", SqlDbType.VarChar, "TR", 4);
                Add(command, "@c_owner", SqlDbType.VarChar, GetOwner(invoice.Customer), 50);
                Add(command, "@c_owner_ssc", SqlDbType.VarChar, GetOwnerSsc(invoice.Customer), 50);
                Add(command, "@c_assigned_to", SqlDbType.VarChar, "", 50);
                Add(command, "@c_creation_date", SqlDbType.SmallDateTime, DateTime.Now);
                Add(command, "@c_resolved_date", SqlDbType.SmallDateTime, new DateTime(1900, 1, 1));
                Add(command, "@c_issue_description", SqlDbType.VarChar, "", 500);
                Add(command, "@c_refuser", SqlDbType.Bit, true);
                Add(command, "@c_num_litige_client", SqlDbType.VarChar, "", -1);
                Add(command, "@c_note_debit", SqlDbType.VarChar, "", 50);
                Add(command, "@c_montant_note_debit", SqlDbType.Float, 0d);
                Add(command, "@c_decote", SqlDbType.Float, -99999d);
                Add(command, "@c_sous_reason_code", SqlDbType.VarChar, "", 2);
                Add(command, "@c_commercial", SqlDbType.VarChar, invoice.Commercial, 50);
                Add(command, "@c_num_rma_constructeur", SqlDbType.VarChar, "", 20);
                Add(command, "@c_num_avoir_constructeur", SqlDbType.VarChar, "", -1);
                Add(command, "@c_blocage", SqlDbType.VarChar, "", 30);
                Add(command, "@c_type_encodage", SqlDbType.VarChar, invoice.EntryMethod, 5);
                Add(command, "@c_num_frt", SqlDbType.VarChar, "", 15);
                Add(command, "@c_montant_avoir", SqlDbType.Float, 0d);
                Add(command, "@c_montant_frt", SqlDbType.Float, 0d);
                Add(command, "@c_num_fact_client", SqlDbType.VarChar, "", 20);
                Add(command, "@c_montant_fact_client", SqlDbType.Float, 0d);
                Add(command, "@c_num_fact_mkt", SqlDbType.VarChar, "", 10);
                Add(command, "@c_montant_fact_mkt", SqlDbType.Float, 0d);
                Add(command, "@c_montant_fact_mkt_paye", SqlDbType.Float, 0d);
                Add(command, "@c_deduc_faite", SqlDbType.VarChar, "Non constaté", 13);
                Add(command, "@c_compensation_demande", SqlDbType.NVarChar, "", 3);
                Add(command, "@c_remboursement_client", SqlDbType.NVarChar, "", 50);
                Add(command, "@c_fournisseur", SqlDbType.VarChar, "", 30);
                Add(command, "@c_date_fin_op", SqlDbType.VarChar, "01/01/1900", 10);
                Add(command, "@c_collecte", SqlDbType.Bit, false);
                Add(command, "@c_num_dic", SqlDbType.VarChar, "", 25);
                Add(command, "@c_reception_avoir_frs", SqlDbType.VarChar, "N", 1);
                Add(command, "@c_num_accord_fact_transporteur", SqlDbType.VarChar, "", 20);
                Add(command, "@c_num_fact_IM_transporteur", SqlDbType.VarChar, "", 9);
                Add(command, "@c_FFR", SqlDbType.VarChar, "N", 1);
                Add(command, "@c_num_litige_transporteur", SqlDbType.VarChar, "", 50);
                Add(command, "@c_validation_ISC", SqlDbType.VarChar, "A", 1);
                connection.Open(); command.ExecuteNonQuery();
            }
        }

        private void InsertFedexSkus(int issueId, FedexInvoice invoice)
        {
            const string sql = @"SELECT TOP (1) sl.line_nbr,p.groupnam marque,sl.sku,sl.mfr_part_nbr,sl.product_descr,
sl.qty_shipped,sl.extended_sales,sl.cost,sl.rpl,sl.rbl,p.buyer_name,ISNULL(sl.rb2,0) rb2,ISNULL(sl.rb3,0) rb3
FROM dbo.shipment_line sl LEFT JOIN dbo.product p ON sl.sku=p.sku
WHERE sl.invoice_nbr=@INVOICE AND sl.branch_customer_nbr=@CUSTOMER
AND CONVERT(varchar(8),sl.invoice_date,112)=@DATE AND sl.line_nbr<>0 ORDER BY sl.qty_shipped DESC;";
            using (var connection = new SqlConnection(sql_connexion))
            using (var command = new SqlCommand(sql, connection))
            {
                Add(command, "@INVOICE", SqlDbType.VarChar, invoice.Invoice, 20);
                Add(command, "@CUSTOMER", SqlDbType.VarChar, invoice.Customer, 20);
                Add(command, "@DATE", SqlDbType.Char, invoice.InvoiceDate.ToString("yyyyMMdd"), 8);
                connection.Open();
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read()) InsertFedexSku(issueId, invoice, reader);
                }
            }
        }

        private void InsertFedexSku(int issueId, FedexInvoice invoice, IDataRecord row)
        {
            int quantity = ToInt(row["qty_shipped"]);
            double sales = ToDouble(row["extended_sales"]);
            double cost = ToDouble(row["cost"]);
            int skuId;
            using (var connection = new SqlConnection(sql_connexion))
            using (var command = new SqlCommand("dbo.USP_Insert_Sku_4", connection))
            {
                command.CommandType = CommandType.StoredProcedure;
                Add(command, "@sku", SqlDbType.Char, Convert.ToString(row["sku"]), 7);
                Add(command, "@so_line_nbr", SqlDbType.Int, ToInt(row["line_nbr"]));
                Add(command, "@id_issue", SqlDbType.Int, issueId);
                Add(command, "@acheteur", SqlDbType.VarChar, Convert.ToString(row["buyer_name"]), 50);
                Add(command, "@qty_shipped", SqlDbType.Int, quantity);
                Add(command, "@qty_concerned", SqlDbType.Int, quantity);
                Add(command, "@mfr_part_nbr", SqlDbType.VarChar, Convert.ToString(row["mfr_part_nbr"]), 35);
                Add(command, "@description", SqlDbType.VarChar, Convert.ToString(row["product_descr"]), 100);
                Add(command, "@marque", SqlDbType.VarChar, Convert.ToString(row["marque"]), 35);
                Add(command, "@pu_facture", SqlDbType.Float, sales);
                Add(command, "@avg_total", SqlDbType.Float, cost);
                Add(command, "@rpl_u", SqlDbType.Float, ToDouble(row["rpl"]));
                Add(command, "@rbl_u", SqlDbType.Float, ToDouble(row["rbl"]));
                Add(command, "@rb2_u", SqlDbType.Float, ToDouble(row["rb2"]));
                Add(command, "@rb3_u", SqlDbType.Float, ToDouble(row["rb3"]));
                Add(command, "@new_avg", SqlDbType.Float, quantity == 0 ? 0d : cost / quantity);
                Add(command, "@new_pu", SqlDbType.Float, quantity == 0 ? 0d : sales / quantity);
                Add(command, "@cmt_sku_pline", SqlDbType.NVarChar, "", 500);
                SqlParameter output = command.Parameters.Add("@@num", SqlDbType.Int); output.Direction = ParameterDirection.Output;
                connection.Open(); command.ExecuteNonQuery(); skuId = Convert.ToInt32(output.Value);
            }
            const string linkSql = @"INSERT INTO dbo.T_sku_invoice_nbr
(id_issue,id_sku,invoice_nbr_sku,invoice_date_sku,order_nbr_sku,ship_date_sku,cust_po_sku)
VALUES(@ISSUE,@SKU,@INVOICE,@INVOICE_DATE,@ORDER,@SHIP_DATE,@PO);";
            ExecuteNonQuery(linkSql,
                new SqlParameter("@ISSUE", SqlDbType.Int) { Value = issueId }, new SqlParameter("@SKU", SqlDbType.Int) { Value = skuId },
                new SqlParameter("@INVOICE", SqlDbType.VarChar, 20) { Value = invoice.Invoice }, new SqlParameter("@INVOICE_DATE", SqlDbType.DateTime) { Value = invoice.InvoiceDate },
                new SqlParameter("@ORDER", SqlDbType.VarChar, 20) { Value = invoice.Order }, new SqlParameter("@SHIP_DATE", SqlDbType.DateTime) { Value = invoice.ShipDate },
                new SqlParameter("@PO", SqlDbType.VarChar, 50) { Value = invoice.CustomerPo });
        }

        private void ProcessInternalBlockage(string filePath, int emailDatabaseId, string subject)
        {
            DataTable table = ReadWorksheet(filePath, "blocage", true);
            if (!table.Columns.Contains("N°")) throw new InvalidDataException("Column N° not found in worksheet blocage");
            foreach (DataRow row in table.Rows)
                if (row["N°"] != DBNull.Value && int.TryParse(Convert.ToString(row["N°"]), out int issueId))
                    InsertComment(issueId, subject, "robot_mail", emailDatabaseId);
        }

        private void ProcessCashDiscount(string filePath, int emailDatabaseId)
        {
            using (DataSet dataSet = ReadWorkbook(filePath, true))
            {
                DataTable table = dataSet.Tables.Cast<DataTable>().FirstOrDefault(x => x.Columns.Contains("N° dossier IMIT"));
                if (table == null) throw new InvalidDataException("Column N° dossier IMIT not found");
                foreach (DataRow row in table.Rows)
                    if (row["N° dossier IMIT"] != DBNull.Value && int.TryParse(Convert.ToString(row["N° dossier IMIT"]), out int issueId))
                        InsertComment(issueId, "Réponse", "robot_mail", emailDatabaseId);
            }
        }

        private void ProcessComputacenter(string filePath, int emailDatabaseId)
        {
            int treatmentId = InsertComputacenterTreatment();
            DataTable destination = CreateComputacenterTable();
            using (DataSet workbook = ReadWorkbook(filePath, false))
            {
                List<DataTable> selectedSheets = SelectComputacenterSheets(workbook);
                foreach (DataTable sheet in selectedSheets) AppendComputacenterRows(sheet, destination, treatmentId);
            }
            ExecuteNonQuery("TRUNCATE TABLE dbo.T_ComputaCenter_Data_tmp;");
            using (var bulk = new SqlBulkCopy(sql_connexion))
            {
                bulk.BulkCopyTimeout = 300; bulk.DestinationTableName = "dbo.T_ComputaCenter_Data_tmp";
                foreach (DataColumn column in destination.Columns) bulk.ColumnMappings.Add(column.ColumnName, column.ColumnName);
                bulk.WriteToServer(destination);
            }
            ArchiveEmailInDatabase(emailDatabaseId);
            WriteLog("       Computacenter rows imported : " + destination.Rows.Count);
        }

        private int InsertComputacenterTreatment()
        {
            using (var connection = new SqlConnection(sql_connexion))
            using (var command = new SqlCommand("dbo.USP_INSERT_TRAITEMENT_COMPUTACENTER", connection))
            {
                command.CommandType = CommandType.StoredProcedure;
                SqlParameter output = command.Parameters.Add("@@id", SqlDbType.Int); output.Direction = ParameterDirection.Output;
                connection.Open(); command.ExecuteNonQuery(); return Convert.ToInt32(output.Value);
            }
        }

        private static List<DataTable> SelectComputacenterSheets(DataSet workbook)
        {
            string current = DateTime.Now.Day == 1 ? DateTime.Now.AddMonths(-1).ToString("MMMM", CultureInfo.CurrentCulture) : DateTime.Now.ToString("MMMM", CultureInfo.CurrentCulture);
            string previous = DateTime.Now.Day == 1 ? DateTime.Now.AddMonths(-2).ToString("MMMM", CultureInfo.CurrentCulture) : DateTime.Now.AddMonths(-1).ToString("MMMM", CultureInfo.CurrentCulture);
            return workbook.Tables.Cast<DataTable>().Where(x => StartsWithMonth(x.TableName, current) || StartsWithMonth(x.TableName, previous)).ToList();
        }

        private static bool StartsWithMonth(string sheetName, string month)
        {
            string a = RemoveDiacritics((sheetName ?? "").Trim('\'', '$')).ToLowerInvariant();
            string b = RemoveDiacritics(month ?? "").ToLowerInvariant();
            return a.StartsWith(b.Length >= 4 ? b.Substring(0, 4) : b);
        }

        private static DataTable CreateComputacenterTable()
        {
            var table = new DataTable();
            table.Columns.Add("id_traitement", typeof(int)); table.Columns.Add("date exp", typeof(DateTime));
            table.Columns.Add("Agce liv", typeof(string)); table.Columns.Add("cp", typeof(double)); table.Columns.Add("colis", typeof(double));
            table.Columns.Add("bl", typeof(string)); table.Columns.Add("poids", typeof(double)); table.Columns.Add("situation", typeof(string));
            table.Columns.Add("date", typeof(DateTime)); table.Columns.Add("heure", typeof(DateTime)); table.Columns.Add("Identite", typeof(string));
            table.Columns.Add("Ville", typeof(string)); table.Columns.Add("commentaires", typeof(string)); return table;
        }

        private static void AppendComputacenterRows(DataTable source, DataTable destination, int treatmentId)
        {
            // Le classeur historique commence son tableau à la ligne 3.
            int headerIndex = FindHeaderRow(source, "date exp");
            if (headerIndex < 0) return;
            string[] headers = source.Rows[headerIndex].ItemArray.Select(x => Convert.ToString(x).Trim()).ToArray();
            for (int rowIndex = headerIndex + 1; rowIndex < source.Rows.Count; rowIndex++)
            {
                DataRow row = source.Rows[rowIndex];
                string situation = GetCell(row, headers, "situation");
                bool previousSheetRule = true;
                if (previousSheetRule && (situation.Equals("livré", StringComparison.OrdinalIgnoreCase) || situation.Equals("Retour stock", StringComparison.OrdinalIgnoreCase))) continue;
                DataRow target = destination.NewRow();
                target["id_traitement"] = treatmentId;
                target["date exp"] = ToDbDate(GetCell(row, headers, "date exp")); target["Agce liv"] = GetCell(row, headers, "Agce liv");
                target["cp"] = ToDbDouble(GetCell(row, headers, "cp")); target["colis"] = ToDbDouble(GetCell(row, headers, "colis"));
                target["bl"] = GetCell(row, headers, "bl"); target["poids"] = ToDbDouble(GetCell(row, headers, "poids"));
                target["situation"] = situation; target["date"] = ToDbDate(GetCell(row, headers, "date"));
                target["heure"] = ToDbDate(GetCell(row, headers, "heure")); target["Identite"] = GetCell(row, headers, "Identite");
                target["Ville"] = GetCell(row, headers, "Ville"); target["commentaires"] = GetCell(row, headers, "commentaires");
                destination.Rows.Add(target);
            }
        }

        private bool ImportTrackingFile(int carrier, string destinationTable, string filePath)
        {
            string[] allowed = { "T_Tracking_Auto_DHL", "T_Tracking_Auto_CALBERSON", "T_Tracking_Auto_CHRONOPOST" };
            if (!allowed.Contains(destinationTable, StringComparer.OrdinalIgnoreCase)) throw new InvalidOperationException("Invalid tracking destination table");
            DataTable table;
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension == ".txt" || extension == ".csv")
                table = ReadDelimitedTable(filePath, ';', carrier == 3 ? "DETAIL DES ENVOIS INTERNATIONAUX" : "");
            else
                using (DataSet dataSet = ReadWorkbook(filePath, true)) table = dataSet.Tables[0].Copy();
            using (var bulk = new SqlBulkCopy(sql_connexion))
            {
                bulk.BulkCopyTimeout = 300; bulk.DestinationTableName = destinationTable; bulk.WriteToServer(table);
            }
            return true;
        }

        private string PrepareClaimIdFile(string filePath)
        {
            if (!Path.GetExtension(filePath).Equals(".zip", StringComparison.OrdinalIgnoreCase)) return filePath;
            return ExtractZipSafely(filePath, working_Folder);
        }

        private void ProcessClaimId(string filePath, string subject)
        {
            DataTable table = ReadWorksheet(filePath, "Sku", true);
            ExecuteNonQuery("TRUNCATE TABLE dbo.HP_ClaimID;");
            using (var bulk = new SqlBulkCopy(sql_connexion))
            {
                bulk.BulkCopyTimeout = 300; bulk.DestinationTableName = "dbo.HP_ClaimID"; bulk.WriteToServer(table);
            }
            var result = new DataTable();
            using (var connection = new SqlConnection(sql_connexion))
            using (var adapter = new SqlDataAdapter("SELECT * FROM dbo.vw_HPClaim_ID ORDER BY sceo", connection)) adapter.Fill(result);
            if (result.Rows.Count == 0) return;
            string folder = string.IsNullOrWhiteSpace(output_Files_Folder) ? tempFolder : output_Files_Folder;
            Directory.CreateDirectory(folder);
            string output = Path.Combine(folder, Path.GetFileNameWithoutExtension(filePath) + ".xlsx");
            WriteWorksheet(result, output, "Sku");
            WriteLog("       Claim ID result generated : " + output + " - Subject : " + subject);
        }

        private void TrackBusinessMethod(string methodName, string context, Action action)
        {
            TrackBusinessMethod<object>(methodName, context, () =>
            {
                action();
                return null;
            });
        }

        private T TrackBusinessMethod<T>(string methodName, string context, Func<T> action)
        {
            DateTime startedAtUtc = DateTime.UtcNow;
            bool succeeded = false;
            string errorMessage = "";
            try
            {
                T result = action();
                succeeded = true;
                return result;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                throw;
            }
            finally
            {
                long durationMs = Convert.ToInt64((DateTime.UtcNow - startedAtUtc).TotalMilliseconds);
                try
                {
                    InsertBusinessMethodUsage(methodName, startedAtUtc, succeeded, durationMs, context, errorMessage);
                }
                catch (Exception trackingException)
                {
                    WriteLog("       Unable to record business method usage for " + methodName + " : " + trackingException.Message);
                }
            }
        }

        private void InsertBusinessMethodUsage(string methodName, DateTime executedAtUtc, bool succeeded,
            long durationMs, string context, string errorMessage)
        {
            const string sql = @"INSERT INTO dbo.T_IMCA_BusinessMethodUsage
(CountryCode,MethodName,ExecutedAtUtc,Succeeded,DurationMs,Context,ErrorMessage)
VALUES(@COUNTRY,@METHOD,@EXECUTED,@SUCCEEDED,@DURATION,@CONTEXT,@ERROR);";
            ExecuteNonQuery(sql,
                new SqlParameter("@COUNTRY", SqlDbType.VarChar, 2) { Value = Truncate(country.ToUpperInvariant(), 2) },
                new SqlParameter("@METHOD", SqlDbType.NVarChar, 128) { Value = Truncate(methodName, 128) },
                new SqlParameter("@EXECUTED", SqlDbType.DateTime2) { Value = executedAtUtc },
                new SqlParameter("@SUCCEEDED", SqlDbType.Bit) { Value = succeeded },
                new SqlParameter("@DURATION", SqlDbType.BigInt) { Value = durationMs },
                new SqlParameter("@CONTEXT", SqlDbType.NVarChar, 500) { Value = Truncate(context, 500) },
                new SqlParameter("@ERROR", SqlDbType.NVarChar, 2000) { Value = Truncate(errorMessage, 2000) });
        }

        private void SendMonthlyBusinessUsageReportIfRequired()
        {
            DateTime reportMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).AddMonths(-1);
            if (!TryReserveMonthlyReport(reportMonth)) return;
            try
            {
                string html = BuildMonthlyBusinessUsageReport(reportMonth);
                SendBusinessUsageReport(reportMonth, html);
                CompleteMonthlyReport(reportMonth, "S", "");
                WriteLog("   Monthly business usage report sent for " + reportMonth.ToString("yyyy-MM"));
            }
            catch (Exception ex)
            {
                CompleteMonthlyReport(reportMonth, "E", ex.Message);
                throw;
            }
        }

        private bool TryReserveMonthlyReport(DateTime reportMonth)
        {
            using (var connection = new SqlConnection(sql_connexion))
            {
                connection.Open();
                using (SqlTransaction transaction = connection.BeginTransaction(IsolationLevel.Serializable))
                using (var command = new SqlCommand(@"SELECT Status,LastAttemptAtUtc
FROM dbo.T_IMCA_BusinessMethodMonthlyReport WITH (UPDLOCK,HOLDLOCK)
WHERE CountryCode=@COUNTRY AND ReportMonth=@MONTH;", connection, transaction))
                {
                    command.Parameters.Add("@COUNTRY", SqlDbType.VarChar, 2).Value = country.ToUpperInvariant();
                    command.Parameters.Add("@MONTH", SqlDbType.Date).Value = reportMonth;
                    string status = null;
                    DateTime? lastAttempt = null;
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            status = Convert.ToString(reader["Status"]);
                            if (reader["LastAttemptAtUtc"] != DBNull.Value) lastAttempt = Convert.ToDateTime(reader["LastAttemptAtUtc"]);
                        }
                    }
                    if (status == "S" || (status == "P" && lastAttempt.HasValue && lastAttempt.Value > DateTime.UtcNow.AddHours(-2)))
                    {
                        transaction.Commit();
                        return false;
                    }
                    string sql = status == null
                        ? @"INSERT INTO dbo.T_IMCA_BusinessMethodMonthlyReport(CountryCode,ReportMonth,Status,LastAttemptAtUtc)
VALUES(@COUNTRY,@MONTH,'P',SYSUTCDATETIME());"
                        : @"UPDATE dbo.T_IMCA_BusinessMethodMonthlyReport
SET Status='P',LastAttemptAtUtc=SYSUTCDATETIME(),ErrorMessage=NULL
WHERE CountryCode=@COUNTRY AND ReportMonth=@MONTH;";
                    using (var reserve = new SqlCommand(sql, connection, transaction))
                    {
                        reserve.Parameters.Add("@COUNTRY", SqlDbType.VarChar, 2).Value = country.ToUpperInvariant();
                        reserve.Parameters.Add("@MONTH", SqlDbType.Date).Value = reportMonth;
                        reserve.ExecuteNonQuery();
                    }
                    transaction.Commit();
                    return true;
                }
            }
        }

        private string BuildMonthlyBusinessUsageReport(DateTime reportMonth)
        {
            DateTime nextMonth = reportMonth.AddMonths(1);
            const string sql = @"SELECT c.MethodName,c.Description,c.MonitoringStartDate,
COUNT(u.Id) ExecutionCount,
SUM(CASE WHEN u.Succeeded=1 THEN 1 ELSE 0 END) SuccessCount,
SUM(CASE WHEN u.Succeeded=0 THEN 1 ELSE 0 END) ErrorCount,
MIN(u.ExecutedAtUtc) FirstExecutionUtc,MAX(u.ExecutedAtUtc) LastExecutionUtc,
AVG(CAST(u.DurationMs AS float)) AverageDurationMs
FROM dbo.T_IMCA_BusinessMethodCatalog c
LEFT JOIN dbo.T_IMCA_BusinessMethodUsage u
 ON u.CountryCode=c.CountryCode AND u.MethodName=c.MethodName
 AND u.ExecutedAtUtc>=@START AND u.ExecutedAtUtc<@END
WHERE c.CountryCode=@COUNTRY AND c.Enabled=1 AND c.MonitoringStartDate<@END
GROUP BY c.MethodName,c.Description,c.MonitoringStartDate
ORDER BY c.MethodName;";
            var html = new StringBuilder();
            html.Append("<h2>Rapport mensuel d'utilisation des traitements IMIT - ")
                .Append(WebUtility.HtmlEncode(country.ToUpperInvariant())).Append(" - ")
                .Append(reportMonth.ToString("MMMM yyyy", CultureInfo.GetCultureInfo("fr-FR"))).Append("</h2>")
                .Append("<table border='1' cellspacing='0' cellpadding='5'><tr><th>Méthode</th><th>Description</th><th>Statut</th><th>Exécutions</th><th>Succès</th><th>Erreurs</th><th>Première utilisation UTC</th><th>Dernière utilisation UTC</th><th>Durée moyenne ms</th></tr>");
            int rowCount = 0;
            using (var connection = new SqlConnection(sql_connexion))
            using (var command = new SqlCommand(sql, connection))
            {
                command.Parameters.Add("@COUNTRY", SqlDbType.VarChar, 2).Value = country.ToUpperInvariant();
                command.Parameters.Add("@START", SqlDbType.DateTime2).Value = reportMonth;
                command.Parameters.Add("@END", SqlDbType.DateTime2).Value = nextMonth;
                connection.Open();
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        rowCount++;
                        int executionCount = Convert.ToInt32(reader["ExecutionCount"]);
                        int successCount = Convert.ToInt32(reader["SuccessCount"]);
                        int errorCount = Convert.ToInt32(reader["ErrorCount"]);
                        string usageStatus = executionCount == 0 ? "Non utilisée sur la période"
                            : errorCount == executionCount ? "Toujours en erreur"
                            : errorCount > 0 ? "Utilisée avec erreurs"
                            : "Utilisée avec succès";
                        string firstExecution = reader["FirstExecutionUtc"] == DBNull.Value ? "" : Convert.ToDateTime(reader["FirstExecutionUtc"]).ToString("yyyy-MM-dd HH:mm:ss.fff");
                        string lastExecution = reader["LastExecutionUtc"] == DBNull.Value ? "" : Convert.ToDateTime(reader["LastExecutionUtc"]).ToString("yyyy-MM-dd HH:mm:ss.fff");
                        string averageDuration = reader["AverageDurationMs"] == DBNull.Value ? "" : Convert.ToDouble(reader["AverageDurationMs"]).ToString("0", CultureInfo.InvariantCulture);
                        html.Append("<tr><td>").Append(WebUtility.HtmlEncode(Convert.ToString(reader["MethodName"])))
                            .Append("</td><td>").Append(WebUtility.HtmlEncode(Convert.ToString(reader["Description"])))
                            .Append("</td><td>").Append(WebUtility.HtmlEncode(usageStatus))
                            .Append("</td><td>").Append(executionCount).Append("</td><td>")
                            .Append(successCount).Append("</td><td>").Append(errorCount)
                            .Append("</td><td>").Append(firstExecution).Append("</td><td>")
                            .Append(lastExecution).Append("</td><td>").Append(averageDuration).Append("</td></tr>");
                    }
                }
            }
            if (rowCount == 0) html.Append("<tr><td colspan='9'>Aucune méthode n'est déclarée dans le catalogue pour ce pays.</td></tr>");
            html.Append("</table><p>Les méthodes à 0 exécution restent visibles grâce au catalogue de référence.</p>");
            return html.ToString();
        }

        private void SendBusinessUsageReport(DateTime reportMonth, string html)
        {
            if (string.IsNullOrWhiteSpace(archive_Bcc_Mailbox)) throw new InvalidOperationException("IMIT report sender mailbox is empty");
            List<Recipient> recipients = BuildRecipients(email_in_case_of_technical_issue);
            if (recipients.Count == 0) throw new InvalidOperationException("Monthly report recipient is empty or invalid");
            var message = new Message
            {
                Subject = global_application_name + " - Rapport utilisation méthodes " + country.ToUpperInvariant() + " - " + reportMonth.ToString("yyyy-MM"),
                Body = new ItemBody { ContentType = BodyType.Html, Content = html },
                ToRecipients = recipients
            };
            var request = new Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody { Message = message, SaveToSentItems = true };
            graphService.Users[archive_Bcc_Mailbox].SendMail.PostAsync(request).GetAwaiter().GetResult();
        }

        private void CompleteMonthlyReport(DateTime reportMonth, string status, string error)
        {
            const string sql = @"UPDATE dbo.T_IMCA_BusinessMethodMonthlyReport
SET Status=@STATUS,SentAtUtc=CASE WHEN @STATUS='S' THEN SYSUTCDATETIME() ELSE SentAtUtc END,
ErrorMessage=@ERROR
WHERE CountryCode=@COUNTRY AND ReportMonth=@MONTH;";
            ExecuteNonQuery(sql,
                new SqlParameter("@STATUS", SqlDbType.Char, 1) { Value = status },
                new SqlParameter("@ERROR", SqlDbType.NVarChar, 2000) { Value = Truncate(error, 2000) },
                new SqlParameter("@COUNTRY", SqlDbType.VarChar, 2) { Value = country.ToUpperInvariant() },
                new SqlParameter("@MONTH", SqlDbType.Date) { Value = reportMonth });
        }
        private void ProcessPendingForwards()
        {
            List<ForwardRequest> forwards =
                GetPendingForwards();

            if (forwards.Count == 0)
            {
                if (IsTrue(debug))
                {
                    WriteLog(
                        "       No pending forward");
                }

                return;
            }

            TrackBusinessMethod(
                "ProcessPendingForwards",
                forwards.Count +
                " email(s) à transférer",
                () => ProcessPendingForwardRequests(
                    forwards));
        }

        private void ProcessPendingForwardRequests(
    List<ForwardRequest> forwards)
        {
            var errors = new List<Exception>();

            foreach (ForwardRequest request in forwards)
            {
                try
                {
                    UpdateForwardStatus(
                        request.Id,
                        "P");

                    ForwardWithGraph(
                        request);

                    UpdateForwardStatus(
                        request.Id,
                        "F");

                    WriteLog(
                        "       Forward completed. " +
                        "Database ID : " +
                        request.Id);
                }
                catch (Exception ex)
                {
                    try
                    {
                        UpdateForwardStatus(
                            request.Id,
                            "E");
                    }
                    catch
                    {
                    }

                    string type =
                        IsPermissionError(ex)
                            ? "SEND PERMISSION"
                            : "FORWARD";

                    string details =
                        "T_contenu_email ID " +
                        request.Id +
                        " - Source mailbox: " +
                        request.SourceMailbox +
                        " - Requested sender: " +
                        request.From +
                        " - " +
                        ex.Message;

                    WriteLog(
                        "       " +
                        type +
                        " error : " +
                        details);

                    SendTechnicalAlert(
                        nameof(ProcessPendingForwards),
                        details,
                        type);

                    errors.Add(
                        new TechnicalAlertAlreadySentException(
                            details,
                            ex));
                }
            }

            if (errors.Count > 0)
            {
                throw new AggregateException(
                    errors.Count +
                    " forward error(s).",
                    errors);
            }
        }


        private List<ForwardRequest> GetPendingForwards()
        {
            const string sql = @"SELECT id,ISNULL(EwsID,'') EwsID,ISNULL(boite_mail,'') boite_mail,
                                ISNULL(sujet,'') sujet,ISNULL(body,'') body,
                                ISNULL(new_to,'') new_to,ISNULL(new_cc,'') new_cc,ISNULL(new_from,'') new_from,
                                ISNULL(new_from_nom,'') new_from_nom,ISNULL(new_extra_body,'') new_extra_body,
                                fichier
                                FROM dbo.T_contenu_email
                                WHERE top_a_envoyer='O' ORDER BY date_demande_envoi,id;";
            var result = new List<ForwardRequest>();
            using (var connection = new SqlConnection(sql_connexion))
            using (var command = new SqlCommand(sql, connection))
            {
                command.CommandTimeout = 300;
                connection.Open();
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        result.Add(new ForwardRequest
                        {
                            Id = Convert.ToInt32(reader["id"]),
                            MessageId = Convert.ToString(reader["EwsID"]),
                            SourceMailbox = Convert.ToString(reader["boite_mail"]).Trim(),
                            Subject = Convert.ToString(reader["sujet"]),
                            StoredBody = Convert.ToString(reader["body"]),
                            To = Convert.ToString(reader["new_to"]),
                            Cc = Convert.ToString(reader["new_cc"]),
                            From = Convert.ToString(reader["new_from"]).Trim(),
                            FromName = Convert.ToString(reader["new_from_nom"]),
                            Comment = Convert.ToString(reader["new_extra_body"]),
                            MimeContent = reader["fichier"] == DBNull.Value ? Array.Empty<byte>() : (byte[])reader["fichier"]
                        });
                    }
                }
            }
            return result;
        }

        private void ForwardWithGraph(ForwardRequest request)
        {
            const string sendingMailbox = imit_mailbox_fr;
            List<Recipient> recipients = BuildRecipients(request.To);
            if (recipients.Count == 0)
                throw new InvalidOperationException("No valid recipient for T_contenu_email ID " + request.Id);

            string requestedSender = string.IsNullOrWhiteSpace(request.From) ? sendingMailbox : request.From;
            Recipient senderRecipient = new Recipient
            {
                EmailAddress = new EmailAddress
                {
                    Address = requestedSender,
                    Name = request.FromName ?? ""
                }
            };

            string subject = request.Subject ?? "";
            ItemBody body;
            List<Microsoft.Graph.Models.Attachment> attachments;

            if (request.MimeContent != null && request.MimeContent.Length > 0)
            {
                MimeMessage source = LoadMimeMessage(request.MimeContent);
                if (!string.IsNullOrWhiteSpace(source.Subject)) subject = source.Subject;
                body = BuildForwardBody(source, request);
                attachments = BuildGraphAttachments(source);
                WriteLog("       Forward ID " + request.Id + " reconstructed from stored MIME. MIME size: " +
                    request.MimeContent.Length + " byte(s). Attachments: " + attachments.Count);
            }
            else
            {
                body = BuildForwardBodyWithoutMime(request);
                attachments = new List<Microsoft.Graph.Models.Attachment>();
                WriteLog("       Warning: stored MIME is empty for T_contenu_email ID " + request.Id +
                    ". Message reconstructed from database subject and body without attachments.");
            }

            if (!subject.StartsWith("TR:", StringComparison.OrdinalIgnoreCase) &&
                !subject.StartsWith("FW:", StringComparison.OrdinalIgnoreCase))
                subject = "TR: " + subject;

            WriteLog("       Forward ID " + request.Id + " - sending mailbox: " + sendingMailbox +
                " - original mailbox: " + request.SourceMailbox + " - requested sender: " + requestedSender);

            var message = new Message
            {
                Subject = subject,
                Body = body,

                From = senderRecipient,

                ToRecipients = recipients,

                CcRecipients =
                    BuildRecipients(request.Cc),

                BccRecipients =
                    BuildRecipients(
                        archive_Bcc_Mailbox),

                Attachments = attachments
            };

            var sendRequest = new Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody
            {
                Message = message,
                SaveToSentItems = true
            };
            graphService.Users[sendingMailbox].SendMail.PostAsync(sendRequest).GetAwaiter().GetResult();
        }

        private static MimeMessage LoadMimeMessage(byte[] mimeContent)
        {
            if (mimeContent == null || mimeContent.Length == 0)
                throw new InvalidOperationException("The stored MIME content is empty");
            using (var stream = new MemoryStream(mimeContent, false))
                return MimeMessage.Load(stream);
        }

        private static ItemBody BuildForwardBody(MimeMessage source, ForwardRequest request)
        {
            string commentHtml = ConvertTextToHtml(request.Comment);
            string separator = string.IsNullOrWhiteSpace(commentHtml) ? "" : "<br/><br/>";
            if (!string.IsNullOrWhiteSpace(source.HtmlBody))
                return new ItemBody { ContentType = BodyType.Html, Content = commentHtml + separator + source.HtmlBody };

            string originalText = !string.IsNullOrWhiteSpace(source.TextBody) ? source.TextBody : request.StoredBody;
            return new ItemBody
            {
                ContentType = BodyType.Html,
                Content = commentHtml + separator + ConvertTextToHtml(originalText)
            };
        }

        private static ItemBody BuildForwardBodyWithoutMime(ForwardRequest request)
        {
            string commentHtml = ConvertTextToHtml(request.Comment);
            string bodyHtml = ConvertStoredBodyToHtml(request.StoredBody);
            var content = new StringBuilder();
            if (!string.IsNullOrWhiteSpace(commentHtml)) content.Append(commentHtml);
            if (!string.IsNullOrWhiteSpace(commentHtml) && !string.IsNullOrWhiteSpace(bodyHtml)) content.Append("<br/><br/>");
            if (!string.IsNullOrWhiteSpace(bodyHtml)) content.Append(bodyHtml);
            if (content.Length == 0) content.Append("Le contenu du message d'origine n'est plus disponible.");
            return new ItemBody { ContentType = BodyType.Html, Content = content.ToString() };
        }

        private static string ConvertStoredBodyToHtml(string body)
        {
            if (string.IsNullOrWhiteSpace(body)) return "";
            return LooksLikeHtml(body) ? body : ConvertTextToHtml(body);
        }

        private static bool LooksLikeHtml(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return false;
            return Regex.IsMatch(value, @"<\s*(html|body|div|p|br|table|span|a|img)(?:\s|>|/)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        private static string ConvertTextToHtml(string text)
        {
            return WebUtility.HtmlEncode(text ?? "").Replace("\r\n", "<br/>").Replace("\n", "<br/>");
        }

        private static List<Microsoft.Graph.Models.Attachment> BuildGraphAttachments(MimeMessage source)
        {
            var result = new List<Microsoft.Graph.Models.Attachment>();
            foreach (MimeEntity entity in source.BodyParts)
            {
                var part = entity as MimePart;
                if (part == null) continue;
                bool isInline = !part.IsAttachment && !string.IsNullOrWhiteSpace(part.ContentId);
                if (!part.IsAttachment && !isInline) continue;
                using (var stream = new MemoryStream())
                {
                    part.Content.DecodeTo(stream);
                    result.Add(new Microsoft.Graph.Models.FileAttachment
                    {
                        OdataType = "#microsoft.graph.fileAttachment",
                        Name = string.IsNullOrWhiteSpace(part.FileName) ? (isInline ? "inline" : "attachment") : part.FileName,
                        ContentType = part.ContentType?.MimeType ?? "application/octet-stream",
                        ContentBytes = stream.ToArray(),
                        ContentId = part.ContentId,
                        IsInline = isInline
                    });
                }
            }
            return result;
        }

        private string DownloadAttachmentByPrefix(string messageId, string prefix, string destinationFolder)
        {
            if (string.IsNullOrWhiteSpace(destinationFolder)) destinationFolder = tempFolder;
            Directory.CreateDirectory(destinationFolder);
            AttachmentCollectionResponse response = graphService.Users[mailboxAddress].Messages[messageId].Attachments.GetAsync().GetAwaiter().GetResult();
            foreach (Microsoft.Graph.Models.Attachment item in response?.Value ?? new List<Microsoft.Graph.Models.Attachment>())
            {
                if (!(item is FileAttachment file) || !(file.Name ?? "").StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) continue;
                if (file.ContentBytes == null && !string.IsNullOrWhiteSpace(file.Id))
                    file = graphService.Users[mailboxAddress].Messages[messageId].Attachments[file.Id].GetAsync().GetAwaiter().GetResult() as FileAttachment;
                if (file?.ContentBytes == null) throw new InvalidOperationException("Attachment content is empty : " + item.Name);
                string path = Path.Combine(destinationFolder, CleanFileName(file.Name));
                File.WriteAllBytes(path, file.ContentBytes);
                if (Path.GetExtension(path).Equals(".zip", StringComparison.OrdinalIgnoreCase)) return ExtractZipSafely(path, destinationFolder);
                return path;
            }
            WriteLog("       No attachment starting with '" + prefix + "' was found");
            return "";
        }

        private MailFolder GetInputFolder()
        {
            string requested = IsSpanishErmaMailbox() ? "Notificaciones eRMA" : sharedmailbox_folder_in;
            if (requested.Equals("Inbox", StringComparison.OrdinalIgnoreCase))
            {
                return ExecuteGraphFolderCallWithRetry(
                    () => graphService.Users[mailboxAddress].MailFolders["inbox"].GetAsync().GetAwaiter().GetResult(),
                    "Read Inbox folder for " + mailboxAddress);
            }

            MailFolder folder = FindFolderDirectlyUnderInbox(requested);
            if (folder == null) folder = FindFolderByNameRecursive(requested, true);
            if (folder == null) throw new DirectoryNotFoundException("Folder not found : " + requested);
            return folder;
        }

        private MailFolder TryGetErrorFolderAfterMessageFailure()
        {
            try
            {
                MailFolder folder = FindFolderDirectlyUnderInbox(sharedmailbox_folder_error);
                if (folder != null)
                {
                    if (IsTrue(debug)) WriteLog("       Error folder found directly under Inbox : " + folder.DisplayName);
                    return folder;
                }

                if (IsTrue(debug)) WriteLog("       Error folder not found directly under Inbox. Recursive search started.");
                folder = FindFolderByNameRecursive(sharedmailbox_folder_error, false);
                if (folder != null) return folder;

                string notFoundError = mailboxAddress + " - Required error folder '" +
                    sharedmailbox_folder_error + "' was not found after a message processing failure. " +
                    "The failed message cannot be moved, but mailbox processing will continue.";
                WriteLog("       " + notFoundError);
                SendTechnicalAlert(nameof(ReadCurrentMailbox), notFoundError, "ERROR FOLDER NOT FOUND");
                return null;
            }
            catch (Exception folderException)
            {
                string lookupError = mailboxAddress + " - Unable to search folder '" +
                    sharedmailbox_folder_error + "' after a message processing failure : " +
                    GetInnermostExceptionMessage(folderException) +
                    ". The failed message cannot be moved, but mailbox processing will continue.";
                WriteLog("       " + lookupError);
                SendTechnicalAlert(nameof(ReadCurrentMailbox), lookupError, "ERROR FOLDER LOOKUP");
                return null;
            }
        }

        private MailFolder FindFolderDirectlyUnderInbox(string folderName)
        {
            if (string.IsNullOrWhiteSpace(folderName)) return null;
            string safeName = folderName.Replace("'", "''");
            MailFolderCollectionResponse response = ExecuteGraphFolderCallWithRetry(
                () => graphService.Users[mailboxAddress].MailFolders["inbox"].ChildFolders.GetAsync(config =>
                {
                    config.QueryParameters.Top = 10;
                    config.QueryParameters.Filter = "displayName eq '" + safeName + "'";
                    config.QueryParameters.Select = new[] { "id", "displayName", "childFolderCount" };
                }).GetAwaiter().GetResult(),
                "Search direct child folder '" + folderName + "' under Inbox for " + mailboxAddress);

            return response?.Value?.FirstOrDefault(x =>
                string.Equals(x.DisplayName, folderName, StringComparison.OrdinalIgnoreCase));
        }

        private MailFolder FindFolderByNameRecursive(string folderName, bool required)
        {
            MailFolderCollectionResponse roots = ExecuteGraphFolderCallWithRetry(
                () => graphService.Users[mailboxAddress].MailFolders.GetAsync(config =>
                {
                    config.QueryParameters.Top = 100;
                    config.QueryParameters.Select = new[] { "id", "displayName", "childFolderCount" };
                }).GetAwaiter().GetResult(),
                "Read root folders for " + mailboxAddress);

            foreach (MailFolder root in roots?.Value ?? new List<MailFolder>())
            {
                MailFolder found = FindFolderRecursive(root, folderName);
                if (found != null) return found;
            }
            if (required) throw new DirectoryNotFoundException("Folder not found : " + folderName);
            return null;
        }

        private MailFolder FindFolderRecursive(MailFolder folder, string target)
        {
            if (string.Equals(folder.DisplayName, target, StringComparison.OrdinalIgnoreCase)) return folder;
            if ((folder.ChildFolderCount ?? 0) == 0) return null;

            MailFolderCollectionResponse children = ExecuteGraphFolderCallWithRetry(
                () => graphService.Users[mailboxAddress].MailFolders[folder.Id].ChildFolders.GetAsync(config =>
                {
                    config.QueryParameters.Top = 100;
                    config.QueryParameters.Select = new[] { "id", "displayName", "childFolderCount" };
                }).GetAwaiter().GetResult(),
                "Read child folders of '" + (folder.DisplayName ?? folder.Id) + "' for " + mailboxAddress);

            foreach (MailFolder child in children?.Value ?? new List<MailFolder>())
            {
                MailFolder found = FindFolderRecursive(child, target);
                if (found != null) return found;
            }
            return null;
        }

        private T ExecuteGraphFolderCallWithRetry<T>(Func<T> action, string operation)
        {
            const int maximumAttempts = 3;
            Exception lastException = null;
            for (int attempt = 1; attempt <= maximumAttempts; attempt++)
            {
                try
                {
                    return action();
                }
                catch (Exception ex) when (IsTransientGraphFolderError(ex))
                {
                    lastException = ex;
                    if (attempt >= maximumAttempts) break;
                    int delayMilliseconds = attempt * 2000;
                    WriteLog("       Temporary Graph folder error during " + operation +
                        ". Attempt " + attempt + "/" + maximumAttempts +
                        ". Retry in " + delayMilliseconds + " ms : " + GetInnermostExceptionMessage(ex));
                    System.Threading.Thread.Sleep(delayMilliseconds);
                }
            }
            throw new InvalidOperationException(
                "Graph folder operation failed after " + maximumAttempts + " attempts : " + operation +
                " - " + GetInnermostExceptionMessage(lastException), lastException);
        }

        private static bool IsTransientGraphFolderError(Exception exception)
        {
            Exception current = exception;
            while (current != null)
            {
                if (current is HttpRequestException || current is TimeoutException ||
                    current is System.Threading.Tasks.TaskCanceledException) return true;

                string message = current.Message ?? "";
                if (message.IndexOf("12002", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    message.IndexOf("timed out", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    message.IndexOf("operation was canceled", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    message.IndexOf("An error occurred while sending the request", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
                current = current.InnerException;
            }
            return false;
        }

        private static string GetInnermostExceptionMessage(Exception exception)
        {
            if (exception == null) return "";
            Exception current = exception;
            while (current.InnerException != null) current = current.InnerException;
            return current.Message ?? "";
        }

        private void MarkMessageAsProcessed(string messageId)
        {
            var processedProperty = new SingleValueLegacyExtendedProperty
            {
                Id = processedPropertyId,
                Value = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ", CultureInfo.InvariantCulture)
            };
            var update = new Message
            {
                SingleValueExtendedProperties = new List<SingleValueLegacyExtendedProperty>
                {
                    processedProperty
                }
            };
            graphService.Users[mailboxAddress].Messages[messageId].PatchAsync(update).GetAwaiter().GetResult();
        }

        private static bool IsMessageAlreadyProcessed(Message message)
        {
            return message?.SingleValueExtendedProperties?.Any(property =>
                string.Equals(property.Id, processedPropertyId, StringComparison.OrdinalIgnoreCase) &&
                !string.IsNullOrWhiteSpace(property.Value)) == true;
        }

        private static string EscapeODataString(string value)
        {
            return (value ?? "").Replace("'", "''");
        }

        private void MarkAsRead(string messageId)
        {
            graphService.Users[mailboxAddress].Messages[messageId].PatchAsync(new Message { IsRead = true }).GetAwaiter().GetResult();
        }

        private void MoveMessage(string messageId, string destinationId)
        {
            var body = new Microsoft.Graph.Users.Item.Messages.Item.Move.MovePostRequestBody { DestinationId = destinationId };
            graphService.Users[mailboxAddress].Messages[messageId].Move.PostAsync(body).GetAwaiter().GetResult();
        }

        private int InsertComment(int issueId, string comment, string commentName, int emailDatabaseId, string messageWebLink = "")
        {
            int commentId;
            using (var connection = new SqlConnection(sql_connexion))
            using (var command = new SqlCommand("dbo.USP_Insert_Comment", connection))
            {
                command.CommandType = CommandType.StoredProcedure;
                command.CommandTimeout = 300;
                Add(command, "@id_issue", SqlDbType.Int, issueId);
                Add(command, "@comment", SqlDbType.VarChar, Truncate(comment, 2000), 2000);
                Add(command, "@date_comment", SqlDbType.SmallDateTime, DateTime.Now);
                Add(command, "@comment_name", SqlDbType.VarChar, Truncate(commentName, 50), 50);
                Add(command, "@id_fichier", SqlDbType.Int, emailDatabaseId);
                SqlParameter output = command.Parameters.Add("@@num", SqlDbType.Int);
                output.Direction = ParameterDirection.Output;
                connection.Open();
                command.ExecuteNonQuery();
                commentId = output.Value == null || output.Value == DBNull.Value ? 0 : Convert.ToInt32(output.Value);
            }

            if (commentId <= 0)
                throw new InvalidOperationException("USP_Insert_Comment did not return an id_comment for issue " + issueId);

            LinkEmailToComment(emailDatabaseId, commentId);
            LinkEmailAttachmentsToComment(emailDatabaseId, commentId);

            if (string.IsNullOrWhiteSpace(messageWebLink))
                messageWebLink = GetMessageWebLinkFromGraph(emailDatabaseId);

            UpdateCommentWebLink(commentId, messageWebLink);
            WriteLog("       Email database ID " + emailDatabaseId + " linked to comment ID " + commentId + " for issue " + issueId);
            return commentId;
        }

        private string GetMessageWebLinkFromGraph(int emailDatabaseId)
        {
            const string sql = @"SELECT ISNULL(EwsID,'') AS GraphMessageId,
       ISNULL(boite_mail,'') AS Mailbox
FROM dbo.T_contenu_email
WHERE id=@EMAIL_ID;";

            string graphMessageId = "";
            string sourceMailbox = "";
            using (var connection = new SqlConnection(sql_connexion))
            using (var command = new SqlCommand(sql, connection))
            {
                command.CommandTimeout = 300;
                command.Parameters.Add("@EMAIL_ID", SqlDbType.Int).Value = emailDatabaseId;
                connection.Open();
                using (SqlDataReader reader = command.ExecuteReader(CommandBehavior.SingleRow))
                {
                    if (!reader.Read())
                        throw new InvalidOperationException("T_contenu_email row not found while resolving webLink. ID : " + emailDatabaseId);

                    graphMessageId = Convert.ToString(reader["GraphMessageId"]).Trim();
                    sourceMailbox = Convert.ToString(reader["Mailbox"]).Trim();
                }
            }

            if (string.IsNullOrWhiteSpace(graphMessageId) || string.IsNullOrWhiteSpace(sourceMailbox))
            {
                WriteLog("       Warning: Graph message ID or source mailbox is empty for email database ID " + emailDatabaseId + ".");
                return "";
            }

            Message graphMessage = graphService.Users[sourceMailbox].Messages[graphMessageId].GetAsync(config =>
            {
                config.QueryParameters.Select = new[] { "webLink" };
            }).GetAwaiter().GetResult();

            return graphMessage?.WebLink ?? "";
        }

        private void UpdateCommentWebLink(int commentId, string messageWebLink)
        {
            if (commentId <= 0) throw new ArgumentOutOfRangeException(nameof(commentId));
            if (string.IsNullOrWhiteSpace(messageWebLink))
            {
                WriteLog("       Warning: Graph did not return a webLink for comment ID " + commentId + ". T_comment.fichier was not updated.");
                return;
            }

            const string sql = @"UPDATE dbo.T_comment
SET fichier=@WEB_LINK
WHERE id_comment=@COMMENT_ID;
IF @@ROWCOUNT=0 THROW 50002, 'T_comment row not found while updating fichier', 1;";

            ExecuteNonQuery(sql,
                new SqlParameter("@WEB_LINK", SqlDbType.NVarChar, -1) { Value = messageWebLink },
                new SqlParameter("@COMMENT_ID", SqlDbType.Int) { Value = commentId });

            WriteLog("       T_comment.fichier updated with Graph webLink for comment ID " + commentId);
        }

        private void LinkEmailToComment(int emailDatabaseId, int commentId)
        {
            if (emailDatabaseId <= 0) throw new ArgumentOutOfRangeException(nameof(emailDatabaseId));
            if (commentId <= 0) throw new ArgumentOutOfRangeException(nameof(commentId));

            const string sql = @"UPDATE dbo.T_contenu_email
SET id_comment = CASE WHEN ISNULL(id_comment,0)=0 THEN @COMMENT_ID ELSE id_comment END
WHERE id=@EMAIL_ID;
IF @@ROWCOUNT=0 THROW 50001, 'T_contenu_email row not found while linking comment', 1;";

            ExecuteNonQuery(sql,
                new SqlParameter("@COMMENT_ID", SqlDbType.Int) { Value = commentId },
                new SqlParameter("@EMAIL_ID", SqlDbType.Int) { Value = emailDatabaseId });
        }

        private void LinkEmailAttachmentsToComment(int emailDatabaseId, int commentId)
        {
            if (emailDatabaseId <= 0) throw new ArgumentOutOfRangeException(nameof(emailDatabaseId));
            if (commentId <= 0) throw new ArgumentOutOfRangeException(nameof(commentId));

            string graphMessageId = GetStoredGraphMessageId(emailDatabaseId);
            if (string.IsNullOrWhiteSpace(graphMessageId))
            {
                WriteLog("       No Graph message ID found for email database ID " + emailDatabaseId + ". Attachment links were not created.");
                return;
            }

            AttachmentCollectionResponse response = graphService.Users[mailboxAddress]
                .Messages[graphMessageId].Attachments.GetAsync(config =>
                {
                    config.QueryParameters.Top = 999;
                    config.QueryParameters.Select = new[] { "id", "name" };
                }).GetAwaiter().GetResult();

            int linkedCount = 0;
            foreach (Microsoft.Graph.Models.Attachment attachment in response?.Value ?? new List<Microsoft.Graph.Models.Attachment>())
            {
                if (attachment == null || string.IsNullOrWhiteSpace(attachment.Id)) continue;
                InsertAttachmentLinkIfMissing(commentId, attachment.Name, attachment.Id);
                linkedCount++;
            }

            WriteLog("       Attachment link(s) created for comment ID " + commentId + " : " + linkedCount);
        }

        private string GetStoredGraphMessageId(int emailDatabaseId)
        {
            const string sql = @"SELECT ISNULL(EwsID,'') FROM dbo.T_contenu_email WHERE id=@EMAIL_ID;";
            using (var connection = new SqlConnection(sql_connexion))
            using (var command = new SqlCommand(sql, connection))
            {
                command.CommandTimeout = 300;
                command.Parameters.Add("@EMAIL_ID", SqlDbType.Int).Value = emailDatabaseId;
                connection.Open();
                return Convert.ToString(command.ExecuteScalar()).Trim();
            }
        }

        private void InsertAttachmentLinkIfMissing(int commentId, string attachmentName, string attachmentGraphId)
        {
            const string sql = @"IF NOT EXISTS
(
    SELECT 1
    FROM dbo.T_contenu_email_fichiers_joints
    WHERE id_comment=@COMMENT_ID
      AND EwsID COLLATE Latin1_General_CS_AS=@ATTACHMENT_ID
)
BEGIN
    INSERT INTO dbo.T_contenu_email_fichiers_joints
    (
        id_comment,
        nom_pieces_jointes,
        EwsID
    )
    VALUES
    (
        @COMMENT_ID,
        @ATTACHMENT_NAME,
        @ATTACHMENT_ID
    );
END;";

            ExecuteNonQuery(sql,
                new SqlParameter("@COMMENT_ID", SqlDbType.Int) { Value = commentId },
                new SqlParameter("@ATTACHMENT_NAME", SqlDbType.NVarChar, 500) { Value = Truncate(attachmentName, 500) },
                new SqlParameter("@ATTACHMENT_ID", SqlDbType.NVarChar, 1000) { Value = attachmentGraphId });
        }

        private void InsertHistory(int issueId, int status, string comment, string user)
        {
            const string sql = @"INSERT INTO dbo.T_histo(id_issue,comment,histo_date,user_name,id_status)
VALUES(@ISSUE,@COMMENT,GETDATE(),@USER,@STATUS);";
            ExecuteNonQuery(sql, new SqlParameter("@ISSUE", SqlDbType.Int) { Value = issueId },
                new SqlParameter("@COMMENT", SqlDbType.VarChar, 2000) { Value = comment ?? "" },
                new SqlParameter("@USER", SqlDbType.VarChar, 50) { Value = user ?? "" },
                new SqlParameter("@STATUS", SqlDbType.Int) { Value = status });
        }

        private string GetOwner(string customer)
        {
            return ExecuteScalarString(@"SELECT ISNULL(e.employee_name,'') FROM dbo.customer c INNER JOIN dbo.employee e ON c.os_salesman=e.employee WHERE c.branch_customer_nbr=@CUSTOMER;", customer);
        }

        private string GetOwnerSsc(string customer)
        {
            return ExecuteScalarString(@"SELECT ISNULL(gest_retro,'') FROM dbo.T_gestionnaire_client_retro_desk WHERE br_cust_nbr=@CUSTOMER;", customer);
        }

        private string ExecuteScalarString(string sql, string customer)
        {
            using (var connection = new SqlConnection(sql_connexion))
            using (var command = new SqlCommand(sql, connection))
            {
                Add(command, "@CUSTOMER", SqlDbType.VarChar, customer ?? "", 20); connection.Open(); return Convert.ToString(command.ExecuteScalar());
            }
        }

        private int FindExistingFedexIssue(string key, int currentId)
        {
            const string sql = @"SELECT TOP (1) c.id_issue FROM dbo.T_comment c WHERE c.id_contenu_email IN
(SELECT TOP (1) id FROM dbo.T_contenu_email WHERE id<>@ID AND num_fiche_fedex_colis=@KEY ORDER BY id) ORDER BY c.id_issue;";
            using (var connection = new SqlConnection(sql_connexion))
            using (var command = new SqlCommand(sql, connection))
            {
                Add(command, "@ID", SqlDbType.Int, currentId); Add(command, "@KEY", SqlDbType.VarChar, key ?? "", -1);
                connection.Open(); object value = command.ExecuteScalar(); return value == null || value == DBNull.Value ? 0 : Convert.ToInt32(value);
            }
        }

        private void UpdateConversationInformation(string referenceId, string key, int emailId)
        {
            const string sql = "UPDATE dbo.T_contenu_email SET ReferenceID=@REFERENCE,num_fiche_fedex_colis=@KEY WHERE id=@ID;";
            ExecuteNonQuery(sql, new SqlParameter("@REFERENCE", SqlDbType.VarChar, -1) { Value = referenceId ?? "" },
                new SqlParameter("@KEY", SqlDbType.VarChar, -1) { Value = key ?? "" }, new SqlParameter("@ID", SqlDbType.Int) { Value = emailId });
        }

        private void ArchiveEmailInDatabase(int emailId)
        {
            ExecuteNonQuery("UPDATE dbo.T_contenu_email SET archiver=1 WHERE id=@ID;", new SqlParameter("@ID", SqlDbType.Int) { Value = emailId });
        }

        private void UpdateForwardStatus(int emailId, string status)
        {
            ExecuteNonQuery("UPDATE dbo.T_contenu_email SET top_a_envoyer=@STATUS WHERE id=@ID;",
                new SqlParameter("@STATUS", SqlDbType.VarChar, 1) { Value = status }, new SqlParameter("@ID", SqlDbType.Int) { Value = emailId });
        }

        private void UpdateMailboxRefreshInformation(int id)
        {
            const string sql = @"UPDATE dbo.T_SharedMailboxes SET date_dernier_raf=GETDATE(),
dt_heure_filtre=COALESCE((SELECT MAX(dt_time_received) FROM dbo.T_contenu_email WHERE id_mailboxe=@ID),dt_heure_filtre)
WHERE id_mailboxe=@ID;";
            ExecuteNonQuery(sql, new SqlParameter("@ID", SqlDbType.Int) { Value = id });
        }

        private void ExecuteNonQuery(string sql, params SqlParameter[] parameters)
        {
            using (var connection = new SqlConnection(sql_connexion))
            using (var command = new SqlCommand(sql, connection))
            {
                command.CommandTimeout = 300; if (parameters != null && parameters.Length > 0) command.Parameters.AddRange(parameters);
                connection.Open(); command.ExecuteNonQuery();
            }
        }

        private void SendTechnicalAlert(string method, string error, string type)
        {
            try
            {
                WriteLog("       " + type + " error in " + method + " : " + error);
                if (graphService == null || string.IsNullOrWhiteSpace(email_in_case_of_technical_issue)) return;
                string sender = !string.IsNullOrWhiteSpace(mailboxAddress) ? mailboxAddress : archive_Bcc_Mailbox;
                if (string.IsNullOrWhiteSpace(sender)) return;
                var message = new Message
                {
                    Subject = global_application_name + " - Erreur " + type + " dans " + method,
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Html,
                        Content = "<b>Mailbox :</b> " + WebUtility.HtmlEncode(sender) + "<br/><b>Méthode :</b> " + WebUtility.HtmlEncode(method) + "<br/><b>Message :</b> " + WebUtility.HtmlEncode(error)
                    },
                    ToRecipients = BuildRecipients(email_in_case_of_technical_issue)
                };
                var body = new Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody { Message = message, SaveToSentItems = true };
                graphService.Users[sender].SendMail.PostAsync(body).GetAwaiter().GetResult();
            }
            catch (Exception ex) { WriteLog("Error sending technical alert : " + ex.Message); }
        }

        private string GetImcaParameter(string connectionString, string parameter)
        {
            if (string.IsNullOrWhiteSpace(parameter)) return "";
            const string sql = @"SELECT ISNULL(VALUE,'') FROM PCM_TAB_IMCA_PARAMETER_GLOBAL WHERE SK_VALID=0 AND PARAMETER=@PARAMETER;";
            using (var connection = new SqlConnection(connectionString))
            using (var command = new SqlCommand(sql, connection))
            {
                Add(command, "@PARAMETER", SqlDbType.NVarChar, parameter, 255); connection.Open(); return Convert.ToString(command.ExecuteScalar());
            }
        }

        private static DataSet ReadWorkbook(string filePath, bool firstRowHeader)
        {
            using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
            {
                return reader.AsDataSet(new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = firstRowHeader }
                });
            }
        }

        private static DataTable ReadWorksheet(string filePath, string worksheetName, bool firstRowHeader)
        {
            using (DataSet dataSet = ReadWorkbook(filePath, firstRowHeader))
            {
                DataTable table = dataSet.Tables.Cast<DataTable>().FirstOrDefault(x => string.Equals(x.TableName.Trim('\'', '$'), worksheetName.Trim('\'', '$'), StringComparison.OrdinalIgnoreCase));
                if (table == null) throw new InvalidDataException("Worksheet not found : " + worksheetName);
                return table.Copy();
            }
        }

        private static void WriteWorksheet(DataTable table, string filePath, string worksheetName)
        {
            using (var workbook = new XLWorkbook())
            {
                workbook.Worksheets.Add(table, CleanWorksheetName(worksheetName)); workbook.SaveAs(filePath);
            }
        }

        private static DataTable ReadDelimitedTable(string filePath, char delimiter, string marker)
        {
            List<string> lines = File.ReadAllLines(filePath).ToList();
            if (!string.IsNullOrWhiteSpace(marker))
            {
                int index = lines.FindIndex(x => x.IndexOf(marker, StringComparison.OrdinalIgnoreCase) >= 0);
                if (index >= 0) lines = lines.Skip(index + 1).ToList();
            }
            var table = new DataTable();
            if (lines.Count == 0) return table;
            string[] headers = ParseDelimitedLine(lines[0], delimiter).ToArray();
            foreach (string header in headers) table.Columns.Add(UniqueColumnName(table, string.IsNullOrWhiteSpace(header) ? "Column" : header));
            foreach (string line in lines.Skip(1))
            {
                string[] values = ParseDelimitedLine(line, delimiter).ToArray(); DataRow row = table.NewRow();
                for (int i = 0; i < Math.Min(values.Length, table.Columns.Count); i++) row[i] = values[i]; table.Rows.Add(row);
            }
            return table;
        }

        private static IEnumerable<string> ParseDelimitedLine(string line, char delimiter)
        {
            var value = new StringBuilder(); bool quoted = false;
            for (int i = 0; i < (line ?? "").Length; i++)
            {
                char c = line[i];
                if (c == '"')
                {
                    if (quoted && i + 1 < line.Length && line[i + 1] == '"') { value.Append('"'); i++; }
                    else quoted = !quoted;
                }
                else if (c == delimiter && !quoted) { yield return value.ToString(); value.Clear(); }
                else value.Append(c);
            }
            yield return value.ToString();
        }

        private static string ExtractZipSafely(string zipPath, string destinationFolder)
        {
            Directory.CreateDirectory(destinationFolder);
            string root = Path.GetFullPath(destinationFolder).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            string firstFile = "";
            using (ZipArchive archive = ZipFile.OpenRead(zipPath))
            {
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    if (string.IsNullOrEmpty(entry.Name)) continue;
                    string destination = Path.GetFullPath(Path.Combine(destinationFolder, entry.FullName));
                    if (!destination.StartsWith(root, StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("Unsafe ZIP entry : " + entry.FullName);
                    Directory.CreateDirectory(Path.GetDirectoryName(destination)); entry.ExtractToFile(destination, true);
                    if (string.IsNullOrWhiteSpace(firstFile)) firstFile = destination;
                }
            }
            return firstFile;
        }

        private static string FindBodyValue(string body, params string[] labels)
        {
            foreach (string line in (body ?? "").Replace("\r\n", "\n").Split('\n'))
            {
                string normalized = line.Trim().ToLowerInvariant();
                if (!labels.Any(x => normalized.Contains(x.ToLowerInvariant()))) continue;
                int index = line.LastIndexOf(':'); if (index >= 0 && index + 1 < line.Length) return line.Substring(index + 1).Trim();
            }
            return "";
        }

        private static string FindSecondValue(string body, string marker, int length)
        {
            int first = body.IndexOf(marker, StringComparison.OrdinalIgnoreCase); if (first < 0) return "";
            int second = body.IndexOf(marker, first + marker.Length, StringComparison.OrdinalIgnoreCase); if (second < 0) return "";
            int start = second + marker.Length; return start >= body.Length ? "" : body.Substring(start, Math.Min(length, body.Length - start)).Trim();
        }

        private static string FindFixedValue(string body, string marker, int length)
        {
            int index = body.IndexOf(marker, StringComparison.OrdinalIgnoreCase); if (index < 0) return "";
            int start = index + marker.Length; return start >= body.Length ? "" : body.Substring(start, Math.Min(length, body.Length - start)).Trim();
        }

        private static string FindSectionValue(string body, string startMarker, string endMarker)
        {
            int start = body.IndexOf(startMarker, StringComparison.OrdinalIgnoreCase); if (start < 0) return ""; start += startMarker.Length;
            int end = body.IndexOf(endMarker, start, StringComparison.OrdinalIgnoreCase); if (end < 0) end = body.Length;
            return body.Substring(start, end - start).Trim();
        }

        private static string GetFedexKey(string subject)
        {
            string[] parts = (subject ?? "").Split('/'); return parts.Length >= 5 ? parts[4].Trim() + "#" + parts[1].Trim() : "";
        }

        private static int FindHeaderRow(DataTable table, string expectedHeader)
        {
            for (int i = 0; i < Math.Min(table.Rows.Count, 20); i++)
                if (table.Rows[i].ItemArray.Any(x => string.Equals(Convert.ToString(x).Trim(), expectedHeader, StringComparison.OrdinalIgnoreCase))) return i;
            return -1;
        }

        private static string GetCell(DataRow row, string[] headers, string name)
        {
            int index = Array.FindIndex(headers, x => string.Equals(x, name, StringComparison.OrdinalIgnoreCase)); return index >= 0 && index < row.ItemArray.Length ? Convert.ToString(row[index]).Trim() : "";
        }

        private static object ToDbDate(string value)
        {
            if (DateTime.TryParse(value, CultureInfo.CurrentCulture, DateTimeStyles.None, out DateTime date) || DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out date)) return date;
            if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double serial)) return DateTime.FromOADate(serial);
            return DBNull.Value;
        }

        private static object ToDbDouble(string value)
        {
            if (double.TryParse(value, NumberStyles.Any, CultureInfo.CurrentCulture, out double number) || double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out number)) return number;
            return DBNull.Value;
        }

        private static double ToDouble(object value)
        {
            if (value == null || value == DBNull.Value) return 0d; return double.TryParse(Convert.ToString(value), NumberStyles.Any, CultureInfo.CurrentCulture, out double result) ? result : 0d;
        }

        private static int ToInt(object value)
        {
            if (value == null || value == DBNull.Value) return 0; return int.TryParse(Convert.ToString(value), out int result) ? result : Convert.ToInt32(ToDouble(value));
        }

        private static List<Recipient> BuildRecipients(string addresses)
        {
            return (addresses ?? "").Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).Where(IsEmail)
                .Distinct(StringComparer.OrdinalIgnoreCase).Select(x => new Recipient { EmailAddress = new EmailAddress { Address = x } }).ToList();
        }

        private static IEnumerable<string> GetRecipientAddresses(IEnumerable<Recipient> recipients)
        {
            return recipients == null ? Enumerable.Empty<string>() : recipients.Select(x => x?.EmailAddress?.Address).Where(x => !string.IsNullOrWhiteSpace(x));
        }

        private static string GetRecipients(IEnumerable<Recipient> recipients) { return string.Join(";", GetRecipientAddresses(recipients)); }
        private static string GetSender(Message email) { return email.From?.EmailAddress?.Address ?? email.Sender?.EmailAddress?.Address ?? email.From?.EmailAddress?.Name ?? ""; }
        private static string GetTextBody(Message email) { return email.Body?.Content ?? email.BodyPreview ?? ""; }
        private static bool IsEmail(string value) { return !string.IsNullOrWhiteSpace(value) && Regex.IsMatch(value, @"^[^\s@]+@[^\s@]+\.[^\s@]+$"); }
        private static string GetDomain(string address) { int index = (address ?? "").LastIndexOf('@'); return index >= 0 ? address.Substring(index).Trim().ToLowerInvariant() : ""; }
        private static string NormalizeDomain(string value) { value = (value ?? "").Trim().ToLowerInvariant(); return value.Length == 0 ? "" : value.StartsWith("@") ? value : "@" + value; }
        private static string Truncate(string value, int length) { value = value ?? ""; return value.Length <= length ? value : value.Substring(0, length); }
        private static string CleanFileName(string value) { return string.Join("_", (value ?? "attachment").Split(Path.GetInvalidFileNameChars())).Trim(); }
        private static string CleanWorksheetName(string value) { string name = Regex.Replace(value ?? "Data", @"[\[\]:*?/\\]", "_"); return name.Length <= 31 ? name : name.Substring(0, 31); }
        private static string UniqueColumnName(DataTable table, string name) { string result = name; int index = 2; while (table.Columns.Contains(result)) result = name + "_" + index++; return result; }
        private static string RemoveDiacritics(string value) { return new string((value ?? "").Normalize(NormalizationForm.FormD).Where(c => CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark).ToArray()).Normalize(NormalizationForm.FormC); }
        private bool IsSpanishErmaMailbox()
        {
            return country.Equals("ES", StringComparison.OrdinalIgnoreCase) && mailboxId == 4 && !IsCurrentCountryImitMailbox();
        }
        private static bool IsTrue(string value) { return string.Equals(value?.Trim(), "TRUE", StringComparison.OrdinalIgnoreCase); }
        private static bool IsRefreshDue(DateTime date, int minutes) { return minutes <= 0 || DateTime.Now >= date.AddMinutes(minutes); }

        private DateTime ParseStartDate()
        {
            if (string.IsNullOrWhiteSpace(start_date_scan)) return new DateTime(1900, 1, 1);
            if (!DateTime.TryParseExact(start_date_scan, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime result))
                throw new InvalidOperationException("Invalid start_date_scan. Expected yyyyMMdd");
            return result;
        }

        private static bool IsPermissionError(Exception exception)
        {
            if (exception is ApiException api && (api.ResponseStatusCode == 401 || api.ResponseStatusCode == 403)) return true;
            string text = exception?.ToString() ?? "";
            return new[] { "ErrorAccessDenied", "AccessDenied", "SendAsDenied", "Authorization_RequestDenied", "403" }
                .Any(x => text.IndexOf(x, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private void ValidateCountryConfiguration()
        {
            if (string.IsNullOrWhiteSpace(sql_connexion)) throw new InvalidOperationException("sql_connexion is empty");
            if (string.IsNullOrWhiteSpace(tempFolder)) throw new InvalidOperationException("temp_folder is empty");
            if (string.IsNullOrWhiteSpace(email_in_case_of_technical_issue)) throw new InvalidOperationException("technical email is empty");
        }

        private void ValidateMailboxConfiguration()
        {
            if (mailboxId <= 0) throw new InvalidOperationException("id_mailboxe is invalid");
            if (string.IsNullOrWhiteSpace(mailboxAddress)) throw new InvalidOperationException("mailboxe is empty");
        }

        private void WriteLog(string message)
        {
            if (string.IsNullOrWhiteSpace(logsFolder)) logsFolder = AppDomain.CurrentDomain.BaseDirectory;
            Directory.CreateDirectory(logsFolder);
            string file = Path.Combine(logsFolder, "IMCA_" + sessionName + "_" + DateTime.Now.ToString("dd_MM_yyyy") + "_" + country + "_" + global_application_name + ".txt");
            File.AppendAllText(file, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - " + message + Environment.NewLine);
        }

        private static string GetServicePath()
        {
            string location = System.Reflection.Assembly.GetEntryAssembly()?.Location;
            return string.IsNullOrWhiteSpace(location) ? AppDomain.CurrentDomain.BaseDirectory : Path.GetDirectoryName(location);
        }

        private static void Add(SqlCommand command, string name, SqlDbType type, object value, int size = 0)
        {
            SqlParameter parameter = size == 0 ? command.Parameters.Add(name, type) : command.Parameters.Add(name, type, size);
            parameter.Value = value ?? DBNull.Value;
        }
    }
}
