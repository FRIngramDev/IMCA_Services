using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
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
        private const int retry_delay_hours = 2;
        private string country = "", name = "", active = "", debug = "", start_date_scan = "", number_Of_Mails = "10";
        private string sharedmailbox_folder_in = "Inbox", sharedmailbox_folder_out = "Archives";
        private string sql_connexion_parameter_global = "", sql_connexion = "";
        private string sql_gestion_cdes_parameter_global = "", sql_gestion_cdes = "";
        private string sql_dss_copie_parameter_global = "", sql_dss_copie = "";
        private string email_in_case_of_technical_issue_parameter_global = "", email_in_case_of_technical_issue = "";
        private string fr_graph_send_as_parameter_global = "", fr_graph_send_as = "";
        private string credit_managers_contentieux = "", contentieux_recipients = "";
        private string path_archives = "";
        private string dss_con_openrowset_parameter_global = "";
        private string dss_con_openrowset = "";
        private string uri_webservice = "";
        private HashSet<string> credit_managers_contentieux_list = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private string logsFolder = "", tempFolder = "", sessionName = "";
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

            public string email_in_case_of_technical_issue_parameter_global { get; set; } = "";

            public string fr_graph_send_as_parameter_global { get; set; } = "";

            public string credit_managers_contentieux { get; set; } = "";

            public string contentieux_recipients { get; set; } = "";
            public string path_archives { get; set; } = "";
            public string dss_con_openrowset_parameter_global { get; set; } = "";
            public string uri_webservice { get; set; } = "";
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
                        throw new AggregateException(
                            errors.Count + " credit mailbox processing error(s).",
                            errors);
                    }

                    WriteLog(
                        "   Country processing completed : " +
                        country.ToUpperInvariant());
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
            email_in_case_of_technical_issue_parameter_global = i.email_in_case_of_technical_issue_parameter_global ?? "";
            fr_graph_send_as_parameter_global = i.fr_graph_send_as_parameter_global ?? "";
            credit_managers_contentieux = i.credit_managers_contentieux ?? "";
            contentieux_recipients = i.contentieux_recipients ?? "";
            path_archives = i.path_archives ?? "";
            dss_con_openrowset_parameter_global = i.dss_con_openrowset_parameter_global ?? "";
            uri_webservice = i.uri_webservice ?? "";
            sql_connexion = GetImcaParameter(imca, sql_connexion_parameter_global);
            sql_gestion_cdes = GetImcaParameter(imca, sql_gestion_cdes_parameter_global);
            sql_dss_copie = GetImcaParameter(imca, sql_dss_copie_parameter_global);
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
            Exception last = null;
            for (int i = 1;
            i <= 3;
            i++) try
                {
                    return action();
                }
                catch (Exception ex) when (IsTransientGraphError(ex))
                {
                    last = ex;
                    if (i >= 3) break;
                    System.Threading.Thread.Sleep(i * 2000);
                }
            throw new InvalidOperationException("Graph operation failed after 3 attempts : " + operation + " - " + GetInnermostExceptionMessage(last), last);
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

        private void WriteLog(string m)
        {
            Directory.CreateDirectory(logsFolder);
            File.AppendAllText(Path.Combine(logsFolder, "IMCA_" + sessionName + "_" + DateTime.Now.ToString("dd_MM_yyyy") + "_" + country + "_" + global_application_name + ".txt"), DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - " + m + Environment.NewLine);
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

        private static bool IsChangeKeyConflict(Exception ex)
        {
            string t = ex?.ToString() ?? "";
            return t.IndexOf("change key", StringComparison.OrdinalIgnoreCase) >= 0 || t.IndexOf("ErrorIrresolvableConflict", StringComparison.OrdinalIgnoreCase) >= 0 || (ex is ApiException a && (a.ResponseStatusCode == 409 || a.ResponseStatusCode == 412));
        }

        private static bool IsTransientGraphError(Exception ex)
        {
            while (ex != null)
            {
                if (ex is HttpRequestException || ex is TimeoutException || ex is System.Threading.Tasks.TaskCanceledException) return true;
                string t = ex.Message ?? "";
                if (t.IndexOf("12002", StringComparison.OrdinalIgnoreCase) >= 0 || t.IndexOf("timed out", StringComparison.OrdinalIgnoreCase) >= 0 || t.IndexOf("An error occurred while sending the request", StringComparison.OrdinalIgnoreCase) >= 0) return true;
                ex = ex.InnerException;
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
    }
}
