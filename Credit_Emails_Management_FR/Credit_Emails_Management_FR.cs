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
namespace CREDIT_EMAILS_MANAGEMENT_FR
{
    public class CREDIT_EMAILS_MANAGEMENT_FR
    {
        private const string global_application_name = "CREDIT_EMAILS_MANAGEMENT_FR";
        private const string immutable_id_preference = "IdType=\"ImmutableId\"";
        private const string processed_property_id = "String {8BF48C6E-2C72-46F0-965D-919A7C2E54A9} Name IMCACreditProcessed";
        private const string credit_card_mailbox_name = "Cartes Bleues";
        private const int retry_delay_hours = 2;
        private string country = "", name = "", active = "", debug = "", start_date_scan = "", number_Of_Mails = "10";
        private string sharedmailbox_folder_in = "Inbox", sharedmailbox_folder_out = "Archives";
        private string sql_connexion_parameter_global = "", sql_connexion = "";
        private string sql_gestion_cdes_parameter_global = "", sql_gestion_cdes = "";
        private string sql_dss_copie_parameter_global = "", sql_dss_copie = "";
        private string email_in_case_of_technical_issue_parameter_global = "", email_in_case_of_technical_issue = "";
        private string fr_graph_send_as_parameter_global = "", fr_graph_send_as = "";
        private string credit_managers_contentieux = "", contentieux_recipients = "";
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
            sql_connexion = GetImcaParameter(imca, sql_connexion_parameter_global);
            sql_gestion_cdes = GetImcaParameter(imca, sql_gestion_cdes_parameter_global);
            sql_dss_copie = GetImcaParameter(imca, sql_dss_copie_parameter_global);
            email_in_case_of_technical_issue = GetImcaParameter(imca, email_in_case_of_technical_issue_parameter_global);
            fr_graph_send_as = GetImcaParameter(imca, fr_graph_send_as_parameter_global);
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
            inputFolderName = b.InputFolder;
            outputFolderName = b.OutputFolder;
            allowedSenders = SplitValues(b.SenderAllowed).ToHashSet(StringComparer.OrdinalIgnoreCase);
        }

        private DateTime? ReadCurrentMailbox()
        {
            ValidateMailboxConfiguration();

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
                q.QueryParameters.Filter = "receivedDateTime gt " + date.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ", CultureInfo.InvariantCulture);
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

            // Comportement historique : sans numéro d'autorisation,
            // aucun traitement métier n'est exécuté et aucune alerte technique
            // n'est envoyée. Le message est journalisé puis archivé normalement.
            if (string.IsNullOrWhiteSpace(data.AuthorizationNumber))
            {
                WriteLog(
                    "       Authorization number is missing. " +
                    "Business processing skipped and message archived.");

                FinalizeGraphMessage(email.Id, outputFolderId);

                UpsertProcessingState(
                    email,
                    "S",
                    null,
                    "Authorization number missing - message archived without business processing");

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
            MailFolderCollectionResponse res = ExecuteGraphWithRetry(() => graphService.Users[mailboxAddress].MailFolders["inbox"].ChildFolders.GetAsync(q =>
            {
                q.QueryParameters.Top = 100;
                q.QueryParameters.Filter = "displayName eq '" + EscapeODataString(name) + "'";
            }
            ).GetAwaiter().GetResult(), "Find folder " + name);
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
            if (string.IsNullOrWhiteSpace(sql_connexion) || string.IsNullOrWhiteSpace(sql_gestion_cdes) || string.IsNullOrWhiteSpace(sql_dss_copie) || string.IsNullOrWhiteSpace(email_in_case_of_technical_issue) || string.IsNullOrWhiteSpace(fr_graph_send_as)) throw new InvalidOperationException("Country configuration is incomplete");
            if (credit_managers_contentieux_list.Count == 0 || BuildRecipients(contentieux_recipients).Count == 0) throw new InvalidOperationException("Contentieux configuration is invalid");
        }

        private void ValidateMailboxConfiguration()
        {
            if (mailboxId <= 0 || string.IsNullOrWhiteSpace(mailboxAddress) || string.IsNullOrWhiteSpace(inputFolderName) || string.IsNullOrWhiteSpace(outputFolderName) || allowedSenders.Count == 0) throw new InvalidOperationException("Mailbox configuration is incomplete");
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
