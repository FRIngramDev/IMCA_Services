using Microsoft.Graph;
using Microsoft.Graph.Models;
using MimeKit;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;

namespace APPLE_EMAILS_MANAGEMENT_FR
{
    public class APPLE_EMAILS_MANAGEMENT_FR
    {
        private string country = "";
        private string sk_valid = "";
        private string name = "";
        private string active = "";
        private string debug = "";
        private string start_date_scan = "";
        private string number_of_mails = "100";
        private string Rep_Easystock = "";

        private int id_mailboxe;
        private string nom_mailboxe = "";
        private string sharedmailbox_name = "";
        private DateTime dt_heure_filtre = new DateTime(1900, 1, 1);
        private int raffraichissement_min;
        private DateTime date_dernier_raf = new DateTime(1900, 1, 1);
        private bool is_integration_auto;
        private string typologie = "";
        private string sharedmailbox_folder_in = "";
        private string sharedmailbox_folder_out = "";
        private const string sharedmailbox_folder_error = "Erreur";

        private string email_in_case_of_technical_issue_parameter_global = "";
        private string email_in_case_of_technical_issue = "";
        private string sql_connexion_parameter_global = "";
        private string sql_connexion = "";
        private string logs_folder = "";
        private string temp_folder = "";
        private string global_session_name = "";
        private const string global_application_name = "APPLE_EMAILS_MANAGEMENT_FR";
        private GraphServiceClient graphService;

        public class JSON_file
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
            public string email_in_case_of_technical_issue_parameter_global { get; set; } = "";
            public string sql_connexion_parameter_global { get; set; } = "";
            public string Rep_Easystock { get; set; } = "";
        }

        private sealed class SharedMailboxConfiguration
        {
            public int IdMailboxe { get; set; }
            public string NomMailboxe { get; set; } = "";
            public string Mailboxe { get; set; } = "";
            public DateTime DtHeureFiltre { get; set; }
            public int Ordre { get; set; }
            public int RaffraichissementMin { get; set; }
            public DateTime DateDernierRaf { get; set; }
            public bool IsIntegrationAuto { get; set; }
            public string Typologie { get; set; } = "";
        }

        private sealed class AppleMailToSend
        {
            public int Id { get; set; }
            public string Sujet { get; set; } = "";
            public string Body { get; set; } = "";
            public string AffecteA { get; set; } = "";
            public byte[] Fichier { get; set; }
        }

        private sealed class TechnicalAlertAlreadySentException : Exception
        {
            public TechnicalAlertAlreadySentException(string message, Exception inner)
                : base(message, inner) { }
        }

        private GraphServiceClient Connexion_Microsoft_Graph()
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            return new class_dev_tools.Ews_Modern_Auth().Get_Graph_Service();
        }

        public void Read_Email_with_Graph(
            string sql_con,
            string logs,
            string tmp_folder,
            string session_name)
        {
            string servicePath = GetServicePath();
            setlogs_folder(Path.Combine(servicePath, logs ?? ""));
            settemp_folder(Path.Combine(servicePath, tmp_folder ?? ""));
            setGlobalSessionName(session_name);

            try
            {
                string json = get_IMCA_paramters(
                    sql_con,
                    global_application_name);

                if (string.IsNullOrWhiteSpace(json))
                {
                    throw new InvalidOperationException(
                        "No parameters found for " +
                        global_application_name);
                }

                JSON_file parameters =
                    JsonConvert.DeserializeObject<JSON_file>(json);

                if (parameters?.countries == null ||
                    parameters.countries.Count == 0)
                {
                    throw new InvalidOperationException(
                        global_application_name +
                        " parameters are empty or invalid");
                }

                foreach (Country p in parameters.countries)
                {
                    setParamCountry(p.country);
                    setParamSK_Valid(p.sk_valid);
                    setParamName(p.name);
                    setParamActive(p.active);
                    setParamDebug(p.debug);
                    setStartDateScan(p.start_date_scan);
                    setNumber_of_mails(p.number_of_mails);
                    setsharedmailbox_folder_in(
                        p.sharedmailbox_folder_in);
                    setsharedmailbox_folder_out(
                        p.sharedmailbox_folder_out);
                    setRep_Easystock(p.Rep_Easystock);
                    setSqlConnexionParam(
                        p.sql_connexion_parameter_global);
                    setEmailInCaseOfTechnicalIssueParam(
                        p.email_in_case_of_technical_issue_parameter_global);
                    setSqlConnexion(
                        get_IMCA_paramters(
                            sql_con,
                            sql_connexion_parameter_global));
                    setEmailInCaseOfTechnicalIssue(
                        get_IMCA_paramters(
                            sql_con,
                            email_in_case_of_technical_issue_parameter_global));

                    if (!IsTrue(active))
                    {
                        continue;
                    }

                    try
                    {
                        WriteToFile(
                            name.ToUpperInvariant() +
                            "(" + country.ToUpperInvariant() + ")" +
                            " at " +
                            DateTime.Now.ToString(
                                "dd/MM/yyyy HH:mm:ss"));

                        WriteToFile(
                            "   Debug Parameter is set to " +
                            debug.ToUpperInvariant());

                        ValidateRequiredCountryParameters();
                        graphService = Connexion_Microsoft_Graph();

                        if (!Directory.Exists(temp_folder))
                        {
                            Directory.CreateDirectory(temp_folder);
                        }

                        List<SharedMailboxConfiguration> sharedMailboxes =
                            GetActiveSharedMailboxes();

                        if (sharedMailboxes.Count == 0)
                        {
                            WriteToFile(
                                "   No active shared mailbox found in " +
                                "T_SharedMailboxes");
                            continue;
                        }

                        List<Exception> mailboxErrors =
                            new List<Exception>();

                        foreach (SharedMailboxConfiguration mailbox
                            in sharedMailboxes)
                        {
                            SetCurrentMailbox(mailbox);

                            if (!IsRefreshDue(
                                date_dernier_raf,
                                raffraichissement_min))
                            {
                                if (IsTrue(debug))
                                {
                                    WriteToFile(
                                        "   Mailbox skipped because refresh " +
                                        "is not due : " + nom_mailboxe +
                                        " - Last refresh : " +
                                        date_dernier_raf.ToString(
                                            "dd/MM/yyyy HH:mm:ss") +
                                        " - Refresh interval : " +
                                        raffraichissement_min +
                                        " minute(s)");
                                }
                                continue;
                            }

                            try
                            {
                                WriteToFile(
                                    "   Processing mailbox " +
                                    nom_mailboxe + " : " +
                                    sharedmailbox_name +
                                    " - Order : " + mailbox.Ordre);

                                ReadCurrentSharedMailbox();
                                UpdateMailboxRefreshInformation(
                                    id_mailboxe);
                            }
                            catch (Exception mailboxException)
                            {
                                WriteToFile(
                                    "   Error reading mailbox " +
                                    sharedmailbox_name + " : " +
                                    mailboxException.Message);

                                mailboxErrors.Add(
                                    new Exception(
                                        "Mailbox " + sharedmailbox_name +
                                        " : " + mailboxException.Message,
                                        mailboxException));

                                if (!(mailboxException is
                                    TechnicalAlertAlreadySentException))
                                {
                                    SendTechnicalIssueMail(
                                        nameof(Read_Email_with_Graph),
                                        sharedmailbox_name + " - " +
                                        mailboxException.Message,
                                        "MAILBOX PROCESSING");
                                }
                            }
                        }

                        try
                        {
                            Transfert_de_mail_communications_APPLE();
                        }
                        catch (Exception appleException)
                        {
                            mailboxErrors.Add(appleException);
                        }

                        if (mailboxErrors.Count > 0)
                        {
                            throw new AggregateException(
                                mailboxErrors.Count +
                                " APPLE_EMAILS processing error(s).",
                                mailboxErrors);
                        }
                    }
                    catch (Exception ex)
                    {
                        WriteToFile(
                            "   Error get emails : " + ex.Message);
                        throw;
                    }
                }
            }
            catch (Exception ex)
            {
                WriteToFile(
                    "Global error Read_Email_with_Graph : " +
                    ex.Message);
                throw;
            }
            finally
            {
                graphService = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void ReadCurrentSharedMailbox()
        {
            ValidateRequiredMailboxParameters();
            int nbMailInserted = 0;
            int nbMailAlreadyPresent = 0;
            int nbMailArchived = 0;

            if (IsTrue(debug))
            {
                WriteToFile(
                    "   Connexion to " + sharedmailbox_name);
                WriteToFile(
                    "   Sharedmailbox_folder_in  : " +
                    sharedmailbox_folder_in);
                WriteToFile(
                    "   Sharedmailbox_folder_out : " +
                    sharedmailbox_folder_out);
                WriteToFile(
                    "   Extracting the " + number_of_mails +
                    " oldest messages");
                WriteToFile(
                    "   Filter date : " +
                    dt_heure_filtre.ToString(
                        "dd/MM/yyyy HH:mm:ss"));
            }

            MailFolder inputFolder = GetInputFolder();
            MailFolder archiveFolder = null;

            if (IsCompubaseMailbox())
            {
                archiveFolder = GetChildFolderByName(
                    sharedmailbox_folder_out);
            }

            MailFolder errorFolder = GetChildFolderByName(
                sharedmailbox_folder_error);

            MessageCollectionResponse messages =
                GetMessagesToProcess(inputFolder.Id);

            if (messages?.Value == null ||
                messages.Value.Count == 0)
            {
                if (IsTrue(debug))
                {
                    WriteToFile("       No emails found");
                }
                return;
            }

            foreach (Message emailSummary in messages.Value)
            {
                try
                {
                    Message email =
                        GetCompleteMessage(emailSummary.Id);

                    WriteToFile(
                        "       Subject : " +
                        (email.Subject ?? "<no subject>"));

                    if (IsCompubaseMailbox())
                    {
                        ProcessCompubaseEmail(email);
                        MarkEmailAsRead(email.Id);
                        MoveEmail(email.Id, archiveFolder.Id);
                        nbMailArchived++;
                        WriteToFile(
                            "       Compubase email archived");
                    }
                    else
                    {
                        byte[] mime = GetMimeContent(email.Id);
                        int idContenuEmail =
                            InsertEmailIfNew(email, mime);

                        if (idContenuEmail > 0)
                        {
                            nbMailInserted++;
                            WriteToFile(
                                "       Email inserted in " +
                                "T_contenu_email. ID : " +
                                idContenuEmail);
                        }
                        else
                        {
                            nbMailAlreadyPresent++;
                            WriteToFile(
                                "       Email already present in " +
                                "T_contenu_email");
                        }
                    }

                    System.Threading.Thread.Sleep(500);
                }
                catch (Exception ex)
                {
                    WriteToFile(
                        "       Error processing email : " +
                        ex.Message);

                    if (!(ex is
                        TechnicalAlertAlreadySentException))
                    {
                        SendTechnicalIssueMail(
                            nameof(ReadCurrentSharedMailbox),
                            sharedmailbox_name + " - Mail : " +
                            (emailSummary.Subject ?? "<no subject>") +
                            " - " + ex.Message,
                            "EMAIL PROCESSING");
                    }

                    try
                    {
                        MarkEmailAsRead(emailSummary.Id);
                        MoveEmail(
                            emailSummary.Id,
                            errorFolder.Id);
                    }
                    catch (Exception moveException)
                    {
                        WriteToFile(
                            "       Error moving email to Erreur " +
                            "folder : " + moveException.Message);
                    }
                }
            }

            if (IsTrue(debug))
            {
                WriteToFile(
                    "       Inserted : " + nbMailInserted +
                    " - Already present : " +
                    nbMailAlreadyPresent +
                    " - Archived Compubase : " +
                    nbMailArchived);
            }
        }

        private bool IsCompubaseMailbox()
        {
            return string.Equals(sharedmailbox_name?.Trim(),
                "fr_compubase@ingrammicro.com", StringComparison.OrdinalIgnoreCase);
        }

        private void ProcessCompubaseEmail(Message email)
        {
            string subject = email.Subject ?? "";
            if (subject.IndexOf("compubase ingram stock report", StringComparison.OrdinalIgnoreCase) < 0)
            {
                WriteToFile("No Compubase processing matched. Message will be archived.");
                return;
            }

            if (string.IsNullOrWhiteSpace(Rep_Easystock))
                throw new InvalidOperationException("Rep_Easystock is empty in JSON parameters");

            Directory.CreateDirectory(Rep_Easystock);
            string destination = DownloadMatchingAttachment(email.Id, "stockreport_ingram", Rep_Easystock);
            WriteToFile("Compubase attachment copied to " + destination);
        }

        private string DownloadMatchingAttachment(string messageId, string prefix, string destinationFolder)
        {
            AttachmentCollectionResponse response = graphService.Users[sharedmailbox_name]
                .Messages[messageId].Attachments.GetAsync().GetAwaiter().GetResult();

            foreach (Microsoft.Graph.Models.Attachment attachment in response?.Value ?? new List<Microsoft.Graph.Models.Attachment>())
            {
                Microsoft.Graph.Models.FileAttachment file = attachment as Microsoft.Graph.Models.FileAttachment;
                if (file == null || !(file.Name ?? "").StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) continue;

                if (file.ContentBytes == null && !string.IsNullOrWhiteSpace(file.Id))
                    file = graphService.Users[sharedmailbox_name].Messages[messageId]
                        .Attachments[file.Id].GetAsync().GetAwaiter().GetResult()
                        as Microsoft.Graph.Models.FileAttachment;

                if (file?.ContentBytes == null)
                    throw new InvalidOperationException("Attachment content is empty : " + attachment.Name);

                string path = Path.Combine(destinationFolder, CleanFileName(file.Name));
                if (File.Exists(path)) File.Delete(path);
                File.WriteAllBytes(path, file.ContentBytes);
                return path;
            }

            throw new InvalidOperationException("No attachment starting with '" + prefix + "' was found");
        }

        private int InsertEmailIfNew(Message email, byte[] mime)
        {
            if (EmailAlreadyExists(email.Id, id_mailboxe)) return 0;

            using (SqlConnection con = new SqlConnection(sql_connexion))
            using (SqlCommand cmd = new SqlCommand("dbo.USP_ADD_EMAIL_IN_DB", con))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = 300;
                cmd.Parameters.Add("@id_equipe", SqlDbType.Int).Value = 0;
                cmd.Parameters.Add("@EwsID", SqlDbType.NVarChar, -1).Value = email.Id ?? "";
                cmd.Parameters.Add("@id_boite_mail", SqlDbType.Int).Value = id_mailboxe;
                cmd.Parameters.Add("@body", SqlDbType.NVarChar, -1).Value = GetTextBody(email);
                cmd.Parameters.Add("@from", SqlDbType.NVarChar, -1).Value = email.From?.EmailAddress?.Address ?? email.From?.EmailAddress?.Name ?? "";
                cmd.Parameters.Add("@sujet", SqlDbType.NVarChar, -1).Value = email.Subject ?? "";
                cmd.Parameters.Add("@to", SqlDbType.NVarChar, -1).Value = GetRecipients(email.ToRecipients);
                cmd.Parameters.Add("@cc", SqlDbType.NVarChar, -1).Value = GetRecipients(email.CcRecipients);
                cmd.Parameters.Add("@new_to", SqlDbType.NVarChar, -1).Value = "";
                cmd.Parameters.Add("@new_cc", SqlDbType.NVarChar, -1).Value = "";
                cmd.Parameters.Add("@new_from", SqlDbType.NVarChar, -1).Value = "";
                cmd.Parameters.Add("@new_from_nom", SqlDbType.NVarChar, -1).Value = "";
                cmd.Parameters.Add("@new_extra_body", SqlDbType.NVarChar, -1).Value = "";
                cmd.Parameters.Add("@top_a_envoyer", SqlDbType.NVarChar, 10).Value = "";
                cmd.Parameters.Add("@sent_element", SqlDbType.Bit).Value = false;

                string nomFichier = CleanFileName(
                    string.IsNullOrWhiteSpace(email.Subject)
                        ? "email"
                        : email.Subject);

                if (nomFichier.Length > 255)
                {
                    nomFichier = nomFichier.Substring(0, 255);
                }

                cmd.Parameters.Add("@nom_fic", SqlDbType.NVarChar, 255).Value = nomFichier;
                cmd.Parameters.Add("@extension", SqlDbType.NVarChar, 10).Value = "eml";
                cmd.Parameters.Add("@document", SqlDbType.Image).Value = mime ?? new byte[0];
                cmd.Parameters.Add("@affecte_a", SqlDbType.NVarChar, -1).Value = "";
                cmd.Parameters.Add("@dt_time_received", SqlDbType.DateTime).Value = email.ReceivedDateTime?.LocalDateTime ?? email.CreatedDateTime?.LocalDateTime ?? DateTime.Now;
                cmd.Parameters.Add("@hasAttachement", SqlDbType.Bit).Value = email.HasAttachments == true;
                cmd.Parameters.Add("@conversationId", SqlDbType.NVarChar, -1).Value = email.ConversationId ?? "";
                SqlParameter output = cmd.Parameters.Add("@@id", SqlDbType.Int);
                output.Direction = ParameterDirection.Output;
                con.Open();
                cmd.ExecuteNonQuery();
                return output.Value == DBNull.Value ? 0 : Convert.ToInt32(output.Value);
            }
        }

        private bool EmailAlreadyExists(string graphId, int mailboxId)
        {
            const string sql = @"SELECT TOP (1) id FROM dbo.T_contenu_email
WHERE EwsID COLLATE Latin1_General_CS_AS = @EWSID AND id_mailboxe = @ID;";
            using (SqlConnection con = new SqlConnection(sql_connexion))
            using (SqlCommand cmd = new SqlCommand(sql, con))
            {
                cmd.Parameters.Add("@EWSID", SqlDbType.NVarChar, -1).Value = graphId ?? "";
                cmd.Parameters.Add("@ID", SqlDbType.Int).Value = mailboxId;
                con.Open();
                return cmd.ExecuteScalar() != null;
            }
        }

        public void Transfert_de_mail_communications_APPLE()
        {
            WriteToFile(
                "   Starting Apple communication processing");

            SharedMailboxConfiguration appleMailbox =
                GetAppleMailbox();

            if (appleMailbox == null)
            {
                WriteToFile(
                    "   No active mailbox with typologie APPLE found");
                return;
            }

            SetCurrentMailbox(appleMailbox);
            AppleMailToSend mail = GetNextAppleMailToSend();

            if (mail == null)
            {
                if (IsTrue(debug))
                {
                    WriteToFile(
                        "   No Apple communication to send");
                }
                return;
            }

            WriteToFile(
                "   Processing Apple communication ID : " +
                mail.Id + " - Group : " + mail.AffecteA);

            try
            {
                UpdateAppleValue(
                    mail.Id,
                    "top_a_envoyer",
                    "P");

                List<string> contacts =
                    GetAppleContacts(mail.AffecteA);

                if (contacts.Count == 0)
                {
                    throw new InvalidOperationException(
                        "No Apple contact found for group : " +
                        mail.AffecteA);
                }

                WriteToFile(
                    "   Apple contacts found : " + contacts.Count);

                MimeMessage source =
                    LoadMimeMessage(mail.Fichier);
                List<Microsoft.Graph.Models.Attachment> attachments =
                    BuildGraphAttachments(source);
                string sentTo = "";

                foreach (string contact in contacts)
                {
                    Message message = new Message
                    {
                        Subject =
                            !string.IsNullOrWhiteSpace(source.Subject)
                                ? source.Subject
                                : mail.Sujet,
                        Body = BuildAppleBody(source, mail),
                        ToRecipients = BuildRecipients(contact),
                        ReplyTo = BuildRecipients(sharedmailbox_name),
                        Attachments = CloneAttachments(attachments)
                    };

                    var request =
                        new Microsoft.Graph.Users.Item.SendMail
                            .SendMailPostRequestBody
                        {
                            Message = message,
                            SaveToSentItems = true
                        };

                    graphService.Users[sharedmailbox_name]
                        .SendMail
                        .PostAsync(request)
                        .GetAwaiter()
                        .GetResult();

                    sentTo = string.IsNullOrWhiteSpace(sentTo)
                        ? contact
                        : sentTo + ";" + contact;

                    UpdateAppleValue(
                        mail.Id,
                        "envoyer_a",
                        sentTo);

                    WriteToFile(
                        "       Apple communication sent to : " +
                        contact);
                }

                UpdateAppleValue(
                    mail.Id,
                    "top_a_envoyer",
                    "F");

                WriteToFile(
                    "   Apple communication ID " + mail.Id +
                    " successfully sent to " + contacts.Count +
                    " contact(s)");
            }
            catch (Exception ex)
            {
                try
                {
                    UpdateAppleValue(
                        mail.Id,
                        "top_a_envoyer",
                        "E");
                }
                catch
                {
                }

                SendTechnicalIssueMail(
                    nameof(Transfert_de_mail_communications_APPLE),
                    "Apple mail ID " + mail.Id + " - " +
                    ex.Message,
                    "APPLE");

                throw new TechnicalAlertAlreadySentException(
                    "Apple communication failed : " + ex.Message,
                    ex);
            }
        }

        private SharedMailboxConfiguration GetAppleMailbox()
        {
            const string sql = @"SELECT TOP (1) id_mailboxe, nom_mailboxe, mailboxe, ordre,
ISNULL(dt_heure_filtre,'19000101') dt_heure_filtre,
ISNULL(raffraichissement_min,0) raffraichissement_min,
ISNULL(date_dernier_raf,'19000101') date_dernier_raf,
ISNULL(is_integration_auto,0) is_integration_auto,
ISNULL(typologie,'') typologie
FROM dbo.T_SharedMailboxes
WHERE actif=1 AND UPPER(LTRIM(RTRIM(ISNULL(typologie,''))))='APPLE'
ORDER BY ordre;";
            return ExecuteMailboxQuery(sql).FirstOrDefault();
        }

        private AppleMailToSend GetNextAppleMailToSend()
        {
            const string sql = @"
SELECT TOP (1)
    c.id,
    ISNULL(c.sujet, '') AS sujet,
    ISNULL(c.body, '') AS body,
    ISNULL(c.affecte_a, '') AS affecte_a,
    c.fichier
FROM dbo.T_contenu_email AS c WITH (NOLOCK)
WHERE c.id_mailboxe = @ID_MAILBOXE
  AND c.top_a_envoyer = 'O'
  AND c.affecte_a IS NOT NULL
  AND LTRIM(RTRIM(c.affecte_a)) <> ''
ORDER BY c.id;";

            using (SqlConnection con =
                new SqlConnection(sql_connexion))
            using (SqlCommand cmd =
                new SqlCommand(sql, con))
            {
                cmd.CommandTimeout = 300;
                cmd.Parameters.Add(
                    "@ID_MAILBOXE",
                    SqlDbType.Int).Value = id_mailboxe;

                con.Open();

                using (SqlDataReader reader =
                    cmd.ExecuteReader(CommandBehavior.SingleRow))
                {
                    if (!reader.Read())
                    {
                        return null;
                    }

                    return new AppleMailToSend
                    {
                        Id = Convert.ToInt32(reader["id"]),
                        Sujet = Convert.ToString(reader["sujet"]),
                        Body = Convert.ToString(reader["body"]),
                        AffecteA = Convert.ToString(reader["affecte_a"]),
                        Fichier = reader["fichier"] == DBNull.Value
                            ? null
                            : (byte[])reader["fichier"]
                    };
                }
            }
        }

        private List<string> GetAppleContacts(string group)
        {
            const string sql = @"SELECT DISTINCT LTRIM(RTRIM(email)) email
FROM dbo.T_Contacts_Communications_APPLE
WHERE nom_groupe=@GROUP AND NULLIF(LTRIM(RTRIM(email)),'') IS NOT NULL
ORDER BY email;";
            List<string> contacts = new List<string>();
            using (SqlConnection con = new SqlConnection(sql_connexion))
            using (SqlCommand cmd = new SqlCommand(sql, con))
            {
                cmd.Parameters.Add("@GROUP", SqlDbType.NVarChar, 255).Value = group;
                con.Open();
                using (SqlDataReader r = cmd.ExecuteReader())
                    while (r.Read()) contacts.Add(Convert.ToString(r["email"]));
            }
            return contacts;
        }

        private void UpdateAppleValue(int id, string column, string value)
        {
            if (column != "top_a_envoyer" && column != "envoyer_a")
                throw new InvalidOperationException("Invalid Apple update column");

            string sql = "UPDATE dbo.T_contenu_email SET [" + column + "]=@VALUE WHERE id=@ID;";
            using (SqlConnection con = new SqlConnection(sql_connexion))
            using (SqlCommand cmd = new SqlCommand(sql, con))
            {
                cmd.Parameters.Add("@VALUE", SqlDbType.NVarChar, -1).Value = value ?? "";
                cmd.Parameters.Add("@ID", SqlDbType.Int).Value = id;
                con.Open();
                cmd.ExecuteNonQuery();
            }
        }

        private static MimeMessage LoadMimeMessage(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0)
                throw new InvalidOperationException("The stored MIME content is empty");
            using (MemoryStream stream = new MemoryStream(bytes, false))
                return MimeMessage.Load(stream);
        }

        private static ItemBody BuildAppleBody(
            MimeMessage source,
            AppleMailToSend mail)
        {
            if (!string.IsNullOrWhiteSpace(source.HtmlBody))
            {
                return new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = source.HtmlBody
                };
            }

            if (!string.IsNullOrWhiteSpace(source.TextBody))
            {
                return new ItemBody
                {
                    ContentType = BodyType.Text,
                    Content = source.TextBody
                };
            }

            return new ItemBody
            {
                ContentType = BodyType.Text,
                Content = mail.Body ?? ""
            };
        }

        private static List<Microsoft.Graph.Models.Attachment> BuildGraphAttachments(MimeMessage source)
        {
            List<Microsoft.Graph.Models.Attachment> result = new List<Microsoft.Graph.Models.Attachment>();
            foreach (MimeEntity entity in source.BodyParts)
            {
                MimePart part = entity as MimePart;
                if (part == null || (!part.IsAttachment && string.IsNullOrWhiteSpace(part.ContentId))) continue;
                using (MemoryStream stream = new MemoryStream())
                {
                    part.Content.DecodeTo(stream);
                    result.Add(new Microsoft.Graph.Models.FileAttachment
                    {
                        OdataType = "#microsoft.graph.fileAttachment",
                        Name = string.IsNullOrWhiteSpace(part.FileName) ? "attachment" : part.FileName,
                        ContentType = part.ContentType.MimeType,
                        ContentBytes = stream.ToArray(),
                        ContentId = part.ContentId,
                        IsInline = !part.IsAttachment
                    });
                }
            }
            return result;
        }

        private static List<Microsoft.Graph.Models.Attachment> CloneAttachments(IEnumerable<Microsoft.Graph.Models.Attachment> source)
        {
            return source.Cast<Microsoft.Graph.Models.FileAttachment>()
                .Select(a => (Microsoft.Graph.Models.Attachment)new Microsoft.Graph.Models.FileAttachment
                {
                    OdataType = "#microsoft.graph.fileAttachment",
                    Name = a.Name,
                    ContentType = a.ContentType,
                    ContentBytes = a.ContentBytes,
                    ContentId = a.ContentId,
                    IsInline = a.IsInline
                }).ToList();
        }

        private List<SharedMailboxConfiguration> GetActiveSharedMailboxes()
        {
            const string sql = @"SELECT id_mailboxe, nom_mailboxe, mailboxe, ordre,
ISNULL(dt_heure_filtre,'19000101') dt_heure_filtre,
ISNULL(raffraichissement_min,0) raffraichissement_min,
ISNULL(date_dernier_raf,'19000101') date_dernier_raf,
ISNULL(is_integration_auto,0) is_integration_auto,
ISNULL(typologie,'') typologie
FROM dbo.T_SharedMailboxes
WHERE actif=1
AND LOWER(LTRIM(RTRIM(mailboxe))) NOT IN
('fr_aging_report@ingrammicro.com','marketplace_data_fr@ingrammicro.com')
ORDER BY ordre;";
            return ExecuteMailboxQuery(sql);
        }

        private List<SharedMailboxConfiguration> ExecuteMailboxQuery(string sql)
        {
            List<SharedMailboxConfiguration> result = new List<SharedMailboxConfiguration>();
            using (SqlConnection con = new SqlConnection(sql_connexion))
            using (SqlCommand cmd = new SqlCommand(sql, con))
            {
                con.Open();
                using (SqlDataReader r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        result.Add(new SharedMailboxConfiguration
                        {
                            IdMailboxe = Convert.ToInt32(r["id_mailboxe"]),
                            NomMailboxe = Convert.ToString(r["nom_mailboxe"]).Trim(),
                            Mailboxe = Convert.ToString(r["mailboxe"]).Trim(),
                            Ordre = Convert.ToInt32(r["ordre"]),
                            DtHeureFiltre = Convert.ToDateTime(r["dt_heure_filtre"]),
                            RaffraichissementMin = Convert.ToInt32(r["raffraichissement_min"]),
                            DateDernierRaf = Convert.ToDateTime(r["date_dernier_raf"]),
                            IsIntegrationAuto = Convert.ToBoolean(r["is_integration_auto"]),
                            Typologie = Convert.ToString(r["typologie"]).Trim()
                        });
                    }
                }
            }
            return result;
        }

        private void SetCurrentMailbox(SharedMailboxConfiguration m)
        {
            setIdMailboxe(m.IdMailboxe);
            setNomMailboxe(m.NomMailboxe);
            setSharedMailboxName(m.Mailboxe);
            setDateHeureFiltre(m.DtHeureFiltre);
            setRaffraichissementMin(m.RaffraichissementMin);
            setDateDernierRaf(m.DateDernierRaf);
            setIntegrationAuto(m.IsIntegrationAuto);
            setTypologie(m.Typologie);
        }

        private void UpdateMailboxRefreshInformation(int id)
        {
            const string sql = @"UPDATE dbo.T_SharedMailboxes SET date_dernier_raf=GETDATE(),
dt_heure_filtre=COALESCE((SELECT MAX(dt_time_received) FROM dbo.T_contenu_email WHERE id_mailboxe=@ID),dt_heure_filtre)
WHERE id_mailboxe=@ID;";
            using (SqlConnection con = new SqlConnection(sql_connexion))
            using (SqlCommand cmd = new SqlCommand(sql, con))
            {
                cmd.Parameters.Add("@ID", SqlDbType.Int).Value = id;
                con.Open();
                cmd.ExecuteNonQuery();
            }
        }

        private MessageCollectionResponse GetMessagesToProcess(string folderId)
        {
            int top;
            if (!int.TryParse(number_of_mails, out top) || top <= 0) top = 100;
            DateTime filter = dt_heure_filtre > new DateTime(1900, 1, 1) ? dt_heure_filtre : ParseStartDate();
            return graphService.Users[sharedmailbox_name].MailFolders[folderId].Messages.GetAsync(c =>
            {
                c.QueryParameters.Top = top;
                c.QueryParameters.Orderby = new[] { "receivedDateTime asc" };
                c.QueryParameters.Filter = "receivedDateTime ge " + filter.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);
                c.QueryParameters.Select = new[] { "id", "subject", "receivedDateTime", "createdDateTime", "hasAttachments" };
            }).GetAwaiter().GetResult();
        }

        private Message GetCompleteMessage(string id)
        {
            return graphService.Users[sharedmailbox_name].Messages[id].GetAsync(c =>
            {
                c.QueryParameters.Select = new[] { "id", "subject", "body", "bodyPreview", "from", "toRecipients", "ccRecipients", "receivedDateTime", "createdDateTime", "conversationId", "hasAttachments" };
            }).GetAwaiter().GetResult();
        }

        private byte[] GetMimeContent(string id)
        {
            using (Stream stream = graphService.Users[sharedmailbox_name].Messages[id].Content.GetAsync().GetAwaiter().GetResult())
            using (MemoryStream memory = new MemoryStream())
            {
                stream.CopyTo(memory);
                return memory.ToArray();
            }
        }

        private MailFolder GetInputFolder()
        {
            return string.Equals(sharedmailbox_folder_in, "Inbox", StringComparison.OrdinalIgnoreCase)
                ? graphService.Users[sharedmailbox_name].MailFolders["inbox"].GetAsync().GetAwaiter().GetResult()
                : GetChildFolderByName(sharedmailbox_folder_in);
        }

        private MailFolder GetChildFolderByName(string nameFolder)
        {
            string safe = (nameFolder ?? "").Replace("'", "''");
            MailFolderCollectionResponse folders = graphService.Users[sharedmailbox_name].MailFolders["inbox"].ChildFolders
                .GetAsync(c => c.QueryParameters.Filter = "displayName eq '" + safe + "'").GetAwaiter().GetResult();
            MailFolder folder = folders?.Value?.FirstOrDefault();
            if (folder == null) throw new DirectoryNotFoundException("Folder not found : " + nameFolder);
            return folder;
        }

        private void MarkEmailAsRead(string id)
        {
            graphService.Users[sharedmailbox_name].Messages[id].PatchAsync(new Message { IsRead = true }).GetAwaiter().GetResult();
        }

        private void MoveEmail(string id, string destinationId)
        {
            var request = new Microsoft.Graph.Users.Item.Messages.Item.Move.MovePostRequestBody { DestinationId = destinationId };
            graphService.Users[sharedmailbox_name].Messages[id].Move.PostAsync(request).GetAwaiter().GetResult();
        }

        private void SendTechnicalIssueMail(string method, string error, string type)
        {
            try
            {
                WriteToFile("       " + type + " error in " + method + " : " + error);
                if (graphService == null || string.IsNullOrWhiteSpace(email_in_case_of_technical_issue) || string.IsNullOrWhiteSpace(sharedmailbox_name)) return;
                Message message = new Message
                {
                    Subject = global_application_name + " - Erreur " + type + " dans " + method,
                    Body = new ItemBody { ContentType = BodyType.Html, Content = "<b>Mailbox :</b> " + WebUtility.HtmlEncode(sharedmailbox_name) + "<br/><b>Message :</b> " + WebUtility.HtmlEncode(error) },
                    ToRecipients = BuildRecipients(email_in_case_of_technical_issue)
                };
                var request = new Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody { Message = message, SaveToSentItems = true };
                graphService.Users[sharedmailbox_name].SendMail.PostAsync(request).GetAwaiter().GetResult();
            }
            catch (Exception ex) { WriteToFile("Error sending technical alert : " + ex.Message); }
        }

        private static List<Recipient> BuildRecipients(string addresses)
        {
            return (addresses ?? "").Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(a => a.Trim()).Where(a => a.Contains("@")).Distinct(StringComparer.OrdinalIgnoreCase)
                .Select(a => new Recipient { EmailAddress = new EmailAddress { Address = a } }).ToList();
        }

        private static string GetRecipients(IEnumerable<Recipient> recipients)
        {
            return recipients == null ? "" : string.Join(";", recipients.Select(r => r?.EmailAddress?.Address).Where(a => !string.IsNullOrWhiteSpace(a)));
        }

        private static string GetTextBody(Message email) { return email.Body?.ContentType == BodyType.Text ? email.Body.Content ?? "" : email.BodyPreview ?? ""; }
        private static string GetHtmlBody(Message email) { return email.Body?.ContentType == BodyType.Html ? email.Body.Content ?? "" : ""; }

        private DateTime ParseStartDate()
        {
            if (string.IsNullOrWhiteSpace(start_date_scan)) return new DateTime(1900, 1, 1);
            DateTime date;
            if (!DateTime.TryParseExact(start_date_scan, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                throw new InvalidOperationException("Invalid start_date_scan value. Expected yyyyMMdd");
            return date;
        }

        private string get_IMCA_paramters(string cs, string parameter)
        {
            const string sql = "SELECT ISNULL(VALUE,'') FROM PCM_TAB_IMCA_PARAMETER_GLOBAL WHERE SK_VALID=0 AND PARAMETER=@PARAMETER;";
            using (SqlConnection con = new SqlConnection(cs))
            using (SqlCommand cmd = new SqlCommand(sql, con))
            {
                cmd.Parameters.Add("@PARAMETER", SqlDbType.NVarChar, 255).Value = parameter;
                con.Open();
                return Convert.ToString(cmd.ExecuteScalar());
            }
        }

        private void ValidateRequiredCountryParameters()
        {
            if (string.IsNullOrWhiteSpace(sql_connexion)) throw new InvalidOperationException("sql_connexion is empty");
            if (string.IsNullOrWhiteSpace(temp_folder)) throw new InvalidOperationException("temp_folder is empty");
            if (string.IsNullOrWhiteSpace(Rep_Easystock)) WriteToFile("Warning: Rep_Easystock is empty. Compubase processing will fail if triggered.");
        }

        private void ValidateRequiredMailboxParameters()
        {
            if (id_mailboxe <= 0) throw new InvalidOperationException("id_mailboxe is invalid");
            if (string.IsNullOrWhiteSpace(sharedmailbox_name)) throw new InvalidOperationException("mailboxe is empty");
        }

        private void WriteToFile(string message)
        {
            if (string.IsNullOrWhiteSpace(logs_folder)) logs_folder = AppDomain.CurrentDomain.BaseDirectory;
            Directory.CreateDirectory(logs_folder);
            string file = Path.Combine(logs_folder, "IMCA_" + global_session_name + "_" + DateTime.Now.ToString("dd_MM_yyyy") + "_" + country + "_" + global_application_name + ".txt");
            File.AppendAllText(file, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - " + message + Environment.NewLine);
        }

        private static string GetServicePath()
        {
            string location = System.Reflection.Assembly.GetEntryAssembly()?.Location;
            return string.IsNullOrWhiteSpace(location) ? AppDomain.CurrentDomain.BaseDirectory : Path.GetDirectoryName(location);
        }

        private static string CleanFileName(string value)
        {
            return string.Join("_", (value ?? "attachment").Split(Path.GetInvalidFileNameChars())).Trim();
        }

        private static bool IsTrue(string value) { return string.Equals(value?.Trim(), "TRUE", StringComparison.OrdinalIgnoreCase); }
        private static bool IsRefreshDue(DateTime date, int minutes) { return minutes <= 0 || DateTime.Now >= date.AddMinutes(minutes); }

        public void setParamCountry(string value) { country = value ?? ""; }
        public void setParamSK_Valid(string value) { sk_valid = value ?? ""; }
        public void setParamName(string value) { name = value ?? ""; }
        public void setParamActive(string value) { active = value ?? ""; }
        public void setParamDebug(string value) { debug = value ?? ""; }
        public void setStartDateScan(string value) { start_date_scan = value ?? ""; }
        public void setNumber_of_mails(string value) { number_of_mails = string.IsNullOrWhiteSpace(value) ? "100" : value; }
        public void setRep_Easystock(string value) { Rep_Easystock = value ?? ""; }
        public void setIdMailboxe(int value) { id_mailboxe = value; }
        public void setNomMailboxe(string value) { nom_mailboxe = value ?? ""; }
        public void setSharedMailboxName(string value) { sharedmailbox_name = value ?? ""; }
        public void setDateHeureFiltre(DateTime value) { dt_heure_filtre = value; }
        public void setRaffraichissementMin(int value) { raffraichissement_min = value; }
        public void setDateDernierRaf(DateTime value) { date_dernier_raf = value; }
        public void setIntegrationAuto(bool value) { is_integration_auto = value; }
        public void setTypologie(string value) { typologie = value ?? ""; }
        public void setsharedmailbox_folder_in(string value) { sharedmailbox_folder_in = string.IsNullOrWhiteSpace(value) ? "Inbox" : value; }
        public void setsharedmailbox_folder_out(string value) { sharedmailbox_folder_out = string.IsNullOrWhiteSpace(value) ? "Archives" : value; }
        public void setEmailInCaseOfTechnicalIssueParam(string value) { email_in_case_of_technical_issue_parameter_global = value ?? ""; }
        public void setEmailInCaseOfTechnicalIssue(string value) { email_in_case_of_technical_issue = value ?? ""; }
        public void setSqlConnexionParam(string value) { sql_connexion_parameter_global = value ?? ""; }
        public void setSqlConnexion(string value) { sql_connexion = value ?? ""; }
        public void setlogs_folder(string value) { logs_folder = value ?? ""; }
        public void settemp_folder(string value) { temp_folder = value ?? ""; }
        public void setGlobalSessionName(string value) { global_session_name = value ?? ""; }
    }
}
