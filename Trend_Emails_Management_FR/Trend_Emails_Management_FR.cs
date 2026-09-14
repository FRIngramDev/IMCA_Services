using ExcelDataReader;
using Microsoft.Graph;
using Microsoft.Graph.Models;
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
using System.Text.RegularExpressions;

namespace TREND_EMAILS_MANAGEMENT_FR
{
    public class TREND_EMAILS_MANAGEMENT_FR
    {
        public TREND_EMAILS_MANAGEMENT_FR()
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
        private int id_mailboxe = 0;
        private string nom_mailboxe = "";
        private string sharedmailbox_name = "";
        private string sharedmailbox_folder_in = "";
        private string sharedmailbox_folder_out = "";
        private const string sharedmailbox_folder_error = "Erreur";

        // ----- Technical alert parameters -----
        private string email_in_case_of_technical_issue_parameter_global = "";
        private string email_in_case_of_technical_issue = "";

        // ----- Files / logs / session -----
        private string logs_folder = "";
        private string temp_folder = "";
        private string global_session_name = "";

        // ----- Application name -----
        private const string global_application_name =
            "TREND_EMAILS_MANAGEMENT_FR";

        // ----- SQL connection parameters -----
        private string sql_connexion = "";
        private string sql_connexion_parameter_global = "";

        // ----- Microsoft Graph client -----
        private GraphServiceClient graphService = null;

        public class JSON_file
        {
            public string logs_folder { get; set; } = "";
            public List<Country> countries { get; set; } =
                new List<Country>();
        }

        public class Country
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
            public string email_in_case_of_technical_issue_parameter_global { get; set; } = "";
            public string sql_connexion_parameter_global { get; set; } = "";
        }

        private sealed class SharedMailboxConfiguration
        {
            public int IdMailboxe { get; set; }
            public string NomMailboxe { get; set; } = "";
            public string Mailboxe { get; set; } = "";
            public int Ordre { get; set; }
        }

        private sealed class TechnicalAlertAlreadySentException : Exception
        {
            public TechnicalAlertAlreadySentException(
                string message,
                Exception innerException)
                : base(message, innerException)
            {
            }
        }

        private GraphServiceClient Connexion_Microsoft_Graph()
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            class_dev_tools.Ews_Modern_Auth Ews_Modern_Auth =
                new class_dev_tools.Ews_Modern_Auth();

            return Ews_Modern_Auth.Get_Graph_Service();
        }

        public void Read_Email_with_Graph(
            string sql_con,
            string logs,
            string tmp_folder,
            string session_name)
        {
            string global_parameters = "";
            string service_path = GetServicePath();

            setlogs_folder(Path.Combine(service_path, logs ?? ""));
            settemp_folder(Path.Combine(service_path, tmp_folder ?? ""));
            setGlobalSessionName(session_name);

            try
            {
                global_parameters = get_IMCA_paramters(
                    sql_con,
                    global_application_name);

                if (string.IsNullOrWhiteSpace(global_parameters))
                {
                    throw new InvalidOperationException(
                        "No parameters found for " +
                        global_application_name);
                }

                JSON_file param =
                    JsonConvert.DeserializeObject<JSON_file>(
                        global_parameters);

                if (param == null ||
                    param.countries == null ||
                    param.countries.Count == 0)
                {
                    throw new InvalidOperationException(
                        global_application_name +
                        " parameters are empty or invalid");
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
                    setsharedmailbox_folder_in( p.sharedmailbox_folder_in);
                    setsharedmailbox_folder_out(p.sharedmailbox_folder_out);
                    setSqlConnexionParam(p.sql_connexion_parameter_global);
                    setEmailInCaseOfTechnicalIssueParam(p.email_in_case_of_technical_issue_parameter_global);

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

                        setGraphService(Connexion_Microsoft_Graph());

                        if (!Directory.Exists(temp_folder))
                        {
                            Directory.CreateDirectory(temp_folder);
                        }

                        DeleteFiles(temp_folder, global_application_name);

                        List<SharedMailboxConfiguration> sharedMailboxes = GetActiveSharedMailboxes();

                        if (sharedMailboxes.Count == 0)
                        {
                            WriteToFile(
                                "   No active shared mailbox found in " +
                                "T_SharedMailboxes");
                            continue;
                        }

                        List<Exception> mailboxErrors = new List<Exception>();

                        foreach (
                            SharedMailboxConfiguration mailbox in sharedMailboxes)
                        {
                            setIdMailboxe(mailbox.IdMailboxe);
                            setNomMailboxe(mailbox.NomMailboxe);
                            setSharedMailboxName(mailbox.Mailboxe);

                            try
                            {
                                WriteToFile(
                                    "   Processing mailbox " +
                                    nom_mailboxe + " : " +
                                    sharedmailbox_name +
                                    " - Order : " + mailbox.Ordre);

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
                                        sharedmailbox_name + " : " +
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
                setGraphService(null);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void ReadCurrentSharedMailbox()
        {
            ValidateRequiredMailboxParameters();
            int nb_mail = 0;
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
            }

            MailFolder inputFolder = GetInputFolder();
            MailFolder archiveFolder = GetChildFolderByName(sharedmailbox_folder_out);
            MailFolder errorFolder = GetChildFolderByName(sharedmailbox_folder_error);

            MessageCollectionResponse messages = GetMessagesToProcess(inputFolder.Id);

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
                    Message email = GetCompleteMessage(emailSummary.Id);

                    WriteToFile(
                        "       Subject : " +
                        (email.Subject ?? "<no subject>"));

                    ProcessTrendMessage(email);

                    MarkEmailAsRead(email.Id);
                    MoveEmail(email.Id, archiveFolder.Id);
                    nb_mail++;
                    System.Threading.Thread.Sleep(500);
                }
                catch (Exception ex)
                {
                    WriteToFile(
                        "       Error processing email : " +
                        ex.Message);

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
                        MarkEmailAsRead(emailSummary.Id);
                        MoveEmail(emailSummary.Id, errorFolder.Id);
                    }
                    catch (Exception moveException)
                    {
                        WriteToFile(
                            "       Error moving email to Erreur folder : " +
                            moveException.Message);
                    }
                }
            }

            if (IsTrue(debug))
            {
                WriteToFile(
                    "       " + nb_mail +
                    " email(s) have been processed");
            }
        }

        private void ProcessTrendMessage(Message email)
        {
            string mailbox =
                (sharedmailbox_name ?? "")
                    .Trim()
                    .ToLowerInvariant();

            string subject =
                (email.Subject ?? "")
                    .Trim()
                    .ToLowerInvariant();

            if (mailbox == "fr_aging_report@ingrammicro.com")
            {
                if (subject.Contains(
                    "aging3d_sales_daily_proj_fin_de_mois"))
                {
                    ProcessAging3D(
                        email,
                        "Aging3d_Sales_Daily_Proj_fin_de_mois",
                        "T_AGING3D_DAILY_SALES_PROJ_FIN_DE_MOIS_PLUS_2_TMP",
                        "Sales Branch");

                    return;
                }

                if (subject.Contains("aging3d_sales_daily"))
                {
                    ProcessAging3D(
                        email,
                        "Aging3d_Sales_Daily",
                        "T_AGING3D_DAILY_SALES_BR_21_91_TMP",
                        "Sales Branch 21 Et 91");

                    return;
                }

                WriteToFile(
                    "       No Aging processing matched. " +
                    "Message will be archived.");

                return;
            }

            if (mailbox == "marketplace_data_fr@ingrammicro.com")
            {
                DateTime received =
                    email.ReceivedDateTime?.LocalDateTime ??
                    DateTime.Now;

                if (received.DayOfWeek == DayOfWeek.Saturday ||
                    received.DayOfWeek == DayOfWeek.Sunday)
                {
                    WriteToFile(
                        "       Marketplace message received during " +
                        "the weekend. Message will be archived.");

                    return;
                }

                if (subject.Contains("vendors"))
                {
                    ProcessMarketplaceVendors(email);
                    return;
                }

                if (subject.Contains("france_marketplace_last"))
                {
                    ProcessMarketplaceData(email);
                    return;
                }

                WriteToFile(
                    "       No Marketplace processing matched. " +
                    "Message will be archived.");

                return;
            }

            throw new InvalidOperationException(
                "No TREND processing configured for mailbox : " +
                sharedmailbox_name);
        }

        private void ProcessMarketplaceVendors(Message email)
        {
            string file = DownloadMatchingAttachment(
                email.Id,
                "Vendors");

            DataTable data = ReadExcelSheet(file, "sheet1");

            BulkReplaceTable(
                data,
                "cloud_liste_vendor_tmp");

            const string sql = @"
                                TRUNCATE TABLE cloud_liste_vendor;
                                INSERT INTO cloud_liste_vendor
                                    ([code_name], [code_vendor], [vendor_model])
                                SELECT [Vendor], [Vendor Code], [Vendor Model]
                                FROM cloud_liste_vendor_tmp;
                                DELETE FROM cloud_liste_vendor
                                WHERE [code_name] IS NULL;";

            ExecuteSql(
                sql,
                nameof(ProcessMarketplaceVendors));
        }

        private void ProcessMarketplaceData(Message email)
        {
            string file = DownloadMatchingAttachment(
                email.Id,
                "France_Marketplace");

            DataTable data = ReadExcelSheet(file, "sheet1");

            BulkReplaceTable(
                data,
                "marketplace_data_brute_ATLAS_new_tmp");

            const string sql = @"
                                SET LANGUAGE English;
                                UPDATE marketplace_data_brute_ATLAS_new_tmp
                                SET [Fiscal Month] = FORMAT(
                                    CAST('01-' + REPLACE([Fiscal Month], ' ', '') AS DATE),
                                    'yyyyMM');
                                DELETE FROM marketplace_data_brute_ATLAS_new
                                WHERE [Fiscal Month] IN
                                (
                                    SELECT DISTINCT [Fiscal Month]
                                    FROM marketplace_data_brute_ATLAS_new_tmp
                                );
                                INSERT INTO marketplace_data_brute_ATLAS_new
                                (
                                    [VENDOR], [VENDOR_NBR], [Billing Period Type], [BCN],
                                    [Subscription Name], [Fiscal Month], date_facturation,
                                    [Gross Sales, LC], [End-User], [Net Sales, LC]
                                )
                                SELECT
                                    [Vendor], [Vendor Code], [Billing Period Type], [BCN],
                                    [Subscription Name], [Fiscal Month], [Report Date],
                                    [Gross Sales, LC], [End-User Name], [Net Sales, LC]
                                FROM marketplace_data_brute_ATLAS_new_tmp;
                                INSERT INTO marketplace_data_brute_ATLAS_new_date_maj
                                    (date_maj, date_max_facturation)
                                SELECT GETDATE(), MAX(date_facturation)
                                FROM marketplace_data_brute_ATLAS_new;";

            ExecuteSql(
                sql,
                nameof(ProcessMarketplaceData));
        }

        private void ProcessAging3D(
            Message email,
            string attachmentPrefix,
            string destinationTable,
            string worksheetName)
        {
            string attachment = DownloadMatchingAttachment(
                email.Id,
                attachmentPrefix);

            string extractedFile = ExtractExcelFile(
                attachment,
                temp_folder);

            DataTable data = ReadExcelSheet(
                extractedFile,
                worksheetName);

            BulkReplaceTable(data, destinationTable);

            string dateMail = ExtractDateFromSubject(
                email.Subject,
                email.ReceivedDateTime?.LocalDateTime ??
                DateTime.Now);

            const string sql =
                "EXEC dbo.TREND_AGING3D_DAILY " +
                "@table_SQL, @date_mail";

            try
            {
                using (SqlConnection connection =
                    new SqlConnection(sql_connexion))
                using (SqlCommand command = new SqlCommand(
                    "dbo.TREND_AGING3D_DAILY",
                    connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandTimeout = 1800;

                    command.Parameters.Add(
                        "@table_SQL",
                        SqlDbType.NVarChar,
                        255).Value = destinationTable;

                    command.Parameters.Add(
                        "@date_mail",
                        SqlDbType.NVarChar,
                        20).Value = dateMail;

                    connection.Open();
                    command.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                SendTechnicalIssueMail(
                    nameof(ProcessAging3D),
                    sql,
                    ex.Message,
                    "SQL");

                throw new TechnicalAlertAlreadySentException(
                    "SQL error in ProcessAging3D : " + ex.Message,
                    ex);
            }
        }

        private List<SharedMailboxConfiguration>
            GetActiveSharedMailboxes()
        {
            const string sql = @"
                                SELECT
                                    id_mailboxe,
                                    nom_mailboxe,
                                    mailboxe,
                                    ordre
                                FROM dbo.T_SharedMailboxes
                                WHERE actif = 1
                                ORDER BY ordre;";

            try
            {
                List<SharedMailboxConfiguration> result =
                    new List<SharedMailboxConfiguration>();

                using (SqlConnection connection =
                    new SqlConnection(sql_connexion))
                using (SqlCommand command =
                    new SqlCommand(sql, connection))
                {
                    command.CommandTimeout = 300;
                    connection.Open();

                    using (SqlDataReader reader =
                        command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            result.Add(
                                new SharedMailboxConfiguration
                                {
                                    IdMailboxe = Convert.ToInt32(
                                        reader["id_mailboxe"]),
                                    NomMailboxe = Convert.ToString(
                                        reader["nom_mailboxe"]).Trim(),
                                    Mailboxe = Convert.ToString(
                                        reader["mailboxe"]).Trim(),
                                    Ordre = Convert.ToInt32(
                                        reader["ordre"])
                                });
                        }
                    }
                }

                return result;
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

        private MessageCollectionResponse GetMessagesToProcess(
            string folderId)
        {
            int topEmails;

            if (!int.TryParse(number_of_mails, out topEmails) ||
                topEmails <= 0)
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
                        new[] { "receivedDateTime asc" };
                    config.QueryParameters.Select = new[]
                    {
                        "id",
                        "subject",
                        "receivedDateTime",
                        "hasAttachments",
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

        private Message GetCompleteMessage(string messageId)
        {
            return graphService.Users[sharedmailbox_name]
                .Messages[messageId]
                .GetAsync(config =>
                {
                    config.QueryParameters.Select = new[]
                    {
                        "id",
                        "subject",
                        "receivedDateTime",
                        "hasAttachments"
                    };
                })
                .GetAwaiter()
                .GetResult();
        }

        private string DownloadMatchingAttachment(
            string messageId,
            string expectedPrefix)
        {
            AttachmentCollectionResponse response =
                graphService.Users[sharedmailbox_name]
                    .Messages[messageId]
                    .Attachments
                    .GetAsync()
                    .GetAwaiter()
                    .GetResult();

            foreach (Attachment attachment in
                response?.Value ?? new List<Attachment>())
            {
                FileAttachment fileAttachment =
                    attachment as FileAttachment;

                if (fileAttachment == null ||
                    !(fileAttachment.Name ?? "").StartsWith(
                        expectedPrefix,
                        StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (fileAttachment.ContentBytes == null &&
                    !string.IsNullOrWhiteSpace(fileAttachment.Id))
                {
                    fileAttachment =
                        graphService.Users[sharedmailbox_name]
                            .Messages[messageId]
                            .Attachments[fileAttachment.Id]
                            .GetAsync()
                            .GetAwaiter()
                            .GetResult() as FileAttachment;
                }

                if (fileAttachment?.ContentBytes == null)
                {
                    throw new InvalidOperationException(
                        "Attachment content is empty : " +
                        attachment.Name);
                }

                string path = Path.Combine(
                    temp_folder,
                    global_application_name + "_" +
                    CleanFileName(fileAttachment.Name));

                File.WriteAllBytes(
                    path,
                    fileAttachment.ContentBytes);

                return path;
            }

            throw new InvalidOperationException(
                "No attachment starting with '" +
                expectedPrefix + "' was found");
        }

        private static DataTable ReadExcelSheet(
            string filePath,
            string worksheetName)
        {
            if (string.IsNullOrWhiteSpace(filePath) ||
                !File.Exists(filePath))
            {
                throw new FileNotFoundException(
                    "Excel file not found",
                    filePath);
            }

            if (string.IsNullOrWhiteSpace(worksheetName))
            {
                throw new ArgumentException(
                    "Worksheet name is empty.",
                    nameof(worksheetName));
            }

            using (FileStream stream = new FileStream(
                filePath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.Read))
            using (IExcelDataReader reader =
                ExcelReaderFactory.CreateReader(stream))
            {
                DataSet dataSet = reader.AsDataSet(
                    new ExcelDataSetConfiguration
                    {
                        ConfigureDataTable = tableReader =>
                            new ExcelDataTableConfiguration
                            {
                                UseHeaderRow = true
                            }
                    });

                if (dataSet == null ||
                    dataSet.Tables == null ||
                    dataSet.Tables.Count == 0)
                {
                    throw new InvalidOperationException(
                        "No worksheet found in Excel file : " +
                        Path.GetFileName(filePath));
                }

                DataTable worksheet = dataSet.Tables
                    .Cast<DataTable>()
                    .FirstOrDefault(table =>
                        string.Equals(
                            table.TableName?.Trim(),
                            worksheetName.Trim(),
                            StringComparison.OrdinalIgnoreCase));

                if (worksheet == null)
                {
                    string availableWorksheets = string.Join(
                        ", ",
                        dataSet.Tables
                            .Cast<DataTable>()
                            .Select(table => table.TableName));

                    throw new InvalidOperationException(
                        "Worksheet not found : " +
                        worksheetName +
                        ". Available worksheets : " +
                        availableWorksheets);
                }

                if (worksheet.Rows.Count == 0)
                {
                    throw new InvalidOperationException(
                        "No data found in worksheet : " +
                        worksheetName);
                }

                return worksheet.Copy();
            }
        }

        private void BulkReplaceTable(
            DataTable data,
            string destinationTable)
        {
            ValidateSqlTableName(destinationTable);

            using (SqlConnection connection =
                new SqlConnection(sql_connexion))
            {
                connection.Open();

                using (SqlTransaction transaction =
                    connection.BeginTransaction())
                {
                    try
                    {
                        using (SqlCommand truncate = new SqlCommand(
                            "TRUNCATE TABLE " +
                            QuoteIdentifier(destinationTable),
                            connection,
                            transaction))
                        {
                            truncate.CommandTimeout = 300;
                            truncate.ExecuteNonQuery();
                        }

                        using (SqlBulkCopy bulkCopy = new SqlBulkCopy(
                            connection,
                            SqlBulkCopyOptions.CheckConstraints,
                            transaction))
                        {
                            bulkCopy.BulkCopyTimeout = 300;
                            bulkCopy.DestinationTableName =
                                QuoteIdentifier(destinationTable);
                            bulkCopy.WriteToServer(data);
                        }

                        transaction.Commit();
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw;
                    }
                }
            }
        }

        private void ExecuteSql(
            string sql,
            string methodName)
        {
            try
            {
                using (SqlConnection connection =
                    new SqlConnection(sql_connexion))
                using (SqlCommand command =
                    new SqlCommand(sql, connection))
                {
                    command.CommandTimeout = 300;
                    connection.Open();
                    command.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                SendTechnicalIssueMail(
                    methodName,
                    sql,
                    ex.Message,
                    "SQL");

                throw new TechnicalAlertAlreadySentException(
                    "SQL error in " + methodName + " : " +
                    ex.Message,
                    ex);
            }
        }

        private string BuildMessageFilter()
        {
            if (string.IsNullOrWhiteSpace(start_date_scan))
            {
                return "";
            }

            DateTime searchDate;

            if (DateTime.TryParseExact(
                    start_date_scan,
                    "yyyyMMdd",
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out searchDate) ||
                DateTime.TryParse(
                    start_date_scan,
                    out searchDate))
            {
                return "receivedDateTime ge " +
                    searchDate.ToUniversalTime().ToString(
                        "yyyy-MM-ddTHH:mm:ssZ",
                        CultureInfo.InvariantCulture);
            }

            throw new InvalidOperationException(
                "Invalid start_date_scan value : " +
                start_date_scan);
        }

        private MailFolder GetInputFolder()
        {
            if (string.Equals(
                sharedmailbox_folder_in,
                "Inbox",
                StringComparison.OrdinalIgnoreCase))
            {
                return graphService.Users[sharedmailbox_name]
                    .MailFolders["inbox"]
                    .GetAsync()
                    .GetAwaiter()
                    .GetResult();
            }

            return GetChildFolderByName(
                sharedmailbox_folder_in);
        }

        private MailFolder GetChildFolderByName(string folderName)
        {
            string safeFolderName =
                (folderName ?? "").Replace("'", "''");

            MailFolderCollectionResponse folders =
                graphService.Users[sharedmailbox_name]
                    .MailFolders["inbox"]
                    .ChildFolders
                    .GetAsync(config =>
                    {
                        config.QueryParameters.Filter =
                            "displayName eq '" +
                            safeFolderName + "'";
                    })
                    .GetAwaiter()
                    .GetResult();

            MailFolder folder =
                folders?.Value?.FirstOrDefault();

            if (folder == null)
            {
                throw new DirectoryNotFoundException(
                    "Folder not found : " + folderName);
            }

            return folder;
        }

        private void MarkEmailAsRead(string messageId)
        {
            graphService.Users[sharedmailbox_name]
                .Messages[messageId]
                .PatchAsync(new Message { IsRead = true })
                .GetAwaiter()
                .GetResult();
        }

        private void MoveEmail(
            string messageId,
            string destinationFolderId)
        {
            var requestBody =
                new Microsoft.Graph.Users.Item.Messages.Item.Move
                    .MovePostRequestBody
                {
                    DestinationId = destinationFolderId
                };

            graphService.Users[sharedmailbox_name]
                .Messages[messageId]
                .Move
                .PostAsync(requestBody)
                .GetAwaiter()
                .GetResult();
        }

        private string ExtractExcelFile(
            string sourceFile,
            string destinationDirectory)
        {
            if (string.IsNullOrWhiteSpace(sourceFile) ||
                !File.Exists(sourceFile))
            {
                throw new FileNotFoundException(
                    "Attachment file not found",
                    sourceFile);
            }

            string extension =
                Path.GetExtension(sourceFile).ToLowerInvariant();

            if (extension == ".xls" || extension == ".xlsx")
            {
                return sourceFile;
            }

            if (extension != ".zip")
            {
                throw new NotSupportedException(
                    "Unsupported attachment format : " + extension +
                    ". Expected format : .xls, .xlsx or .zip");
            }

            string extractionDirectory = Path.Combine(
                destinationDirectory,
                global_application_name + "_EXTRACT_" +
                Guid.NewGuid().ToString("N"));

            Directory.CreateDirectory(extractionDirectory);

            try
            {
                using (FileStream archiveStream = new FileStream(
                    sourceFile,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.Read))
                using (ZipArchive archive = new ZipArchive(
                    archiveStream,
                    ZipArchiveMode.Read))
                {
                    foreach (ZipArchiveEntry entry in archive.Entries)
                    {
                        if (string.IsNullOrWhiteSpace(entry.Name))
                        {
                            continue;
                        }

                        string destinationPath = Path.GetFullPath(
                            Path.Combine(
                                extractionDirectory,
                                entry.FullName));

                        string extractionRoot = Path.GetFullPath(
                            extractionDirectory +
                            Path.DirectorySeparatorChar);

                        if (!destinationPath.StartsWith(
                            extractionRoot,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            throw new InvalidOperationException(
                                "Invalid path detected in ZIP archive : " +
                                entry.FullName);
                        }

                        string entryDirectory =
                            Path.GetDirectoryName(destinationPath);

                        if (!Directory.Exists(entryDirectory))
                        {
                            Directory.CreateDirectory(entryDirectory);
                        }

                        entry.ExtractToFile(
                            destinationPath,
                            true);
                    }
                }
            }
            catch (InvalidDataException ex)
            {
                throw new InvalidDataException(
                    "The ZIP attachment is invalid or corrupted : " +
                    Path.GetFileName(sourceFile),
                    ex);
            }

            string extractedFile = Directory
                .GetFiles(
                    extractionDirectory,
                    "*.*",
                    SearchOption.AllDirectories)
                .Where(file =>
                {
                    string fileExtension =
                        Path.GetExtension(file);

                    return fileExtension.Equals(
                               ".xls",
                               StringComparison.OrdinalIgnoreCase) ||
                           fileExtension.Equals(
                               ".xlsx",
                               StringComparison.OrdinalIgnoreCase);
                })
                .OrderByDescending(
                    File.GetLastWriteTimeUtc)
                .FirstOrDefault();

            if (string.IsNullOrWhiteSpace(extractedFile))
            {
                throw new FileNotFoundException(
                    "No Excel file was found in ZIP archive : " +
                    Path.GetFileName(sourceFile));
            }

            return extractedFile;
        }

        private void SendTechnicalIssueMail(
            string methodName,
            string sql,
            string message,
            string typeError)
        {
            try
            {
                WriteToFile(
                    "       " + typeError + " error in " +
                    methodName + " : " + message);

                if (graphService == null ||
                    string.IsNullOrWhiteSpace(
                        email_in_case_of_technical_issue) ||
                    string.IsNullOrWhiteSpace(sharedmailbox_name))
                {
                    return;
                }

                Message alert = new Message
                {
                    Subject = global_application_name +
                        " - Erreur " + typeError +
                        " dans " + methodName,
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Html,
                        Content =
                            "<b>Methode :</b> " + methodName +
                            "<br/><b>Mailbox :</b> " +
                            WebUtility.HtmlEncode(sharedmailbox_name) +
                            "<br/><b>Message :</b> " +
                            WebUtility.HtmlEncode(message) +
                            (string.IsNullOrWhiteSpace(sql)
                                ? ""
                                : "<br/><pre>" +
                                  WebUtility.HtmlEncode(sql) +
                                  "</pre>")
                    },
                    ToRecipients = BuildRecipients(
                        email_in_case_of_technical_issue)
                };

                var requestBody =
                    new Microsoft.Graph.Users.Item.SendMail
                        .SendMailPostRequestBody
                    {
                        Message = alert,
                        SaveToSentItems = true
                    };

                graphService.Users[sharedmailbox_name]
                    .SendMail
                    .PostAsync(requestBody)
                    .GetAwaiter()
                    .GetResult();
            }
            catch (Exception ex)
            {
                WriteToFile(
                    "Error sending technical issue email : " +
                    ex.Message);
            }
        }

        private static List<Recipient> BuildRecipients(
            string addresses)
        {
            return (addresses ?? "")
                .Split(
                    new[] { ';', ',' },
                    StringSplitOptions.RemoveEmptyEntries)
                .Select(value => value.Trim())
                .Where(value => value.Contains("@"))
                .Select(value => new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = value
                    }
                })
                .ToList();
        }

        private void ValidateRequiredCountryParameters()
        {
            if (string.IsNullOrWhiteSpace(sql_connexion))
            {
                throw new InvalidOperationException(
                    "sql_connexion is empty. Parameter used : " +
                    sql_connexion_parameter_global);
            }

            if (string.IsNullOrWhiteSpace(temp_folder))
            {
                throw new InvalidOperationException(
                    "temp_folder is empty");
            }
        }

        private void ValidateRequiredMailboxParameters()
        {
            if (id_mailboxe <= 0)
            {
                throw new InvalidOperationException(
                    "id_mailboxe is invalid");
            }

            if (string.IsNullOrWhiteSpace(sharedmailbox_name))
            {
                throw new InvalidOperationException(
                    "mailboxe is empty");
            }
        }

        private string get_IMCA_paramters(
            string connectionString,
            string parameterName)
        {
            const string sql = @"
SELECT ISNULL(VALUE, '')
FROM PCM_TAB_IMCA_PARAMETER_GLOBAL
WHERE SK_VALID = 0
  AND PARAMETER = @PARAMETER;";

            using (SqlConnection connection =
                new SqlConnection(connectionString))
            using (SqlCommand command =
                new SqlCommand(sql, connection))
            {
                command.CommandTimeout = 300;
                command.Parameters.Add(
                    "@PARAMETER",
                    SqlDbType.NVarChar,
                    255).Value = parameterName;
                connection.Open();
                return Convert.ToString(
                    command.ExecuteScalar());
            }
        }

        private void WriteToFile(string message)
        {
            if (string.IsNullOrWhiteSpace(logs_folder))
            {
                logs_folder =
                    AppDomain.CurrentDomain.BaseDirectory;
            }

            Directory.CreateDirectory(logs_folder);

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

        private static string GetServicePath()
        {
            string location =
                System.Reflection.Assembly
                    .GetEntryAssembly()?.Location;

            return string.IsNullOrWhiteSpace(location)
                ? AppDomain.CurrentDomain.BaseDirectory
                : Path.GetDirectoryName(location);
        }

        private static string ExtractDateFromSubject(
            string subject,
            DateTime fallback)
        {
            string source = subject ?? "";
            string ending = source.Length >= 10
                ? source.Substring(source.Length - 10)
                    .Replace('.', '/')
                : "";

            DateTime parsedDate;

            return DateTime.TryParse(
                ending,
                CultureInfo.GetCultureInfo("fr-FR"),
                DateTimeStyles.None,
                out parsedDate)
                ? parsedDate.ToString("dd/MM/yyyy")
                : fallback.ToString("dd/MM/yyyy");
        }

        private static string CleanFileName(string value)
        {
            return string.Join(
                "_",
                (value ?? "attachment")
                    .Split(Path.GetInvalidFileNameChars()))
                .Trim();
        }

        private static void DeleteFiles(
            string folder,
            string prefix)
        {
            if (!Directory.Exists(folder))
            {
                return;
            }

            foreach (string file in Directory.GetFiles(folder))
            {
                try
                {
                    if (Path.GetFileName(file).StartsWith(
                        prefix,
                        StringComparison.OrdinalIgnoreCase))
                    {
                        File.Delete(file);
                    }
                }
                catch
                {
                }
            }

            foreach (string directory in
                Directory.GetDirectories(folder))
            {
                try
                {
                    if (Path.GetFileName(directory).StartsWith(
                        prefix,
                        StringComparison.OrdinalIgnoreCase))
                    {
                        Directory.Delete(
                            directory,
                            true);
                    }
                }
                catch
                {
                }
            }
        }

        private static bool IsTrue(string value)
        {
            return string.Equals(
                value?.Trim(),
                "TRUE",
                StringComparison.OrdinalIgnoreCase);
        }

        private static void ValidateSqlTableName(string value)
        {
            if (string.IsNullOrWhiteSpace(value) ||
                !Regex.IsMatch(
                    value,
                    @"^[A-Za-z_][A-Za-z0-9_]*(\.[A-Za-z_][A-Za-z0-9_]*)?$"))
            {
                throw new InvalidOperationException(
                    "Invalid SQL table name : " + value);
            }
        }

        private static string QuoteIdentifier(string value)
        {
            return string.Join(
                ".",
                value.Split('.')
                    .Select(part =>
                        "[" + part.Replace("]", "]]") + "]"));
        }

        // ----- Set parameters, same readable model as TRACT_SYNDICAL_FR -----
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
            number_of_mails = string.IsNullOrWhiteSpace(value)
                ? "10"
                : value;
        }

        public void setIdMailboxe(int value)
        {
            id_mailboxe = value;
        }

        public void setNomMailboxe(string value)
        {
            nom_mailboxe = value ?? "";
        }

        public void setSharedMailboxName(string value)
        {
            sharedmailbox_name = value ?? "";
        }

        public void setsharedmailbox_folder_in(string value)
        {
            sharedmailbox_folder_in =
                string.IsNullOrWhiteSpace(value)
                    ? "Inbox"
                    : value;
        }

        public void setsharedmailbox_folder_out(string value)
        {
            sharedmailbox_folder_out =
                string.IsNullOrWhiteSpace(value)
                    ? "Archives"
                    : value;
        }

        public void setEmailInCaseOfTechnicalIssueParam(
            string value)
        {
            email_in_case_of_technical_issue_parameter_global =
                value ?? "";
        }

        public void setEmailInCaseOfTechnicalIssue(string value)
        {
            email_in_case_of_technical_issue = value ?? "";
        }

        public void setSqlConnexionParam(string value)
        {
            sql_connexion_parameter_global = value ?? "";
        }

        public void setSqlConnexion(string value)
        {
            sql_connexion = value ?? "";
        }

        public void setlogs_folder(string value)
        {
            logs_folder = value ?? "";
        }

        public void settemp_folder(string value)
        {
            temp_folder = value ?? "";
        }

        public void setGlobalSessionName(string value)
        {
            global_session_name = value ?? "";
        }

        public void setGraphService(GraphServiceClient value)
        {
            graphService = value;
        }
    }
}
