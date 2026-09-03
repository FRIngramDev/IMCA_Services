using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.VisualBasic.FileIO;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace IMEI_MANAGEMENT_FR
{
    public class IMEI_MANAGEMENT_FR
    {
        public IMEI_MANAGEMENT_FR()
        {
        }

        // ----- Current country / processing parameters -----
        private string country = "";                       // Country code being processed
        private string sk_valid = "";                      // SK_VALID parameter
        private string name = "";                          // Readable country / instance name
        private string active = "";                        // Indicates whether processing is active
        private string debug = "";                         // Enables detailed logs when TRUE
        private string start_date_scan = "";               // Start date used to search emails
        private string number_of_mails = "10";             // Maximum number of emails to retrieve

        // ----- Current shared mailbox parameters -----
        private string sharedmailbox_name = "";            // Current shared mailbox address
        private string sharedmailbox_folder_in = "";       // Source folder
        private string sharedmailbox_folder_out = "";      // Archive folder
        private string table_name = "";                    // Business destination table
        private string type_datas = "";                    // Expected data type: IMEI or SN
        private bool has_header = true;                    // TRUE when the CSV contains a header row
        private List<string> column_mapping = new List<string>(); // Required only for files without headers
        private List<string> mandatory_columns = new List<string>(); // Columns that must exist in the CSV before import

        // ----- Technical alert / recipient parameters -----
        private string email_in_case_of_technical_issue_parameter_global = "";
        private string email_in_case_of_technical_issue = "";
        private string email_destinataire = "";

        // ----- Files / logs / session -----
        private string logs_folder = "";
        private string temp_folder = "";
        private string global_session_name = "";

        // ----- Application name -----
        private string global_application_name = "IMEI_MANAGEMENT_FR";

        // ----- SQL connection parameters -----
        private string sql_connexion = "";
        private string sql_connexion_parameter_global = "";

        // ----- Microsoft Graph client -----
        private GraphServiceClient graphService = null;

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

            public List<SharedMailbox> sharedmailboxes { get; set; } = new List<SharedMailbox>();
            public string sharedmailbox_folder_in { get; set; } = "INBOX";
            public string sharedmailbox_folder_out { get; set; } = "Archives";

            public string email_in_case_of_technical_issue_parameter_global { get; set; } = "";
            public string email_destinataire { get; set; } = "";
            public string sql_connexion_parameter_global { get; set; } = "";
        }

        public class SharedMailbox
        {
            public string sharedmailbox_name { get; set; } = "";
            public string table_name { get; set; } = "";
            public string type_datas { get; set; } = "";
            public bool has_header { get; set; } = true;
            public List<string> column_mapping { get; set; } = new List<string>();
            public List<string> mandatory_columns { get; set; } = new List<string>();
            public string active { get; set; } = "TRUE";
        }

        /// <summary>
        /// Creates and returns a Microsoft Graph client.
        /// </summary>
        private GraphServiceClient Connexion_Microsoft_Graph()
        {
            System.Net.ServicePointManager.SecurityProtocol =
                System.Net.SecurityProtocolType.Tls12;

            GraphServiceClient GraphService;

            class_dev_tools.Ews_Modern_Auth Ews_Modern_Auth = new class_dev_tools.Ews_Modern_Auth();

            GraphService = Ews_Modern_Auth.Get_Graph_Service();

            return GraphService;
        }

        /// <summary>
        /// Main entry point used by IMCA to read emails from every configured shared mailbox.
        /// The business processing must be added in ProcessEmailBusinessData.
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
                    setEmailInCaseOfTechnicalIssueParam(p.email_in_case_of_technical_issue_parameter_global);

                    email_in_case_of_technical_issue = get_IMCA_paramters(sql_con, email_in_case_of_technical_issue_parameter_global);

                    sql_connexion = get_IMCA_paramters(sql_con, sql_connexion_parameter_global);

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

                        // Process each mailbox independently so one mailbox failure
                        // does not prevent the other mailboxes from being scanned.
                        foreach (SharedMailbox mailbox in p.sharedmailboxes)
                        {
                            setSharedMailboxName(mailbox.sharedmailbox_name);
                            setTableName(mailbox.table_name);
                            setTypeDatas(mailbox.type_datas);
                            setHasHeader(mailbox.has_header);
                            setColumnMapping(mailbox.column_mapping);
                            setMandatoryColumns(mailbox.mandatory_columns);
                            setSharemailboxActive(mailbox.active);

                            if (mailbox.active.ToUpper().Trim() != "TRUE")
                            {
                                continue;
                            }

                            try
                            {
                                ReadCurrentSharedMailbox();
                            }
                            catch (Exception mailboxException)
                            {
                                WriteToFile(
                                    "   Error reading mailbox " +
                                    sharedmailbox_name + " : " +
                                    mailboxException.Message);
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        WriteToFile("   Error get emails : " + e.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteToFile(
                    "Global error Read_Email_with_Graph : " + ex.Message);
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        /// <summary>
        /// Reads the emails for the currently selected shared mailbox.
        /// </summary>
        private void ReadCurrentSharedMailbox()
        {
            try
            {
                int nb_mail = 0;
                string etat_processing = "";

                ValidateRequiredMailboxParameters();

                if (debug.ToUpper() == "TRUE")
                {
                    WriteToFile(
                        "   Connexion to " + sharedmailbox_name +
                        " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                    WriteToFile("   Sharedmailbox_folder_in  : " + sharedmailbox_folder_in);
                    WriteToFile("   Sharedmailbox_folder_out : " + sharedmailbox_folder_out);
                    WriteToFile("   Table name               : " + table_name);
                    WriteToFile("   Data type                : " + type_datas);
                    WriteToFile("   Temp folder              : " + temp_folder);
                    WriteToFile("   Extracting the " + number_of_mails + " oldest messages");
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

                    foreach (Message email in messages.Value)
                    {
                        try
                        {
                            etat_processing = "OK";

                            if (debug.ToUpper() == "TRUE")
                            {
                                WriteToFile(
                                    "       Subject : " + email.Subject +
                                    " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                            }

                            if (email.HasAttachments == true)
                            {
                                List<string> attachments = DownloadCSVAttachmentsWithGraph(
                                    graphService,
                                    sharedmailbox_name,
                                    email.Id
                                );

                                if (attachments.Count == 0)
                                {
                                    string noCsvAttachmentMessage =
                                        sharedmailbox_name +
                                        " - Mail : " + (email.Subject ?? "<no subject>") +
                                        " - Mail not containing any CSV files.";

                                    SendTechnicalIssueMail(
                                        nameof(ReadCurrentSharedMailbox),
                                        "",
                                        noCsvAttachmentMessage,
                                        "NO CSV ATTACHMENT");

                                    throw new Exception(noCsvAttachmentMessage);
                                }

                                foreach (string attachment in attachments)
                                {

                                    // This is the only section intended to contain mailbox-specific
                                    // and email-specific business processing.
                                    string attachmentResult = ProcessEmailBusinessData(attachment, email.Subject);
                                    if (attachmentResult != "OK")
                                    {
                                        etat_processing = "KO";
                                    }
                                }
                            }
                            else
                            {
                                string noAttachmentMessage =
                                    sharedmailbox_name +
                                    " - Mail : " + (email.Subject ?? "<no subject>") +
                                    " - Mail not containing any attachments.";

                                // Notify technical support before the message is handled
                                // by the email-level catch block and moved to Erreur.
                                SendTechnicalIssueMail(
                                    nameof(ReadCurrentSharedMailbox),
                                    "",
                                    noAttachmentMessage,
                                    "NO ATTACHMENT");

                                throw new Exception(noAttachmentMessage);
                            }

                            MarkEmailAsRead(graphService, sharedmailbox_name, email.Id);
                            switch (etat_processing)
                            {
                                case "OK":
                                    MoveEmail(
                                        graphService,
                                        sharedmailbox_name,
                                        email.Id,
                                        archiveFolder.Id);

                                    nb_mail++;
                                    break;

                                case "KO":
                                    MoveEmail(
                                        graphService,
                                        sharedmailbox_name,
                                        email.Id,
                                        errorFolder.Id);
                                    break;

                                default:
                                    throw new InvalidOperationException(
                                        "Unknown processing status: " + etat_processing);
                            }

                            System.Threading.Thread.Sleep(500);
                        }
                        catch (Exception ex)
                        {
                            WriteToFile(
                                "   Error processing email : " + ex.Message +
                                " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));

                            try
                            {
                                MarkEmailAsRead(
                                    graphService,
                                    sharedmailbox_name,
                                    email.Id);

                                MoveEmail(
                                    graphService,
                                    sharedmailbox_name,
                                    email.Id,
                                    errorFolder.Id);
                            }
                            catch (Exception moveException)
                            {
                                WriteToFile(
                                    "   Error moving email to Erreur folder : " +
                                    moveException.Message);
                            }
                        }

                    }

                    if (debug.ToUpper() == "TRUE")
                    {
                        WriteToFile(
                            "   " + nb_mail +
                            " Email(s) have been processed at " +
                            DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                    }
                }
                else if (debug.ToUpper() == "TRUE")
                {
                    WriteToFile("   No emails found");
                }
            }
            catch (Exception ex)
            {
                WriteToFile("Global error ReadCurrentSharedMailbox : " + ex.Message);
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        /// <summary>
        /// Business processing placeholder.
        /// Add the logic that depends on the current mailbox and message here.
        /// </summary>
        private string ProcessEmailBusinessData(string filePath, string subject)
        {
            const string metadataSql = @"
                                        SELECT c.column_id, c.name AS column_name, c.is_nullable,
                                               CASE WHEN dc.object_id IS NULL THEN 0 ELSE 1 END AS has_default
                                        FROM sys.columns c
                                        INNER JOIN sys.tables t ON t.object_id = c.object_id
                                        INNER JOIN sys.schemas s ON s.schema_id = t.schema_id
                                        LEFT JOIN sys.default_constraints dc
                                               ON dc.parent_object_id = c.object_id
                                              AND dc.parent_column_id = c.column_id
                                        WHERE UPPER(t.name) = UPPER(@TABLE_NAME)
                                          AND UPPER(s.name) = UPPER(@SCHEMA_NAME)
                                          AND c.is_identity = 0
                                          AND c.is_computed = 0
                                          AND c.system_type_id <> 189
                                        ORDER BY c.column_id;";

            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                throw new FileNotFoundException(
                    "IMEI file not found for " + sharedmailbox_name,
                    filePath);
            }

            try
            {
                string schemaName;
                string destinationTableName;
                ParseAndValidateTableName(table_name, out schemaName, out destinationTableName);
                string qualifiedTableName = QuoteSqlIdentifier(schemaName) + "." +
                                            QuoteSqlIdentifier(destinationTableName);

                DataTable metadata = LoadDestinationMetadata(
                    metadataSql,
                    schemaName,
                    destinationTableName);

                if (metadata.Rows.Count == 0)
                {
                    throw new InvalidOperationException(
                        "Table not found or without importable columns : " + qualifiedTableName);
                }

                DataTable importTable = new DataTable();
                List<string> mappedColumns = new List<string>();
                bool extraColumnAlertSent = false;

                using (TextFieldParser reader = new TextFieldParser(filePath))
                {
                    reader.TextFieldType = FieldType.Delimited;
                    reader.SetDelimiters(";");
                    reader.HasFieldsEnclosedInQuotes = true;
                    reader.TrimWhiteSpace = false;

                    Dictionary<string, int> sourcePositions;
                    int expectedSourceColumnCount;

                    if (has_header)
                    {
                        string[] headers;
                        sourcePositions = ReadHeaderMapping(reader, out headers);
                        expectedSourceColumnCount = headers.Length;

                        List<string> destinationNames = metadata.AsEnumerable()
                            .Select(row => NormalizeColumnName(Convert.ToString(row["column_name"])))
                            .ToList();

                        List<string> unexpectedHeaders = headers
                            .Where(header => !string.IsNullOrWhiteSpace(header))
                            .Where(header => !destinationNames.Contains(
                                NormalizeColumnName(CleanCsvValue(header)),
                                StringComparer.OrdinalIgnoreCase))
                            .ToList();

                        if (unexpectedHeaders.Count > 0)
                        {
                            SendNewColumnInformationMail(
                                filePath,
                                unexpectedHeaders,
                                qualifiedTableName);
                            extraColumnAlertSent = true;
                        }
                    }
                    else
                    {
                        sourcePositions = BuildPositionalMapping(column_mapping);
                        expectedSourceColumnCount = column_mapping.Count;
                    }

                    // JSON mandatory columns are checked before any row is loaded.
                    // A missing mandatory column stops the import and returns KO.
                    ValidateMandatoryColumns(
                        sourcePositions,
                        metadata,
                        qualifiedTableName,
                        filePath);

                    BuildImportColumns(
                        metadata,
                        sourcePositions,
                        importTable,
                        mappedColumns);

                    while (!reader.EndOfData)
                    {
                        string[] fields;
                        try
                        {
                            fields = reader.ReadFields();
                        }
                        catch (MalformedLineException ex)
                        {
                            throw new FormatException(
                                "Invalid CSV line : " + ex.Message,
                                ex);
                        }

                        if (fields == null || fields.All(string.IsNullOrWhiteSpace))
                        {
                            continue;
                        }

                        // For a headerless file, additional positions cannot be named.
                        // Inform support once, but continue importing all known positions.
                        if (!has_header &&
                            fields.Length > expectedSourceColumnCount &&
                            !extraColumnAlertSent)
                        {
                            List<string> additionalPositions = new List<string>();
                            for (int index = expectedSourceColumnCount;
                                 index < fields.Length;
                                 index++)
                            {
                                additionalPositions.Add("Position " + (index + 1));
                            }

                            SendNewColumnInformationMail(
                                filePath,
                                additionalPositions,
                                qualifiedTableName);
                            extraColumnAlertSent = true;
                        }

                        DataRow newRow = importTable.NewRow();
                        foreach (string destinationColumn in mappedColumns)
                        {
                            int sourceIndex = sourcePositions[
                                NormalizeColumnName(destinationColumn)];

                            if (sourceIndex >= fields.Length)
                            {
                                throw new InvalidOperationException(
                                    "Insufficient number of fields for column " +
                                    destinationColumn + ".");
                            }

                            string value = CleanCsvValue(fields[sourceIndex]);
                            newRow[destinationColumn] = string.IsNullOrWhiteSpace(value)
                                ? (object)DBNull.Value
                                : value;
                        }

                        if (IsBusinessRowValid(newRow))
                        {
                            importTable.Rows.Add(newRow);
                        }
                    }
                }

                if (importTable.Rows.Count == 0)
                {
                    throw new InvalidOperationException(
                        "No valid row found in file " + filePath);
                }

                BulkInsert(importTable, mappedColumns, qualifiedTableName);

                WriteToFile(
                    "       " + importTable.Rows.Count + " row(s) imported into " +
                    qualifiedTableName + " from " + Path.GetFileName(filePath));

                return "OK";
            }
            catch (Exception ex)
            {
                // ValidateMandatoryColumns already sends a dedicated alert.
                // Do not send a second generic IMPORT email for the same issue.
                if (!(ex is MandatoryColumnMissingException))
                {
                    SendTechnicalIssueMail(
                        nameof(ProcessEmailBusinessData),
                        metadataSql,
                        sharedmailbox_name + " - Table : " + table_name +
                        " - File : " + filePath + " - " + ex.Message,
                        ex is SqlException ? "SQL" : "IMPORT");
                }

                WriteToFile(
                    "       Import error for mail '" + subject);
                return "KO";
            }
        }

        private void SendNewColumnInformationMail(
            string filePath,
            List<string> newColumns,
            string destinationTable)
        {
            string message =
                sharedmailbox_name + " - Table : " + destinationTable +
                " - File : " + Path.GetFileName(filePath) +
                " - New column(s) detected and ignored during import : " +
                string.Join(", ", newColumns) +
                ". The import of known columns will continue. " +
                "Please check whether the destination table and mapping must be updated.";

            SendTechnicalIssueMail(
                nameof(ProcessEmailBusinessData),
                "",
                message,
                "NEW COLUMN");
        }

        private DataTable LoadDestinationMetadata(
            string sql,
            string schemaName,
            string destinationTableName)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(sql_connexion))
                using (SqlCommand command = new SqlCommand(sql, connection))
                using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                {
                    command.CommandTimeout = 300;
                    command.Parameters.Add("@TABLE_NAME", SqlDbType.NVarChar, 128).Value = destinationTableName;
                    command.Parameters.Add("@SCHEMA_NAME", SqlDbType.NVarChar, 128).Value = schemaName;
                    DataTable table = new DataTable();
                    connection.Open();
                    adapter.Fill(table);
                    return table;
                }
            }
            catch (Exception ex)
            {
                WriteToFile("       SQL error in LoadDestinationMetadata : " + ex);
                throw;
            }
        }

        private Dictionary<string, int> ReadHeaderMapping(
            TextFieldParser reader,
            out string[] headers)
        {
            if (reader.EndOfData)
            {
                throw new InvalidOperationException("The CSV file is empty.");
            }

            headers = reader.ReadFields();
            if (headers == null || headers.Length == 0)
            {
                throw new InvalidOperationException("The CSV header is empty.");
            }

            Dictionary<string, int> positions =
                new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            for (int index = 0; index < headers.Length; index++)
            {
                string normalized = NormalizeColumnName(CleanCsvValue(headers[index]));
                if (string.IsNullOrWhiteSpace(normalized))
                {
                    throw new InvalidOperationException(
                        "A column in the CSV does not have a name.");
                }

                if (positions.ContainsKey(normalized))
                {
                    throw new InvalidOperationException(
                        "Duplicate column in CSV : " + headers[index]);
                }

                positions.Add(normalized, index);
            }

            return positions;
        }

        private Dictionary<string, int> BuildPositionalMapping(List<string> configuredMapping)
        {
            if (configuredMapping == null || configuredMapping.Count == 0)
            {
                throw new InvalidOperationException(
                    "column_mapping is required when has_header=false.");
            }

            Dictionary<string, int> positions =
                new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            for (int sourceIndex = 0; sourceIndex < configuredMapping.Count; sourceIndex++)
            {
                string configuredColumn = configuredMapping[sourceIndex];
                if (string.IsNullOrWhiteSpace(configuredColumn))
                {
                    continue;
                }

                string normalized = NormalizeColumnName(configuredColumn);
                if (positions.ContainsKey(normalized))
                {
                    throw new InvalidOperationException(
                        "Duplicate column in mapping : " + configuredColumn);
                }

                positions.Add(normalized, sourceIndex);
            }

            return positions;
        }

        private void ValidateMandatoryColumns(
            Dictionary<string, int> sourcePositions,
            DataTable metadata,
            string qualifiedTableName,
            string filePath)
        {
            if (mandatory_columns == null || mandatory_columns.Count == 0)
            {
                return;
            }

            HashSet<string> destinationColumns = new HashSet<string>(
                metadata.AsEnumerable()
                    .Select(row => NormalizeColumnName(
                        Convert.ToString(row["column_name"]))),
                StringComparer.OrdinalIgnoreCase);

            List<string> invalidConfiguration = mandatory_columns
                .Where(column => !string.IsNullOrWhiteSpace(column))
                .Where(column => !destinationColumns.Contains(
                    NormalizeColumnName(column)))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (invalidConfiguration.Count > 0)
            {
                throw new InvalidOperationException(
                    "Mandatory column(s) configured in JSON do not exist in " +
                    qualifiedTableName + " : " +
                    string.Join(", ", invalidConfiguration));
            }

            List<string> missingMandatoryColumns = mandatory_columns
                .Where(column => !string.IsNullOrWhiteSpace(column))
                .Where(column => !sourcePositions.ContainsKey(
                    NormalizeColumnName(column)))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (missingMandatoryColumns.Count == 0)
            {
                return;
            }

            string message =
                sharedmailbox_name +
                " - Table : " + qualifiedTableName +
                " - File : " + Path.GetFileName(filePath) +
                " - Mandatory column(s) missing from the CSV : " +
                string.Join(", ", missingMandatoryColumns) +
                ". No data has been imported and the email will be moved to the Erreur folder.";

            SendTechnicalIssueMail(
                nameof(ProcessEmailBusinessData),
                "",
                message,
                "MANDATORY COLUMN MISSING");

            // The exception is caught by ProcessEmailBusinessData, which returns KO.
            // ReadCurrentSharedMailbox then moves the source email to Erreur.
            throw new MandatoryColumnMissingException(message);
        }

        private sealed class MandatoryColumnMissingException : Exception
        {
            public MandatoryColumnMissingException(string message)
                : base(message)
            {
            }
        }

        private static void BuildImportColumns(
            DataTable metadata,
            Dictionary<string, int> sourcePositions,
            DataTable importTable,
            List<string> mappedColumns)
        {
            foreach (DataRow column in metadata.Rows)
            {
                string destinationColumn = Convert.ToString(column["column_name"]);
                string normalized = NormalizeColumnName(destinationColumn);
                bool required = !Convert.ToBoolean(column["is_nullable"]) &&
                                !Convert.ToBoolean(column["has_default"]);

                if (!sourcePositions.ContainsKey(normalized))
                {
                    if (required)
                    {
                        throw new InvalidOperationException(
                            "Required SQL column missing from mapping : " + destinationColumn);
                    }
                    continue;
                }

                importTable.Columns.Add(destinationColumn, typeof(string));
                mappedColumns.Add(destinationColumn);
            }

            if (mappedColumns.Count == 0)
            {
                throw new InvalidOperationException(
                    "No column in the file corresponds to the SQL table.");
            }
        }

        private void BulkInsert(
            DataTable importTable,
            List<string> mappedColumns,
            string qualifiedTableName)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(sql_connexion))
                {
                    connection.Open();
                    using (SqlTransaction transaction = connection.BeginTransaction())
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(
                        connection,
                        SqlBulkCopyOptions.CheckConstraints,
                        transaction))
                    {
                        bulkCopy.BulkCopyTimeout = 300;
                        bulkCopy.BatchSize = 1000;
                        bulkCopy.DestinationTableName = qualifiedTableName;
                        foreach (string column in mappedColumns)
                        {
                            bulkCopy.ColumnMappings.Add(column, column);
                        }
                        bulkCopy.WriteToServer(importTable);
                        transaction.Commit();
                    }
                }
            }
            catch (Exception ex)
            {
                WriteToFile("       SQL bulk copy error : " + ex);
                throw;
            }
        }

        private static string CleanCsvValue(string value)
        {
            if (value == null)
            {
                return "";
            }

            string cleaned = value.Trim().Trim('\uFEFF');
            if (cleaned.Length >= 2 && cleaned[0] == '"' && cleaned[cleaned.Length - 1] == '"')
            {
                cleaned = cleaned.Substring(1, cleaned.Length - 2);
            }
            return cleaned.Replace("\"\"", "\"").Trim();
        }

        private static void ParseAndValidateTableName(
            string configuredTableName,
            out string schemaName,
            out string destinationTableName)
        {
            string[] parts = (configuredTableName ?? "").Split('.');
            if (parts.Length == 1)
            {
                schemaName = "dbo";
                destinationTableName = parts[0];
            }
            else if (parts.Length == 2)
            {
                schemaName = parts[0];
                destinationTableName = parts[1];
            }
            else
            {
                throw new InvalidOperationException("Invalid table name : " + configuredTableName);
            }

            const string pattern = @"^[A-Za-z_][A-Za-z0-9_]*$";
            if (!Regex.IsMatch(schemaName, pattern) || !Regex.IsMatch(destinationTableName, pattern))
            {
                throw new InvalidOperationException(
                    "Invalid table or schema name : " + configuredTableName);
            }
        }

        private static string QuoteSqlIdentifier(string identifier)
        {
            return "[" + identifier.Replace("]", "]]") + "]";
        }

        private static string NormalizeColumnName(string columnName)
        {
            return (columnName ?? "").Trim().Trim('\uFEFF').Replace(" ", "_").ToUpperInvariant();
        }

        private bool IsBusinessRowValid(DataRow row)
        {
            string dataType = (type_datas ?? "").Trim().ToUpperInvariant();
            if (dataType == "IMEI")
            {
                return HasValue(row, "IMEI") || HasValue(row, "IMEI1");
            }
            if (dataType == "SN")
            {
                return HasValue(row, "serial_number");
            }
            throw new InvalidOperationException("type_datas must be IMEI or SN.");
        }

        private static bool HasValue(DataRow row, string columnName)
        {
            return row.Table.Columns.Contains(columnName) &&
                   row[columnName] != DBNull.Value &&
                   !string.IsNullOrWhiteSpace(Convert.ToString(row[columnName]));
        }

        private void SendTechnicalIssueMail(string methodName, string sql, string message, string type_error)
        {
            try
            {
                WriteToFile("       	" + type_error + " error in " + methodName + " : " + message);

                if (graphService == null)
                {
                    WriteToFile("Unable to send " + type_error + " technical issue email because GraphServiceClient is null");
                    return;
                }

                if (string.IsNullOrWhiteSpace(email_in_case_of_technical_issue) || !email_in_case_of_technical_issue.Contains("@"))
                {
                    WriteToFile("Unable to send " + type_error + " technical issue email because recipient is empty");
                    return;
                }

                string subject = global_application_name + " - Erreur " + type_error + " dans " + methodName;

                string body =
                    "Une erreur " + type_error + " est survenue dans " + global_application_name + ".<br/><br/>" +
                    "<b>Méthode :</b> " + methodName + "<br/>";
                if (type_error == "SQL")
                {
                    body = body +
                    "<b>Timeout configuré :</b> " + 300 + " secondes<br/>";
                }
                body = body +
                    "<b>Message :</b> " + message + "<br/><br/>";
                if (type_error == "SQL")
                {
                    body = body +
                    "<b>Requête SQL :</b><br/>" +
                    "<pre>" + sql + "</pre>";
                }



                EnvoiEmail_with_Graph(
                    graphService,
                    subject,
                    body,
                    email_in_case_of_technical_issue, "", type_error);
            }
            catch (Exception mailEx)
            {
                WriteToFile("Error sending " + type_error + " technical issue email : " + mailEx.Message);
            }
        }

        private void EnvoiEmail_with_Graph(GraphServiceClient graphService, string subject, string body, string recipient, string attachmentPath = "", string type_error = "")
        {
            if (string.IsNullOrWhiteSpace(recipient) || !recipient.Contains("@"))
            {
                return;
            }

            List<Recipient> toRecipients = BuildRecipients(recipient);



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

            if (!string.IsNullOrWhiteSpace(attachmentPath) && File.Exists(attachmentPath))
            {
                message.Attachments = new List<Attachment>
                {
                    new FileAttachment
                    {
                        OdataType = "#microsoft.graph.fileAttachment",
                        Name = Path.GetFileName(attachmentPath),
                        ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        ContentBytes = File.ReadAllBytes(attachmentPath)
                    }
                };
            }
            if (type_error.ToUpper() != "SQL")
            {
                List<Recipient> ccRecipients = BuildRecipients(email_destinataire);
                if (ccRecipients.Count > 0)
                {
                    message.CcRecipients = ccRecipients;
                }
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

        private List<string> DownloadCSVAttachmentsWithGraph(GraphServiceClient graphService, string mailbox, string messageId)
        {
            List<string> downloadedFiles = new List<string>();

            AttachmentCollectionResponse attachments = graphService.Users[mailbox]
                .Messages[messageId]
                .Attachments
                .GetAsync()
                .GetAwaiter()
                .GetResult();

            if (attachments?.Value == null)
            {
                return downloadedFiles;
            }

            foreach (Attachment attachment in attachments.Value)
            {
                if (attachment is FileAttachment fileAttachment)
                {
                    string extension = Path.GetExtension(fileAttachment.Name ?? "").ToLowerInvariant();
                    if (extension != ".csv")
                    {
                        continue;
                    }

                    if (fileAttachment.ContentBytes == null)
                    {
                        throw new Exception("The content of the attachment is empty : " + fileAttachment.Name);
                    }

                    string fileName = CleanFileName(global_application_name + "_" + fileAttachment.Name);
                    string filePath = Path.Combine(temp_folder, fileName);
                    File.WriteAllBytes(filePath, fileAttachment.ContentBytes);
                    downloadedFiles.Add(filePath);
                }
            }

            return downloadedFiles;
        }

        private string CleanFileName(string fileName)
        {
            return string.Join("_", fileName.Split(Path.GetInvalidFileNameChars())).Trim();
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
                "   Invalid start_date_scan value : " + start_date_scan);
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
            string safeFolderName = EscapeODataString(folderName);

            MailFolderCollectionResponse folders = graphService.Users[mailbox]
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
                throw new Exception("Folder not found : " + folderName);
            }

            return folders.Value.First();
        }

        private void ValidateRequiredCountryParameters(Country p)
        {
            if (p.sharedmailboxes == null || p.sharedmailboxes.Count == 0)
            {
                throw new Exception("sharedmailboxes is empty");
            }

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

            if (string.IsNullOrWhiteSpace(email_in_case_of_technical_issue))
            {
                throw new Exception(
                    "email_in_case_of_technical_issue is empty. Parameter used : " +
                    email_in_case_of_technical_issue_parameter_global);
            }
        }

        private void ValidateRequiredMailboxParameters()
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

            if (string.IsNullOrWhiteSpace(table_name))
            {
                throw new Exception(
                    "table_name is empty for " + sharedmailbox_name);
            }

            if (type_datas.ToUpper().Trim() != "IMEI" &&
                type_datas.ToUpper().Trim() != "SN")
            {
                throw new Exception(
                    "type_datas must be IMEI or SN for " +
                    sharedmailbox_name);
            }
        }

        private string get_IMCA_paramters(
            string sql_con,
            string param_name)
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
                logs_folder = AppDomain.CurrentDomain.BaseDirectory;
            }

            if (!Directory.Exists(logs_folder))
            {
                Directory.CreateDirectory(logs_folder);
            }

            string filePath = Path.Combine(
                logs_folder,
                "IMCA_" + global_session_name + "_" +
                DateTime.Now.ToString("dd_MM_yyyy") + "_" +
                country + "_" + global_application_name + ".txt");

            File.AppendAllText(
                filePath,
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") +
                " - " + message + Environment.NewLine);
        }

        private string EscapeODataString(string value)
        {
            return value == null ? "" : value.Replace("'", "''");
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
                    if (Path.GetFileName(file)
                        .ToLower()
                        .StartsWith(prefixfile.ToLower()))
                    {
                        File.Delete(file);
                    }
                }
                catch
                {
                    // Ignore cleanup errors so they do not block mailbox processing.
                }
            }
        }

        public void setParamCountry(string value) { country = value ?? ""; }
        public void setParamSK_Valid(string value) { sk_valid = value ?? ""; }
        public void setParamName(string value) { name = value ?? ""; }
        public void setParamActive(string value) { active = value ?? ""; }
        public void setParamDebug(string value) { debug = value ?? ""; }
        public void setStartDateScan(string value) { start_date_scan = value ?? ""; }

        public void setNumber_of_mails(string value)
        {
            number_of_mails = string.IsNullOrWhiteSpace(value) ? "10" : value;
        }

        public void setSharedMailboxName(string value)
        {
            sharedmailbox_name = value ?? "";
        }

        public void setsharedmailbox_folder_in(string value)
        {
            sharedmailbox_folder_in =
                string.IsNullOrWhiteSpace(value) ? "INBOX" : value;
        }

        public void setsharedmailbox_folder_out(string value)
        {
            sharedmailbox_folder_out =
                string.IsNullOrWhiteSpace(value) ? "Archives" : value;
        }

        public void setTableName(string value) { table_name = value ?? ""; }
        public void setTypeDatas(string value) { type_datas = value ?? ""; }
        public void setHasHeader(bool value) { has_header = value; }
        public void setColumnMapping(List<string> value)
        {
            column_mapping = value ?? new List<string>();
        }
        public void setMandatoryColumns(List<string> value)
        {
            mandatory_columns = value ?? new List<string>();
        }
        public void setSharemailboxActive(string value) { active = value ?? ""; }

        public void setEmailDestinataire(string value) { email_destinataire = value ?? ""; }
        public void setlogs_folder(string value) { logs_folder = value ?? ""; }
        public void settemp_folder(string value) { temp_folder = value ?? ""; }

        public void setSqlConnexionParam(string value)
        {
            sql_connexion_parameter_global = value ?? "";
        }

        public void setEmailInCaseOfTechnicalIssueParam(string value)
        {
            email_in_case_of_technical_issue_parameter_global = value ?? "";
        }
    }
}
