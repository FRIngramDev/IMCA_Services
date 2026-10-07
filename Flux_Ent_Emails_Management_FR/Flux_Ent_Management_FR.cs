using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using Microsoft.Graph.Models.ODataErrors;
using Newtonsoft.Json;
using Oracle.ManagedDataAccess.Client;
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

namespace FLUX_ENT_EMAILS_MANAGEMENT_FR
{
    public class FLUX_ENT_EMAILS_MANAGEMENT_FR
    {
        private const string global_application_name = "FLUX_ENT_EMAILS_MANAGEMENT_FR";
        private const string immutable_id_preference = "IdType=\"ImmutableId\"";
        private const string processed_property_id =
            "String {8BF48C6E-2C72-46F0-965D-919A7C2E54A9} Name IMCAFluxEntrantProcessed";

        private string country = "", skValid = "", name = "", active = "", debug = "";
        private string numberOfMails = "20", startDateScan = "";
        private string mailboxAddress = "", inputFolderName = "Inbox";
        private string archiveFolderName = "Archives";
        private string sqlConnexionParameterGlobal = "", sqlConnexion = "";
        private string oracleConnexionParameterGlobal = "", oracleConnexion = "";
        private string emailTechnicalParameterGlobal = "", emailTechnical = "";
        private string graphSendAsParameterGlobal = "", graphSendAs = "";
        private string logsFolder = "", tempFolder = "", sessionName = "";
        private GraphServiceClient graphService;
        private static readonly object logSyncRoot = new object();

        public sealed class JsonFile
        {
            public List<Country> countries { get; set; } = new List<Country>();
        }

        public sealed class Country
        {
            public string country { get; set; } = "";
            public string sk_valid { get; set; } = "";
            public string name { get; set; } = "";
            public string active { get; set; } = "TRUE";
            public string debug { get; set; } = "FALSE";
            public string number_of_mails { get; set; } = "20";
            public string start_date_scan { get; set; } = "";
            public string mailbox_address { get; set; } = "";
            public string sharedmailbox_folder_in { get; set; } = "Inbox";
            public string sharedmailbox_folder_out { get; set; } = "Archives";
            public string sql_connexion_parameter_global { get; set; } = "";
            public string oracle_connexion_parameter_global { get; set; } = "";
            public string email_in_case_of_technical_issue_parameter_global { get; set; } = "";
            public string fr_graph_send_as_parameter_global { get; set; } = "";
        }

        public void Close_Completed_Appointments(
            string sql_con,
            string logs,
            string tmp_folder,
            string session_name)
        {
            string root = GetServicePath();
            logsFolder = Path.Combine(root, logs ?? "");
            tempFolder = Path.Combine(root, tmp_folder ?? "");
            sessionName = session_name ?? "";

            try
            {
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
                    if (!IsTrue(active))
                        continue;


                    try
                    {
                        ValidateCommonConfiguration(false);
                        ValidateOracleConfiguration();
                        Directory.CreateDirectory(tempFolder);

                        WriteLog(
                            "   Working folder uses IMCA temp folder : " +
                            tempFolder);
                        WriteLog(
                            "   Starting completed appointment closure");

                        CloseCompletedAppointments();

                        WriteLog(
                            "   Completed appointment closure action completed" +
                            " - Country : " + country.ToUpperInvariant());
                    }
                    catch (Exception ex)
                    {
                        string details = GetDetailedExceptionMessage(ex);
                        WriteLog(
                            "   Close_Completed_Appointments error" +
                            " - Country : " + country.ToUpperInvariant() +
                            " - Details : " + details);
                        try
                        {
                            WriteLog(
                                "   Connecting to Microsoft Graph for " +
                                "completed appointment closure technical alert");
                            graphService = ConnectGraph();
                            TrySendTechnicalAlert(
                                nameof(Close_Completed_Appointments),
                                details);
                        }
                        catch (Exception alertException)
                        {
                            WriteLog(
                                "   Completed appointment closure technical " +
                                "alert could not be sent" +
                                " - Details : " +
                                GetDetailedExceptionMessage(alertException));
                        }
                        finally
                        {
                            graphService = null;
                        }
                    }
                }
            }
            finally
            {
                graphService = null;
            }
        }

        public void Read_Email_with_Graph(
            string sql_con,
            string logs,
            string tmp_folder,
            string session_name)
        {
            string root = GetServicePath();
            logsFolder = Path.Combine(root, logs ?? "");
            tempFolder = Path.Combine(root, tmp_folder ?? "");
            sessionName = session_name ?? "";

            try
            {
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
                    if (!IsTrue(active))
                        continue;

                    WriteLog(
                        name.ToUpperInvariant() +
                        "(" + country.ToUpperInvariant() + ") at " +
                        DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                    WriteLog(
                        "   Debug Parameter is set to " +
                        debug.ToUpperInvariant());

                    ValidateCommonConfiguration(true);
                    Directory.CreateDirectory(tempFolder);
                    WriteLog(
                        "   Working folder uses IMCA temp folder : " +
                        tempFolder);
                    WriteLog("   Connecting to Microsoft Graph");
                    graphService = ConnectGraph();
                    WriteLog("   Microsoft Graph connection established");

                    var mailboxErrors = new List<Exception>();
                    try
                    {
                        WriteLog(
                            "   Mailbox loaded from JSON configuration" +
                            " - " + mailboxAddress);
                        WriteLog(
                            "       Input folder : " + inputFolderName +
                            " - Output folder : " + archiveFolderName +
                            " - Maximum messages : " + numberOfMails);
                        if (IsTrue(debug))
                        {
                            WriteLog(
                                "       Filter date : " +
                                ParseStartDate().ToString(
                                    "dd/MM/yyyy HH:mm:ss"));
                        }

                        try
                        {
                            WriteLog(
                                "   Processing mailbox Flux Entrant : " +
                                mailboxAddress);
                            ProcessIncomingMessages();
                            WriteLog(
                                "   Mailbox processing completed : " +
                                mailboxAddress);
                        }
                        catch (Exception ex)
                        {
                            mailboxErrors.Add(ex);
                            string details = GetDetailedExceptionMessage(ex);
                            WriteLog(
                                "   Error reading mailbox" +
                                " - Mailbox : " + mailboxAddress +
                                " - Folder : " + inputFolderName +
                                " - Details : " + details);
                            TrySendTechnicalAlert(
                                nameof(Read_Email_with_Graph),
                                mailboxAddress +
                                " - Folder : " + inputFolderName +
                                " - " + details);
                        }
                    }
                    finally
                    {
                        graphService = null;
                    }

                    if (mailboxErrors.Count > 0)
                    {
                        WriteLog(
                            "   Country processing completed with " +
                            mailboxErrors.Count +
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

        private void CloseCompletedAppointments()
        {
            const string oracleSql = @"
SELECT COMPANY,DC_ID,CNTL_TYPE,CNTL_ID,ZONE_CODE,ORIG_LOCATION_ID,ASSIGN_SEQ,
       SPLIT_SEQ,STATUS,ONDOCK_ALLOC_SW,HOLD_REASON_CODE,PO_DC_ID,PO_ID,
       PO_LINE_NBR,CONF_LOCATION_ID,LOT_SERIAL_ID,QTY_ORIG_PUT,QTY_CONF_PUT,
       PART_ID,INV_IND,TC_LOCATION_ID,TC_QTY_PUT,TC_CODE,LOT_SERIAL_SW,USER_ID,
       DATE_CREATED,TIME_CREATED,DATE_CONFIRMED,TIME_CONFIRMED,RECEIPT_DATE,
       EXPIRATION_DATE,OVERFLOW_SW,MANUAL_PUT_SW,EXP_DATE_SW,SIZE_CODE,
       STORAGE_CODE,RESTRICTION_CODE,MOVEMENT,BIN_TYPE,XREF_SO_SW,
       DISCREPANCY_SW,GRID_ROW,GRID_COLUMN_NBR,ASN_ID,PACKAGE_ID,QTY_ASSIGN_ADJ,
       START_DATE,START_TIME,REC_SEQ_NBR,TO_DROP_POINT,VENDOR_ID,
       SERIAL_TRACKING_SW,RMA_BR_ID,RMA_ID,SLOC_ID,CUST_PO_LINE_NBR,
       END_USER_RMA,ORIGINAL_PART
FROM CSI.PT_ASSIGN_DETAIL";

            using (var oracleConnection = new OracleConnection(oracleConnexion))
            using (var oracleCommand = new OracleCommand(oracleSql, oracleConnection))
            using (var sqlConnection = new SqlConnection(sqlConnexion))
            {
                oracleCommand.CommandTimeout = 0;
                oracleConnection.Open();
                sqlConnection.Open();

                using (var transaction = sqlConnection.BeginTransaction())
                {
                    try
                    {
                        using (var truncate = new SqlCommand(
                            "TRUNCATE TABLE dbo.T_flux_entrant_RP_IMFIRST;",
                            sqlConnection, transaction))
                        {
                            truncate.CommandTimeout = 0;
                            truncate.ExecuteNonQuery();
                        }

                        int importedRows;
                        using (OracleDataReader reader = oracleCommand.ExecuteReader())
                        using (var bulkCopy = new SqlBulkCopy(
                            sqlConnection, SqlBulkCopyOptions.TableLock, transaction))
                        {
                            bulkCopy.DestinationTableName = "dbo.T_flux_entrant_RP_IMFIRST";
                            bulkCopy.BulkCopyTimeout = 0;
                            bulkCopy.BatchSize = 10000;
                            bulkCopy.WriteToServer(reader);
                        }

                        using (var count = new SqlCommand(
                            "SELECT COUNT(*) FROM dbo.T_flux_entrant_RP_IMFIRST;",
                            sqlConnection, transaction))
                        {
                            importedRows = Convert.ToInt32(count.ExecuteScalar());
                        }

                        int closedAppointments;
                        using (var close = new SqlCommand(@"
UPDATE appointment
SET id_statut=130,
    date_heure_rangement_termine=GETDATE()
FROM dbo.T_flux_entrant_rdv AS appointment
WHERE appointment.id_statut=100
  AND NOT EXISTS
  (
      SELECT 1
      FROM dbo.T_flux_entrant_rdv_detail AS detail
      WHERE detail.id_rdv=appointment.id_rdv
        AND EXISTS
        (
            SELECT 1
            FROM dbo.T_flux_entrant_RP_IMFIRST AS imfirst
            WHERE imfirst.DATE_CONFIRMED IS NULL
              AND ISNULL(imfirst.USER_ID,'')=''
              AND imfirst.PO_ID LIKE '%' + RIGHT(detail.num_PO,5) + '%'
        )
  );", sqlConnection, transaction))
                        {
                            close.CommandTimeout = 0;
                            closedAppointments = close.ExecuteNonQuery();
                        }

                        transaction.Commit();
                        WriteLog(
                            "   Completed appointment closure finished" +
                            " - Oracle row(s) imported : " + importedRows +
                            " - Appointment(s) closed : " + closedAppointments);
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw;
                    }
                }
            }
        }

        private sealed class FluxMailProcessingState
        {
            public string Status { get; set; } = "";
            public int? FileId { get; set; }
        }

        private void ProcessIncomingMessages()
        {
            WriteLog("   Processing incoming mailbox : " + mailboxAddress);
            MailFolder inbox = GetRequiredFolder("inbox", inputFolderName);
            MailFolder archive = GetRequiredFolder(inbox.Id, archiveFolderName);
            List<Message> messages = GetInboxMessages(inbox.Id);
            int imported = 0;
            int resumed = 0;
            int alreadyProcessed = 0;
            int errors = 0;

            WriteLog("       Email(s) found for processing : " + messages.Count);

            foreach (Message summary in messages)
            {
                Message email = null;

                try
                {
                    FluxMailProcessingState state =
                        GetFluxMailProcessingState(summary.Id);

                    if (state == null && IsMessageProcessed(summary))
                    {
                        UpsertLegacyProcessedMessage(summary);
                        EnsureProcessedMessageFinalized(
                            summary.Id,
                            archive.Id);
                        alreadyProcessed++;

                        WriteLog(
                            "       Legacy processed email registered in " +
                            "tracking table" +
                            " - Subject : " +
                            (summary.Subject ?? "<no subject>"));

                        continue;
                    }

                    if (state != null &&
                        state.Status.Equals(
                            "S",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        EnsureProcessedMessageFinalized(
                            summary.Id,
                            archive.Id);

                        alreadyProcessed++;

                        WriteLog(
                            "       Incoming email already imported; " +
                            "Graph finalization checked" +
                            " - File ID : " +
                            (state.FileId.HasValue
                                ? state.FileId.Value.ToString(
                                    CultureInfo.InvariantCulture)
                                : "") +
                            " - Subject : " +
                            (summary.Subject ?? "<no subject>"));

                        continue;
                    }

                    email = GetCompleteMessage(summary.Id);
                    int fileId;

                    if (state != null && state.FileId.HasValue)
                    {
                        fileId = state.FileId.Value;
                        resumed++;

                        WriteLog(
                            "       Resuming incoming email after previous error" +
                            " - File ID : " + fileId +
                            " - Subject : " +
                            (email.Subject ?? "<no subject>"));
                    }
                    else
                    {
                        byte[] mime = GetMimeContent(email.Id);
                        fileId = InsertFluxFileAndTrack(email, mime);
                    }

                    ApplyBusinessClassification(
                        fileId,
                        email.Subject ?? "");

                    MarkMessageAsProcessed(email.Id);

                    SetFluxMailProcessingSuccess(
                        email,
                        fileId);

                    MoveMessage(email.Id, archive.Id);

                    imported++;

                    WriteLog(
                        "       Incoming email imported and archived" +
                        " - File ID : " + fileId +
                        " - Sender : " + GetSender(email) +
                        " - Subject : " +
                        (email.Subject ?? "<no subject>"));
                }
                catch (Exception ex)
                {
                    errors++;

                    Message trackedMessage = email ?? summary;
                    UpsertFluxMailProcessingError(
                        trackedMessage,
                        ex);

                    WriteLog(
                        "       Incoming email processing error" +
                        " - Subject : " +
                        (summary.Subject ?? "<no subject>") +
                        " - Details : " +
                        GetDetailedExceptionMessage(ex));

                    TrySendTechnicalAlert(
                        nameof(ProcessIncomingMessages),
                        mailboxAddress + " - Mail : " +
                        (summary.Subject ?? "<no subject>") + " - " +
                        GetDetailedExceptionMessage(ex));
                }
            }

            WriteLog(
                "       Incoming mailbox summary" +
                " - Found : " + messages.Count +
                " - Imported : " + imported +
                " - Resumed : " + resumed +
                " - Already processed : " + alreadyProcessed +
                " - Errors : " + errors);
        }

        private int InsertFluxFileAndTrack(
            Message email,
            byte[] mime)
        {
            if (mime == null || mime.Length == 0)
                throw new InvalidDataException(
                    "Incoming MIME content is empty");

            DateTime received =
                email.ReceivedDateTime?.LocalDateTime ?? DateTime.Now;

            string fileName =
                "EMAIL_" +
                received.ToString(
                    "yyyyMMdd_HHmmss",
                    CultureInfo.InvariantCulture);

            using (var connection = new SqlConnection(sqlConnexion))
            {
                connection.Open();

                using (SqlTransaction transaction =
                    connection.BeginTransaction())
                {
                    try
                    {
                        int fileId;

                        using (var command = new SqlCommand(
                            "dbo.USP_FLUX_ENTRANT_ADD_FIC",
                            connection,
                            transaction))
                        {
                            command.CommandType =
                                CommandType.StoredProcedure;
                            command.CommandTimeout = 300;

                            command.Parameters.Add(
                                "@id_rdv",
                                SqlDbType.NVarChar).Value = "0";
                            command.Parameters.Add(
                                "@nom_fic",
                                SqlDbType.NVarChar).Value = fileName;
                            command.Parameters.Add(
                                "@document",
                                SqlDbType.Image).Value = mime;
                            command.Parameters.Add(
                                "@extension",
                                SqlDbType.NVarChar).Value = "eml";
                            command.Parameters.Add(
                                "@user_cn",
                                SqlDbType.NVarChar).Value = "cmr";
                            command.Parameters.Add(
                                "@typologie",
                                SqlDbType.NVarChar).Value = "CMR";
                            command.Parameters.Add(
                                "@num_litige",
                                SqlDbType.NVarChar).Value = "";

                            SqlParameter outputId =
                                command.Parameters.Add(
                                    "@@num",
                                    SqlDbType.Int);
                            outputId.Direction =
                                ParameterDirection.Output;

                            command.ExecuteNonQuery();

                            if (outputId.Value == null ||
                                outputId.Value == DBNull.Value)
                            {
                                throw new InvalidOperationException(
                                    "USP_FLUX_ENTRANT_ADD_FIC returned " +
                                    "no file ID");
                            }

                            fileId = Convert.ToInt32(outputId.Value);
                        }

                        UpsertFluxMailProcessingPending(
                            connection,
                            transaction,
                            email,
                            fileId);

                        transaction.Commit();
                        return fileId;
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw;
                    }
                }
            }
        }

        private FluxMailProcessingState GetFluxMailProcessingState(
            string graphMessageId)
        {
            const string sql = @"
SELECT ProcessingStatus,
       FluxFileId
FROM dbo.T_FluxEntrantMailProcessing
WHERE MailboxAddress=@MAILBOX
  AND GraphMessageId=@MESSAGE;";

            using (var connection = new SqlConnection(sqlConnexion))
            using (var command = new SqlCommand(sql, connection))
            {
                command.CommandTimeout = 300;
                command.Parameters.Add(
                    "@MAILBOX",
                    SqlDbType.NVarChar,
                    320).Value = mailboxAddress ?? "";
                command.Parameters.Add(
                    "@MESSAGE",
                    SqlDbType.NVarChar,
                    500).Value = graphMessageId ?? "";

                connection.Open();

                using (SqlDataReader reader = command.ExecuteReader())
                {
                    if (!reader.Read())
                        return null;

                    return new FluxMailProcessingState
                    {
                        Status = Convert.ToString(
                            reader["ProcessingStatus"]).Trim(),
                        FileId = reader["FluxFileId"] == DBNull.Value
                            ? (int?)null
                            : Convert.ToInt32(reader["FluxFileId"])
                    };
                }
            }
        }

        private void UpsertFluxMailProcessingPending(
            SqlConnection connection,
            SqlTransaction transaction,
            Message email,
            int fileId)
        {
            const string sql = @"
MERGE dbo.T_FluxEntrantMailProcessing AS target
USING
(
    SELECT @MAILBOX AS MailboxAddress,
           @GRAPH_MESSAGE_ID AS GraphMessageId
) AS source
ON target.MailboxAddress=source.MailboxAddress
AND target.GraphMessageId=source.GraphMessageId
WHEN MATCHED THEN
    UPDATE SET InternetMessageId=@INTERNET_MESSAGE_ID,
               MessageSubject=@SUBJECT,
               ReceivedDateTimeUtc=@RECEIVED_UTC,
               FluxFileId=@FILE_ID,
               ProcessingStatus='P',
               ErrorMessage=NULL,
               AttemptCount=target.AttemptCount+1,
               LastAttemptAtUtc=SYSUTCDATETIME()
WHEN NOT MATCHED THEN
    INSERT
    (
        MailboxAddress,
        GraphMessageId,
        InternetMessageId,
        MessageSubject,
        ReceivedDateTimeUtc,
        FluxFileId,
        ProcessingStatus,
        ErrorMessage,
        AttemptCount,
        FirstSeenAtUtc,
        LastAttemptAtUtc
    )
    VALUES
    (
        @MAILBOX,
        @GRAPH_MESSAGE_ID,
        @INTERNET_MESSAGE_ID,
        @SUBJECT,
        @RECEIVED_UTC,
        @FILE_ID,
        'P',
        NULL,
        1,
        SYSUTCDATETIME(),
        SYSUTCDATETIME()
    );";

            using (var command = new SqlCommand(
                sql,
                connection,
                transaction))
            {
                AddFluxTrackingParameters(
                    command,
                    email,
                    fileId);
                command.ExecuteNonQuery();
            }
        }

        private void SetFluxMailProcessingSuccess(
            Message email,
            int fileId)
        {
            ExecuteNonQuery(
                sqlConnexion,
                @"
UPDATE dbo.T_FluxEntrantMailProcessing
SET InternetMessageId=@INTERNET_MESSAGE_ID,
    MessageSubject=@SUBJECT,
    ReceivedDateTimeUtc=@RECEIVED_UTC,
    FluxFileId=@FILE_ID,
    ProcessingStatus='S',
    ErrorMessage=NULL,
    LastAttemptAtUtc=SYSUTCDATETIME(),
    ProcessedAtUtc=SYSUTCDATETIME()
WHERE MailboxAddress=@MAILBOX
  AND GraphMessageId=@GRAPH_MESSAGE_ID;",
                BuildFluxTrackingParameters(email, fileId));
        }

        private void UpsertLegacyProcessedMessage(Message email)
        {
            ExecuteNonQuery(
                sqlConnexion,
                @"
MERGE dbo.T_FluxEntrantMailProcessing AS target
USING
(
    SELECT @MAILBOX AS MailboxAddress,
           @GRAPH_MESSAGE_ID AS GraphMessageId
) AS source
ON target.MailboxAddress=source.MailboxAddress
AND target.GraphMessageId=source.GraphMessageId
WHEN MATCHED THEN
    UPDATE SET InternetMessageId=@INTERNET_MESSAGE_ID,
               MessageSubject=@SUBJECT,
               ReceivedDateTimeUtc=@RECEIVED_UTC,
               ProcessingStatus='S',
               ErrorMessage=NULL,
               LastAttemptAtUtc=SYSUTCDATETIME(),
               ProcessedAtUtc=COALESCE(
                   target.ProcessedAtUtc,
                   SYSUTCDATETIME())
WHEN NOT MATCHED THEN
    INSERT
    (
        MailboxAddress,
        GraphMessageId,
        InternetMessageId,
        MessageSubject,
        ReceivedDateTimeUtc,
        FluxFileId,
        ProcessingStatus,
        ErrorMessage,
        AttemptCount,
        FirstSeenAtUtc,
        LastAttemptAtUtc,
        ProcessedAtUtc
    )
    VALUES
    (
        @MAILBOX,
        @GRAPH_MESSAGE_ID,
        @INTERNET_MESSAGE_ID,
        @SUBJECT,
        @RECEIVED_UTC,
        NULL,
        'S',
        NULL,
        1,
        SYSUTCDATETIME(),
        SYSUTCDATETIME(),
        SYSUTCDATETIME()
    );",
                new SqlParameter(
                    "@MAILBOX",
                    SqlDbType.NVarChar,
                    320)
                {
                    Value = mailboxAddress ?? ""
                },
                new SqlParameter(
                    "@GRAPH_MESSAGE_ID",
                    SqlDbType.NVarChar,
                    500)
                {
                    Value = email?.Id ?? ""
                },
                new SqlParameter(
                    "@INTERNET_MESSAGE_ID",
                    SqlDbType.NVarChar,
                    500)
                {
                    Value = email?.InternetMessageId ?? ""
                },
                new SqlParameter(
                    "@SUBJECT",
                    SqlDbType.NVarChar,
                    1000)
                {
                    Value = Truncate(email?.Subject, 1000)
                },
                new SqlParameter(
                    "@RECEIVED_UTC",
                    SqlDbType.DateTime2)
                {
                    Value = email?.ReceivedDateTime?.UtcDateTime
                        ?? DateTime.UtcNow
                });
        }

        private void UpsertFluxMailProcessingError(
            Message email,
            Exception exception)
        {
            if (email == null ||
                string.IsNullOrWhiteSpace(email.Id))
            {
                return;
            }

            FluxMailProcessingState currentState =
                GetFluxMailProcessingState(email.Id);

            int? fileId = currentState?.FileId;
            string details = Truncate(
                GetDetailedExceptionMessage(exception),
                2000);

            ExecuteNonQuery(
                sqlConnexion,
                @"
MERGE dbo.T_FluxEntrantMailProcessing AS target
USING
(
    SELECT @MAILBOX AS MailboxAddress,
           @GRAPH_MESSAGE_ID AS GraphMessageId
) AS source
ON target.MailboxAddress=source.MailboxAddress
AND target.GraphMessageId=source.GraphMessageId
WHEN MATCHED THEN
    UPDATE SET InternetMessageId=@INTERNET_MESSAGE_ID,
               MessageSubject=@SUBJECT,
               ReceivedDateTimeUtc=@RECEIVED_UTC,
               FluxFileId=COALESCE(@FILE_ID,target.FluxFileId),
               ProcessingStatus='E',
               ErrorMessage=@ERROR,
               AttemptCount=target.AttemptCount+1,
               LastAttemptAtUtc=SYSUTCDATETIME()
WHEN NOT MATCHED THEN
    INSERT
    (
        MailboxAddress,
        GraphMessageId,
        InternetMessageId,
        MessageSubject,
        ReceivedDateTimeUtc,
        FluxFileId,
        ProcessingStatus,
        ErrorMessage,
        AttemptCount,
        FirstSeenAtUtc,
        LastAttemptAtUtc
    )
    VALUES
    (
        @MAILBOX,
        @GRAPH_MESSAGE_ID,
        @INTERNET_MESSAGE_ID,
        @SUBJECT,
        @RECEIVED_UTC,
        @FILE_ID,
        'E',
        @ERROR,
        1,
        SYSUTCDATETIME(),
        SYSUTCDATETIME()
    );",
                new SqlParameter(
                    "@MAILBOX",
                    SqlDbType.NVarChar,
                    320)
                {
                    Value = mailboxAddress ?? ""
                },
                new SqlParameter(
                    "@GRAPH_MESSAGE_ID",
                    SqlDbType.NVarChar,
                    500)
                {
                    Value = email.Id ?? ""
                },
                new SqlParameter(
                    "@INTERNET_MESSAGE_ID",
                    SqlDbType.NVarChar,
                    500)
                {
                    Value = email.InternetMessageId ?? ""
                },
                new SqlParameter(
                    "@SUBJECT",
                    SqlDbType.NVarChar,
                    1000)
                {
                    Value = Truncate(email.Subject, 1000)
                },
                new SqlParameter(
                    "@RECEIVED_UTC",
                    SqlDbType.DateTime2)
                {
                    Value = email.ReceivedDateTime?.UtcDateTime
                        ?? DateTime.UtcNow
                },
                new SqlParameter(
                    "@FILE_ID",
                    SqlDbType.Int)
                {
                    Value = fileId.HasValue
                        ? (object)fileId.Value
                        : DBNull.Value
                },
                new SqlParameter(
                    "@ERROR",
                    SqlDbType.NVarChar,
                    2000)
                {
                    Value = details
                });
        }

        private void EnsureProcessedMessageFinalized(
            string messageId,
            string archiveFolderId)
        {
            try
            {
                MarkMessageAsProcessed(messageId);
                MoveMessage(messageId, archiveFolderId);
            }
            catch (Exception ex) when (IsGraphObjectNotFound(ex))
            {
                WriteLog(
                    "       Already processed Graph message is no longer " +
                    "available : " + messageId);
            }
        }

        private void AddFluxTrackingParameters(
            SqlCommand command,
            Message email,
            int fileId)
        {
            foreach (SqlParameter parameter in
                BuildFluxTrackingParameters(email, fileId))
            {
                command.Parameters.Add(parameter);
            }
        }

        private SqlParameter[] BuildFluxTrackingParameters(
            Message email,
            int fileId)
        {
            return new[]
            {
                new SqlParameter(
                    "@MAILBOX",
                    SqlDbType.NVarChar,
                    320)
                {
                    Value = mailboxAddress ?? ""
                },
                new SqlParameter(
                    "@GRAPH_MESSAGE_ID",
                    SqlDbType.NVarChar,
                    500)
                {
                    Value = email?.Id ?? ""
                },
                new SqlParameter(
                    "@INTERNET_MESSAGE_ID",
                    SqlDbType.NVarChar,
                    500)
                {
                    Value = email?.InternetMessageId ?? ""
                },
                new SqlParameter(
                    "@SUBJECT",
                    SqlDbType.NVarChar,
                    1000)
                {
                    Value = Truncate(email?.Subject, 1000)
                },
                new SqlParameter(
                    "@RECEIVED_UTC",
                    SqlDbType.DateTime2)
                {
                    Value = email?.ReceivedDateTime?.UtcDateTime
                        ?? DateTime.UtcNow
                },
                new SqlParameter(
                    "@FILE_ID",
                    SqlDbType.Int)
                {
                    Value = fileId
                }
            };
        }

        private void ApplyBusinessClassification(int fileId, string subject)
        {
            int? appointmentId = TryExtractAppointmentId(subject);
            if (appointmentId.HasValue)
            {
                ExecuteNonQuery(sqlConnexion, @"
UPDATE dbo.T_flux_entrant_fichiers_joints
SET id_rdv=@RDV,
    typologie='RDV_TRANSPORTEUR'
WHERE id_fichier=@FILE;",
                    new SqlParameter("@RDV", SqlDbType.Int) { Value = appointmentId.Value },
                    new SqlParameter("@FILE", SqlDbType.Int) { Value = fileId });
            }

            string disputeNumber = TryExtractBookingDisputeNumber(subject);
            if (!string.IsNullOrWhiteSpace(disputeNumber))
            {
                string appointmentNumber = disputeNumber.Split('-').FirstOrDefault() ?? "";
                if (!int.TryParse(appointmentNumber, NumberStyles.None,
                        CultureInfo.InvariantCulture, out int disputeAppointmentId))
                {
                    throw new InvalidDataException(
                        "Invalid appointment number in booking dispute subject : " + subject);
                }

                ExecuteNonQuery(sqlConnexion, @"
UPDATE dbo.T_flux_entrant_fichiers_joints
SET id_rdv=@RDV,
    typologie='LITIGE',
    num_litige=@DISPUTE
WHERE id_fichier=@FILE;",
                    new SqlParameter("@RDV", SqlDbType.Int) { Value = disputeAppointmentId },
                    new SqlParameter("@DISPUTE", SqlDbType.NVarChar, 200) { Value = disputeNumber },
                    new SqlParameter("@FILE", SqlDbType.Int) { Value = fileId });
            }
        }

        private static int? TryExtractAppointmentId(string subject)
        {
            Match match = Regex.Match(
                subject ?? "", @"RDV\s*n[°o]?\s*(\d+)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (!match.Success) return null;
            return int.TryParse(match.Groups[1].Value, NumberStyles.None,
                CultureInfo.InvariantCulture, out int value) ? value : (int?)null;
        }

        private static string TryExtractBookingDisputeNumber(string subject)
        {
            Match match = Regex.Match(
                subject ?? "",
                @"Ingram Micro\s*-\s*Booking Problems\s+(.+?)\s+upon booking",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            return match.Success ? match.Groups[1].Value.Trim() : "";
        }

        private List<Message> GetInboxMessages(string folderId)
        {
            int top = ParsePositiveInt(numberOfMails, 20);
            DateTime start = ParseStartDate();

            WriteLog(
                "       Listing Graph messages" +
                " - Mailbox : " + mailboxAddress +
                " - Folder : " + inputFolderName +
                " - Filter date : " +
                start.ToString("dd/MM/yyyy HH:mm:ss") +
                " - Top : " + top);

            MessageCollectionResponse response = ExecuteGraphWithRetry(
                () => graphService.Users[mailboxAddress]
                    .MailFolders[folderId].Messages.GetAsync(config =>
                    {
                        AddImmutableHeader(config.Headers);
                        config.QueryParameters.Top = top;
                        config.QueryParameters.Orderby = new[] { "receivedDateTime asc" };
                        config.QueryParameters.Filter =
                            "isRead eq false and receivedDateTime gt " +
                            start.ToUniversalTime().ToString(
                                "yyyy-MM-ddTHH:mm:ss.fffZ", CultureInfo.InvariantCulture);
                        config.QueryParameters.Select = new[]
                        {
                            "id", "subject", "receivedDateTime", "from", "internetMessageId"
                        };
                        config.QueryParameters.Expand = new[]
                        {
                            "singleValueExtendedProperties($filter=id eq '" +
                            EscapeODataString(processed_property_id) + "')"
                        };
                    }).GetAwaiter().GetResult(),
                "List incoming flux messages");
            return (response?.Value ?? new List<Message>())
                .Where(message => !string.IsNullOrWhiteSpace(message.Id))
                .OrderBy(message => message.ReceivedDateTime)
                .ToList();
        }

        private Message GetCompleteMessage(string id)
        {
            return ExecuteGraphWithRetry(
                () => graphService.Users[mailboxAddress].Messages[id]
                    .GetAsync(config =>
                    {
                        AddImmutableHeader(config.Headers);
                        config.QueryParameters.Select = new[]
                        {
                            "id", "subject", "receivedDateTime", "createdDateTime",
                            "from", "sender", "internetMessageId"
                        };
                    }).GetAwaiter().GetResult(),
                "Get complete incoming flux message");
        }

        private byte[] GetMimeContent(string id)
        {
            using (Stream input = ExecuteGraphWithRetry(
                () => graphService.Users[mailboxAddress].Messages[id].Content
                    .GetAsync(config => AddImmutableHeader(config.Headers))
                    .GetAwaiter().GetResult(),
                "Get incoming flux MIME"))
            using (var output = new MemoryStream())
            {
                if (input == null) throw new InvalidDataException("Graph MIME stream is empty");
                input.CopyTo(output);
                return output.ToArray();
            }
        }

        private void MarkMessageAsProcessed(string id)
        {
            ExecuteGraphWithRetry(() =>
            {
                var update = new Message
                {
                    IsRead = true,
                    SingleValueExtendedProperties =
                        new List<SingleValueLegacyExtendedProperty>
                        {
                            new SingleValueLegacyExtendedProperty
                            {
                                Id = processed_property_id,
                                Value = DateTime.UtcNow.ToString(
                                    "yyyy-MM-ddTHH:mm:ss.fffZ", CultureInfo.InvariantCulture)
                            }
                        }
                };
                graphService.Users[mailboxAddress].Messages[id]
                    .PatchAsync(update, config => AddImmutableHeader(config.Headers))
                    .GetAwaiter().GetResult();
                return true;
            }, "Mark incoming flux message as processed");
        }

        private void MoveMessage(string id, string destinationFolderId)
        {
            ExecuteGraphWithRetry(() =>
            {
                graphService.Users[mailboxAddress].Messages[id].Move.PostAsync(
                    new Microsoft.Graph.Users.Item.Messages.Item.Move.MovePostRequestBody
                    {
                        DestinationId = destinationFolderId
                    },
                    config => AddImmutableHeader(config.Headers))
                    .GetAwaiter().GetResult();
                return true;
            }, "Move incoming flux message");
        }

        private MailFolder GetRequiredFolder(string parentFolderId, string folderName)
        {
            if (folderName.Equals("Inbox", StringComparison.OrdinalIgnoreCase) ||
                parentFolderId.Equals("inbox", StringComparison.OrdinalIgnoreCase))
            {
                if (folderName.Equals("Inbox", StringComparison.OrdinalIgnoreCase))
                {
                    return ExecuteGraphWithRetry(
                        () => graphService.Users[mailboxAddress].MailFolders["inbox"]
                            .GetAsync(config => AddImmutableHeader(config.Headers))
                            .GetAwaiter().GetResult(),
                        "Read Inbox folder");
                }
            }

            MailFolderCollectionResponse response = ExecuteGraphWithRetry(
                () => graphService.Users[mailboxAddress].MailFolders[parentFolderId]
                    .ChildFolders.GetAsync(config =>
                    {
                        AddImmutableHeader(config.Headers);
                        config.QueryParameters.Top = 100;
                        config.QueryParameters.Filter =
                            "displayName eq '" + EscapeODataString(folderName) + "'";
                    }).GetAwaiter().GetResult(),
                "Find folder " + folderName);
            MailFolder folder = response?.Value?.FirstOrDefault(item =>
                string.Equals(item.DisplayName, folderName, StringComparison.OrdinalIgnoreCase));
            if (folder == null)
                throw new DirectoryNotFoundException(
                    "Folder not found under Inbox : " + folderName);
            return folder;
        }

        private bool IsMessageProcessed(Message message)
        {
            return message?.SingleValueExtendedProperties?.Any(property =>
                string.Equals(property.Id, processed_property_id,
                    StringComparison.OrdinalIgnoreCase) &&
                !string.IsNullOrWhiteSpace(property.Value)) == true;
        }

        private GraphServiceClient ConnectGraph()
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            return new class_dev_tools.Ews_Modern_Auth().Get_Graph_Service();
        }

        private T ExecuteGraphWithRetry<T>(Func<T> action, string operation)
        {
            const int maxAttempts = 3;
            Exception last = null;
            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                try { return action(); }
                catch (Exception ex) when (IsTransientGraphError(ex))
                {
                    last = ex;
                    if (attempt >= maxAttempts) break;
                    int delay = attempt * 5000;
                    WriteLog(
                        "       Temporary Graph error during " + operation +
                        " - Attempt " + attempt + "/" + maxAttempts +
                        " - Retry in " + delay + " ms" +
                        " - Details : " + GetDetailedExceptionMessage(ex));
                    System.Threading.Thread.Sleep(delay);
                }
            }
            throw new InvalidOperationException(
                "Graph operation failed after " + maxAttempts +
                " attempt(s) : " + operation + " - " +
                GetDetailedExceptionMessage(last), last);
        }

        private void TrySendTechnicalAlert(string method, string details)
        {
            try
            {
                if (graphService == null) graphService = ConnectGraph();
                List<Recipient> recipients = BuildRecipients(emailTechnical);
                if (recipients.Count == 0)
                    throw new InvalidOperationException("Technical alert recipient is invalid");

                var sender = new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = graphSendAs,
                        Name = "Ingram Micro - Flux Entrant Lomme"
                    }
                };
                var message = new Message
                {
                    Subject = "Flux Entrant Lomme - Erreur dans " + method,
                    From = sender,
                    Sender = sender,
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Html,
                        Content =
                            "<b>Application :</b> Flux Entrant Lomme<br/>" +
                            "<b>Mailbox :</b> " + WebUtility.HtmlEncode(mailboxAddress) + "<br/>" +
                            "<b>Operation :</b> " + WebUtility.HtmlEncode(method) + "<br/>" +
                            "<b>Message :</b> " + WebUtility.HtmlEncode(details)
                    },
                    ToRecipients = recipients
                };
                graphService.Users[graphSendAs].SendMail.PostAsync(
                    new Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody
                    {
                        Message = message,
                        SaveToSentItems = true
                    }).GetAwaiter().GetResult();
            }
            catch (Exception alertException)
            {
                WriteLog("   Alert error : " + GetDetailedExceptionMessage(alertException));
            }
        }

        private void ApplyCountryConfiguration(Country item, string imcaConnection)
        {
            country = item.country ?? "";
            skValid = item.sk_valid ?? "";
            name = item.name ?? "";
            active = item.active ?? "";
            debug = item.debug ?? "";
            numberOfMails = string.IsNullOrWhiteSpace(item.number_of_mails)
                ? "20" : item.number_of_mails;
            startDateScan = item.start_date_scan ?? "";
            mailboxAddress = item.mailbox_address ?? "";
            inputFolderName = string.IsNullOrWhiteSpace(item.sharedmailbox_folder_in)
                ? "Inbox" : item.sharedmailbox_folder_in;
            archiveFolderName = string.IsNullOrWhiteSpace(item.sharedmailbox_folder_out)
                ? "Archives" : item.sharedmailbox_folder_out;
            sqlConnexionParameterGlobal = item.sql_connexion_parameter_global ?? "";
            oracleConnexionParameterGlobal = item.oracle_connexion_parameter_global ?? "";
            emailTechnicalParameterGlobal =
                item.email_in_case_of_technical_issue_parameter_global ?? "";
            graphSendAsParameterGlobal = item.fr_graph_send_as_parameter_global ?? "";

            sqlConnexion = GetImcaParameter(imcaConnection, sqlConnexionParameterGlobal);
            oracleConnexion = GetImcaParameter(imcaConnection, oracleConnexionParameterGlobal);
            emailTechnical = GetImcaParameter(imcaConnection, emailTechnicalParameterGlobal);
            graphSendAs = GetImcaParameter(imcaConnection, graphSendAsParameterGlobal);
        }

        private void ValidateCommonConfiguration(bool mailboxRequired)
        {
            var missing = new List<string>();
            if (string.IsNullOrWhiteSpace(sqlConnexion)) missing.Add("sqlConnexion");
            if (string.IsNullOrWhiteSpace(emailTechnical)) missing.Add("emailTechnical");
            if (string.IsNullOrWhiteSpace(graphSendAs)) missing.Add("graphSendAs");

            if (mailboxRequired)
            {
                if (string.IsNullOrWhiteSpace(mailboxAddress)) missing.Add("mailboxAddress");
                if (string.IsNullOrWhiteSpace(inputFolderName)) missing.Add("inputFolderName");
                if (string.IsNullOrWhiteSpace(archiveFolderName)) missing.Add("archiveFolderName");
            }
            if (missing.Count > 0)
                throw new InvalidOperationException(
                    "Common configuration is incomplete : " + string.Join(", ", missing));
        }

        private void ValidateOracleConfiguration()
        {
            if (string.IsNullOrWhiteSpace(oracleConnexion))
            {
                throw new InvalidOperationException(
                    "oracleConnexion is empty for completed appointment closure");
            }
        }

        private void InitializeActionFolders(string logs, string temp, string session)
        {
            string root = GetServicePath();
            logsFolder = Path.Combine(root, logs ?? "");
            tempFolder = Path.Combine(root, temp ?? "");
            sessionName = session ?? "";
            Directory.CreateDirectory(logsFolder);
            Directory.CreateDirectory(tempFolder);
        }

        private string GetImcaParameter(string connectionString, string parameter)
        {
            if (string.IsNullOrWhiteSpace(parameter)) return "";
            return ExecuteScalarString(connectionString,
                "SELECT ISNULL(VALUE,'') FROM PCM_TAB_IMCA_PARAMETER_GLOBAL " +
                "WHERE SK_VALID=0 AND PARAMETER=@P;",
                new SqlParameter("@P", SqlDbType.NVarChar, 255) { Value = parameter });
        }

        private static void ExecuteNonQuery(
            string connectionString, string sql, params SqlParameter[] parameters)
        {
            using (var connection = new SqlConnection(connectionString))
            using (var command = new SqlCommand(sql, connection))
            {
                command.CommandTimeout = 300;
                if (parameters != null && parameters.Length > 0)
                    command.Parameters.AddRange(parameters);
                connection.Open();
                command.ExecuteNonQuery();
            }
        }

        private static string ExecuteScalarString(
            string connectionString, string sql, params SqlParameter[] parameters)
        {
            using (var connection = new SqlConnection(connectionString))
            using (var command = new SqlCommand(sql, connection))
            {
                command.CommandTimeout = 300;
                if (parameters != null && parameters.Length > 0)
                    command.Parameters.AddRange(parameters);
                connection.Open();
                return Convert.ToString(command.ExecuteScalar()).Trim();
            }
        }

        private void WriteLog(string message)
        {
            try
            {
                Directory.CreateDirectory(logsFolder);
                string path = Path.Combine(
                    logsFolder,
                    "IMCA_" + sessionName + "_" +
                    DateTime.Now.ToString("dd_MM_yyyy") + "_" + country + "_" +
                    global_application_name + ".txt");
                string line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") +
                              " - " + (message ?? "") + Environment.NewLine;
                lock (logSyncRoot)
                {
                    File.AppendAllText(path, line, Encoding.UTF8);
                }
            }
            catch
            {
                // Logging must never stop the business processing.
            }
        }

        private static string GetDetailedExceptionMessage(Exception exception)
        {
            if (exception == null) return "Unknown error";
            var details = new List<string>();
            Exception current = exception;
            while (current != null)
            {
                details.Add("Type=" + current.GetType().FullName);
                if (!string.IsNullOrWhiteSpace(current.Message))
                    details.Add("Message=" + current.Message);
                if (current is ODataError graphError)
                {
                    if (!string.IsNullOrWhiteSpace(graphError.Error?.Code))
                        details.Add("GraphCode=" + graphError.Error.Code);
                    if (!string.IsNullOrWhiteSpace(graphError.Error?.Message))
                        details.Add("GraphMessage=" + graphError.Error.Message);
                }
                if (current is ApiException apiException)
                    details.Add("HttpStatus=" + apiException.ResponseStatusCode);
                current = current.InnerException;
            }
            return string.Join(" | ", details.Distinct(StringComparer.OrdinalIgnoreCase));
        }

        private static bool IsGraphObjectNotFound(Exception exception)
        {
            Exception current = exception;

            while (current != null)
            {
                if (current is ApiException apiException &&
                    apiException.ResponseStatusCode == 404)
                {
                    return true;
                }

                string message = current.Message ?? "";

                if (message.IndexOf(
                        "not found",
                        StringComparison.OrdinalIgnoreCase) >= 0 ||
                    message.IndexOf(
                        "specified object was not found",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }

                current = current.InnerException;
            }

            return false;
        }

        private static string Truncate(string value, int maximumLength)
        {
            value = value ?? "";
            return value.Length <= maximumLength
                ? value
                : value.Substring(0, maximumLength);
        }

        private static bool IsTransientGraphError(Exception exception)
        {
            Exception current = exception;
            while (current != null)
            {
                if (current is HttpRequestException || current is TimeoutException ||
                    current is System.Threading.Tasks.TaskCanceledException)
                    return true;
                if (current is ApiException api &&
                    (api.ResponseStatusCode == 408 || api.ResponseStatusCode == 429 ||
                     api.ResponseStatusCode == 500 || api.ResponseStatusCode == 502 ||
                     api.ResponseStatusCode == 503 || api.ResponseStatusCode == 504))
                    return true;
                current = current.InnerException;
            }
            return false;
        }

        private DateTime ParseStartDate()
        {
            return DateTime.TryParseExact(startDateScan, "yyyyMMdd",
                CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date)
                ? date : new DateTime(1900, 1, 1);
        }

        private static int ParsePositiveInt(string value, int fallback)
        {
            return int.TryParse(value, NumberStyles.Integer,
                       CultureInfo.InvariantCulture, out int parsed) && parsed > 0
                ? parsed : fallback;
        }

        private static List<Recipient> BuildRecipients(string value)
        {
            return (value ?? "")
                .Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(item => item.Trim())
                .Where(IsEmail)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Select(address => new Recipient
                {
                    EmailAddress = new EmailAddress { Address = address }
                }).ToList();
        }

        private static bool IsEmail(string value)
        {
            return !string.IsNullOrWhiteSpace(value) &&
                   Regex.IsMatch(value, @"^[^\s@]+@[^\s@]+\.[^\s@]+$");
        }

        private static string GetSender(Message message)
        {
            return message?.From?.EmailAddress?.Address ??
                   message?.Sender?.EmailAddress?.Address ?? "";
        }

        private static bool IsTrue(string value)
        {
            return string.Equals(value?.Trim(), "TRUE", StringComparison.OrdinalIgnoreCase);
        }

        private static void AddImmutableHeader(RequestHeaders headers)
        {
            headers.Add("Prefer", immutable_id_preference);
        }

        private static string EscapeODataString(string value)
        {
            return (value ?? "").Replace("'", "''");
        }

        private static string GetServicePath()
        {
            string location = System.Reflection.Assembly.GetEntryAssembly()?.Location;
            return string.IsNullOrWhiteSpace(location)
                ? AppDomain.CurrentDomain.BaseDirectory
                : Path.GetDirectoryName(location);
        }
    }
}
