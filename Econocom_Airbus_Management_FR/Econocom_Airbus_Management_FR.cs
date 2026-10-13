using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Kiota.Abstractions;
using Microsoft.VisualBasic.FileIO;
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
using System.Threading.Tasks;
using System.Xml;

namespace ECONOCOM_AIRBUS_MANAGEMENT_FR
{
    public class ECONOCOM_AIRBUS_MANAGEMENT_FR
    {
        private const string ApplicationName = "ECONOCOM_AIRBUS_MANAGEMENT_FR";
        private const string ImmutablePreference = "IdType=\"ImmutableId\"";
        private const string ProcessedPropertyId =
            "String {F20F6DB8-48B8-4D83-AB9F-7C7184D7E7A1} Name IMCAEconocomAirbusProcessed";

        private string country = "", skValid = "", name = "", active = "", debug = "";
        private string numberOfMails = "20", startDateScan = "", mailboxTypology = "ECONOCOM";
        private int bundleUpdateDelayDays = 3;
        private string folderIn = "Inbox", folderArchive = "archives", folderDuplicate = "doublons", folderError = "erreur";
        private string sqlGestionParameter = "", sqlGestion = "", sqlDssParameter = "", sqlDss = "";
        private string alertParameter = "", alertRecipients = "";
        private string businessAlertRecipients = "";
        private string sendAsParameter = "", sendAs = "";
        private string econocomSendAs = "fr_econocom_airbus@ingrammicro.com";
        private string imeiConfigurationParameterGlobal = "IMEI_MANAGEMENT_FR";
        private string appleImeiTableName = "dbo.IMEI_APPLE";
        private List<string> sendFilesHours = new List<string>();
        private string oracleOdsParameterGlobal = "DWHP_IMT";
        private string oracleOdsConnection = "";
        private string imXmlOrderStatusUrlParameterGlobal = "ECONOCOM_IMXML_ORDER_STATUS_URL_FR";
        private string imXmlOrderStatusTemplateParameterGlobal = "ECONOCOM_IMXML_ORDER_STATUS_TEMPLATE_FR";
        private string imXmlOrderStatusUrl = "", imXmlOrderStatusTemplate = "";
        private string imXmlOrderDetailUrlParameterGlobal = "ECONOCOM_IMXML_ORDER_DETAIL_URL_FR";
        private string imXmlOrderDetailTemplateParameterGlobal = "ECONOCOM_IMXML_ORDER_DETAIL_TEMPLATE_FR";
        private string imXmlOrderDetailUrl = "", imXmlOrderDetailTemplate = "";
        private string imXmlOrderTrackingUrlParameterGlobal = "ECONOCOM_IMXML_ORDER_TRACKING_URL_FR";
        private string imXmlOrderTrackingTemplateParameterGlobal = "ECONOCOM_IMXML_ORDER_TRACKING_TEMPLATE_FR";
        private string imXmlOrderTrackingUrl = "", imXmlOrderTrackingTemplate = "";
        private int delaiRefreshBlTracking = 180;
        private string imcaConnection = "";
        private List<ImeiSource> imeiSources = new List<ImeiSource>();
        private string logsFolder = "", tempFolder = "", sessionName = "";
        private GraphServiceClient graph;
        private static readonly object LogLock = new object();

        public sealed class JsonFile { public List<Country> countries { get; set; } = new List<Country>(); }
        public sealed class Country
        {
            public string country { get; set; } = "";
            public string sk_valid { get; set; } = "";
            public string name { get; set; } = "";
            public string active { get; set; } = "TRUE";
            public string debug { get; set; } = "FALSE";
            public string number_of_mails { get; set; } = "20";
            public string start_date_scan { get; set; } = "";
            public string sharedmailbox_typologie { get; set; } = "ECONOCOM";
            public int delai_update_bundle { get; set; } = 3;
            public string sharedmailbox_folder_in { get; set; } = "Inbox";
            public string sharedmailbox_folder_archive { get; set; } = "archives";
            public string sharedmailbox_folder_duplicate { get; set; } = "doublons";
            public string sharedmailbox_folder_error { get; set; } = "erreur";
            public string sql_connexion_gestion_cdes_parameter_global { get; set; } = "";
            public string sql_connexion_dss_parameter_global { get; set; } = "";
            public string email_in_case_of_technical_issue_parameter_global { get; set; } = "";
            public string business_alert_recipients { get; set; } = "";
            public string fr_graph_send_as_parameter_global { get; set; } = "";
            public string fr_econocom_graph_send_as { get; set; } = "fr_econocom_airbus@ingrammicro.com";
            public string imei_configuration_parameter_global { get; set; } = "IMEI_MANAGEMENT_FR";
            public string apple_imei_table_name { get; set; } = "dbo.IMEI_APPLE";
            public List<string> send_files_hours { get; set; } = new List<string>();
            public string oracle_ods_parameter_global { get; set; } = "DWHP_IMT";
            public string imxml_order_status_url_parameter_global { get; set; } = "ECONOCOM_IMXML_ORDER_STATUS_URL_FR";
            public string imxml_order_status_template_parameter_global { get; set; } = "ECONOCOM_IMXML_ORDER_STATUS_TEMPLATE_FR";
            public string imxml_order_detail_url_parameter_global { get; set; } = "ECONOCOM_IMXML_ORDER_DETAIL_URL_FR";
            public string imxml_order_detail_template_parameter_global { get; set; } = "ECONOCOM_IMXML_ORDER_DETAIL_TEMPLATE_FR";
            public string imxml_order_tracking_url_parameter_global { get; set; } = "ECONOCOM_IMXML_ORDER_TRACKING_URL_FR";
            public string imxml_order_tracking_template_parameter_global { get; set; } = "ECONOCOM_IMXML_ORDER_TRACKING_TEMPLATE_FR";
            public int Delai_Refresh_BL_Tracking { get; set; } = 180;
        }
        private sealed class ImeiJsonFile
        {
            public List<ImeiCountry> countries { get; set; } = new List<ImeiCountry>();
        }
        private sealed class ImeiCountry
        {
            public string country { get; set; } = "";
            public string active { get; set; } = "TRUE";
            public List<ImeiMailbox> sharedmailboxes { get; set; } = new List<ImeiMailbox>();
        }
        private sealed class ImeiMailbox
        {
            public string active { get; set; } = "TRUE";
            public string table_name { get; set; } = "";
            public string type_datas { get; set; } = "";
        }
        private sealed class ImeiSource
        {
            public string TableName = "";
            public string QualifiedTableName = "";
            public string LookupColumn = "";
            public string SerialColumn = "";
            public string Imei2Column = "";
            public string EidColumn = "";
            public bool IsApple;
        }
        private sealed class ImeiInformation
        {
            public string SerialNumber = "";
            public string Imei2 = "";
            public string Eid = "";
        }
        private sealed class PendingCustomerFile
        {
            public int FileId;
            public int SendType;
            public int MailId;
            public string FileName = "";
            public string CustomerNumber = "";
            public string Header = "";
            public string To = "";
            public string Cc = "";
            public string Subject = "";
            public string Body = "";
        }
        private sealed class BlStatusCandidate
        {
            public int BlId;
            public string CustomerPo = "";
            public string BlNumber = "";
            public string CustomerNumber = "";
        }
        private sealed class BlStatusLookup
        {
            public BlStatusCandidate Candidate;
            public string OrderKey = "";
            public string ShipmentKey = "";
        }
        private sealed class BlCountryCandidate
        {
            public int BlId;
            public string BlNumber = "";
            public string CustomerPo = "";
            public string CustomerNumber = "";
        }
        private sealed class ImXmlOrderDetailResult
        {
            public bool Success;
            public bool OrderNotFound;
            public string Message = "";
            public string DebugResponse = "";
            public string CountryCode = "";
            public string ZipCode = "";
        }
        private sealed class BlTrackingCandidate
        {
            public int BlId;
            public int FileId;
            public string BlNumber = "";
            public string CustomerPo = "";
            public string CustomerNumber = "";
        }
        private enum TrackingOutcome
        {
            Success,
            NotFound,
            NoPackage,
            InvalidResponse,
            ImXmlError,
            InvalidShipDate,
            TechnicalError
        }
        private sealed class ImXmlTrackingResult
        {
            public bool Success;
            public TrackingOutcome Outcome = TrackingOutcome.NotFound;
            public string Message = "";
            public string DebugResponse = "";
            public string TrackingNumber = "";
            public string CarrierCode = "";
            public DateTime? ShipDate;
        }
        private sealed class CarrierMapping
        {
            public string CarrierCode = "";
            public string CarrierName = "";
        }
        private sealed class PurchaseOrderCandidate { public int Id; public string Po = "", EndUserRef = "", Customer = ""; }
        private sealed class BlInformation { public string Number = "", Status = ""; }
        private sealed class ImXmlCredentials { public string Login = "", Password = ""; }
        private sealed class ImXmlOrderStatusResult { public bool Success; public string Message = "", DebugResponse = ""; public List<BlInformation> BlList = new List<BlInformation>(); }
        private enum BlUpsertResult { Inserted, Updated, Unchanged }
        private sealed class SharedMailbox
        {
            public int Id; public string Name = "", Address = ""; public DateTime? FilterDate, LastRefresh;
            public int RefreshMinutes; public bool Active;
        }
        private sealed class ProcessingState
        {
            public string Status = "", ErrorType = "", DestinationFolder = "";
            public int? MailId;
            public DateTime? AlertSentAtUtc;
            public DateTime? FirstSeenAtUtc;
        }
        private sealed class ClientImportConfiguration
        {
            public string CustomerNumber = "";
            public string ExpectedCsvHeader = "";
        }

        private sealed class CsvFileValidation
        {
            public string Name = "";
            public string TempPath = "";
            public List<string[]> Rows = new List<string[]>();
            public Dictionary<string, int> ColumnIndexes =
                new Dictionary<string, int>(
                    StringComparer.OrdinalIgnoreCase);
        }
        private sealed class CsvValidationResult
        {
            public bool IsValid, IsDuplicate; public string ErrorType = "", Message = "";
            public List<string> MissingBundles = new List<string>();
            public List<string> UnexpectedColumns = new List<string>();
            public List<CsvFileValidation> Files = new List<CsvFileValidation>();
        }

        private static readonly Dictionary<string, string[]> CsvColumnAliases =
            new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
            {
                { "order_number", new[] { "Order_Number" } },
                { "order_reference", new[] { "Order_Reference" } },
                { "product_type", new[] { "Product_Type", "Bundle", "Product_Bundle" } },
                { "product_code", new[] { "Product_Code", "Supplier_Product_Ref" } },
                { "end_user_first_name", new[] { "EndUser_FirstName", "EndUser_First_Name" } },
                { "end_user_last_name", new[] { "EndUser_LastName", "EndUser_Last_Name", "EndUser_First_Last_Name" } },
                { "end_user_email", new[] { "EndUser_eMail_Address", "EndUser_Email_Address" } },
                { "end_user_id", new[] { "UserID", "EndUser_ID" } },
                { "label_content", new[] { "Label_Content" } },
                { "comment", new[] { "Comment" } }
            };

        public void Read_Email_with_Graph(string sql_con, string logs, string tmp_folder, string session_name)
        {
            imcaConnection = sql_con;
            string root = GetServicePath();
            logsFolder = Path.Combine(root, logs ?? ""); tempFolder = Path.Combine(root, tmp_folder ?? ""); sessionName = session_name ?? "";
            Directory.CreateDirectory(logsFolder); Directory.CreateDirectory(tempFolder);
            try
            {
                JsonFile cfg = JsonConvert.DeserializeObject<JsonFile>(GetParameter(sql_con, ApplicationName) ?? "");
                if (cfg?.countries == null || cfg.countries.Count == 0) throw new InvalidOperationException(ApplicationName + " parameters are empty or invalid");
                foreach (Country item in cfg.countries)
                {
                    Apply(item, sql_con); if (!IsTrue(active)) continue; Validate();
                    WriteLog(name.ToUpperInvariant() + "(" + country.ToUpperInvariant() + ") at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                    WriteLog("   Debug Parameter is set to " + debug.ToUpperInvariant());
                    WriteLog("   Connecting to Microsoft Graph"); graph = ConnectGraph(); WriteLog("   Microsoft Graph connection established");
                    foreach (SharedMailbox mailbox in LoadMailboxes())
                    {
                        WriteLog("   Mailbox loaded from T_SharedMailboxes - ID " + mailbox.Id + " - " + mailbox.Name + " - " + mailbox.Address);
                        if (!mailbox.Active) { WriteLog("   Mailbox skipped because actif is not 1 : " + mailbox.Address); continue; }
                        if (!IsDue(mailbox)) { WriteLog("   Mailbox skipped, refresh is not due : " + mailbox.Address); continue; }
                        try
                        {
                            ProcessMailbox(mailbox, mailbox.FilterDate ?? ParseStartDate());
                            ExecuteNonQuery(sqlGestion, "UPDATE dbo.T_SharedMailboxes SET date_dernier_raf=GETDATE() WHERE id_mailboxe=@ID;",
                                new SqlParameter("@ID", SqlDbType.Int) { Value = mailbox.Id });
                        }
                        catch (Exception ex)
                        {
                            WriteLog("   Mailbox error - " + mailbox.Address + " - " + Details(ex));
                            SendTechnicalAlert("Read_Email_with_Graph", Details(ex));
                        }
                    }
                }
            }
            finally { graph = null; }
        }

        public void Update_BL_Status_ODS(string sql_con, string logs, string tmp_folder, string session_name)
        {
            imcaConnection = sql_con; string root = GetServicePath();
            logsFolder = Path.Combine(root, logs ?? ""); tempFolder = Path.Combine(root, tmp_folder ?? ""); sessionName = session_name ?? "";
            Directory.CreateDirectory(logsFolder); Directory.CreateDirectory(tempFolder);
            JsonFile cfg = JsonConvert.DeserializeObject<JsonFile>(GetParameter(sql_con, ApplicationName) ?? "");
            if (cfg?.countries == null || cfg.countries.Count == 0) throw new InvalidOperationException(ApplicationName + " parameters are empty or invalid");
            foreach (Country item in cfg.countries)
            {
                Apply(item, sql_con); if (!IsTrue(item.active)) continue; ValidateUpdateBlOdsConfiguration();
                try
                {
                    List<BlStatusCandidate> candidates = LoadBlStatusCandidates();
                    int foundOds = 0, voided = 0, foundDss = 0, invoiced = 0, unchanged = 0, invalid = 0;
                    WriteLog("   BL status candidate(s) : " + candidates.Count + " - Refresh delay : " + delaiRefreshBlTracking + " day(s)");
                    List<BlStatusLookup> validCandidates =
                        BuildValidBlStatusLookups(candidates, ref invalid);

                    Dictionary<string, string> odsStatuses;
                    using (var oracle = new OracleConnection(oracleOdsConnection))
                    {
                        oracle.Open();
                        WriteLog("   Direct Oracle ODS connection established");
                        odsStatuses = LoadBlStatusesFromOracle(
                            oracle,
                            validCandidates);
                    }

                    List<BlStatusLookup> dssCandidates =
                        validCandidates
                            .Where(candidate =>
                                !odsStatuses.TryGetValue(
                                    BuildBlStatusKey(
                                        candidate.OrderKey,
                                        candidate.ShipmentKey),
                                    out string status) ||
                                string.IsNullOrWhiteSpace(status))
                            .ToList();

                    Dictionary<string, string> dssInvoices =
                        LoadLatestInvoicesFromDss(dssCandidates);

                    WriteLog(
                        "   Grouped status lookup completed" +
                        " - Valid BL : " + validCandidates.Count +
                        " - ODS result(s) : " +
                        odsStatuses.Count(pair =>
                            !string.IsNullOrWhiteSpace(pair.Value)) +
                        " - DSS fallback candidate(s) : " +
                        dssCandidates.Count +
                        " - DSS invoice(s) : " + dssInvoices.Count);

                    foreach (BlStatusLookup lookup in validCandidates)
                    {
                        BlStatusCandidate candidate = lookup.Candidate;
                        string odsKey = BuildBlStatusKey(
                            lookup.OrderKey,
                            lookup.ShipmentKey);

                        odsStatuses.TryGetValue(
                            odsKey,
                            out string status);

                        status = status ?? "";

                        if (!string.IsNullOrWhiteSpace(status))
                        {
                            foundOds++;

                            if (status.Equals(
                                    "VOIDED",
                                    StringComparison.OrdinalIgnoreCase))
                            {
                                UpdateBlStatus(candidate, "VOIDED");
                                voided++;
                            }
                            else
                            {
                                unchanged++;
                            }

                            continue;
                        }

                        dssInvoices.TryGetValue(
                            BuildDssInvoiceKey(
                                candidate.CustomerNumber,
                                lookup.OrderKey),
                            out string invoice);

                        invoice = invoice ?? "";

                        if (!string.IsNullOrWhiteSpace(invoice))
                        {
                            UpdateBlStatus(candidate, "INVOICED");
                            foundDss++;
                            invoiced++;

                            WriteLog(
                                "       BL updated from grouped DSS fallback" +
                                " - BL : " + candidate.BlNumber +
                                " - Invoice : " + invoice +
                                " - Status : INVOICED");
                        }
                        else
                        {
                            unchanged++;
                        }
                    }
                    WriteLog("   Update_BL_Status_ODS summary - Candidates : " + candidates.Count + " - Found ODS : " + foundOds + " - VOIDED : " + voided + " - Found DSS : " + foundDss + " - INVOICED : " + invoiced + " - Unchanged : " + unchanged + " - Invalid : " + invalid);
                }
                catch (Exception ex) { WriteLog("   Update_BL_Status_ODS error - " + Details(ex)); SendTechnicalAlert(nameof(Update_BL_Status_ODS), Details(ex)); }
            }
        }
        private void ValidateUpdateBlOdsConfiguration()
        { var m = new List<string>(); if (string.IsNullOrWhiteSpace(sqlGestion)) m.Add("sqlGestion"); if (string.IsNullOrWhiteSpace(sqlDss)) m.Add("sqlDss"); if (string.IsNullOrWhiteSpace(oracleOdsConnection)) m.Add("oracleOdsConnection"); if (m.Count > 0) throw new InvalidOperationException("Update_BL_Status_ODS configuration incomplete : " + string.Join(", ", m)); }
        private List<BlStatusCandidate> LoadBlStatusCandidates()
        {
            const string sql = @"SELECT DISTINCT bl.id_bl,ISNULL(cp.cust_po,'') cust_po,ISNULL(bl.num_bl,'') num_bl,ISNULL(m.cust_nbr,'') cust_nbr
FROM dbo.T_suivi_commande_end_user_cust_po cp JOIN dbo.T_suivi_commande_end_user_fichier f ON cp.id_fichier=f.id_fichier
JOIN dbo.T_suivi_commande_end_user_mail m ON f.id_mail=m.id_mail JOIN dbo.T_suivi_commande_end_user_BL bl ON cp.cust_po=bl.cust_po
WHERE DATEDIFF(DAY,m.date_mail,GETDATE())<=@DELAI AND cp.ouvert=1 AND cp.cde_annule<>1 AND cp.cust_po IS NOT NULL AND bl.num_bl IS NOT NULL
AND (bl.status_bl IS NULL OR bl.status_bl NOT IN ('INVOICED','FORCED BILLED','BILLING ERROR','VOIDED')) ORDER BY bl.id_bl;";
            var rlt = new List<BlStatusCandidate>(); using (var cn = new SqlConnection(sqlGestion)) using (var cmd = new SqlCommand(sql, cn)) { cmd.CommandTimeout = 300; cmd.Parameters.Add("@DELAI", SqlDbType.Int).Value = delaiRefreshBlTracking; cn.Open(); using (var r = cmd.ExecuteReader()) while (r.Read()) rlt.Add(new BlStatusCandidate { BlId = Convert.ToInt32(r["id_bl"]), CustomerPo = Convert.ToString(r["cust_po"]).Trim(), BlNumber = Convert.ToString(r["num_bl"]).Trim(), CustomerNumber = Convert.ToString(r["cust_nbr"]).Trim() }); }
            return rlt;
        }
        private static bool TrySplitBlNumber(string blNumber, out string orderKey, out string shipmentKey)
        {
            orderKey = ""; shipmentKey = ""; string normalized = (blNumber ?? "").Trim().Replace(" ", ""); if (string.IsNullOrWhiteSpace(normalized)) return false;
            string[] parts = normalized.Split(new[] { '-' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length >= 2) { orderKey = parts[0].Trim(); shipmentKey = parts[parts.Length - 1].Trim(); }
            else { if (normalized.Length < 9) return false; orderKey = normalized.Substring(0, normalized.Length - 2); shipmentKey = normalized.Substring(normalized.Length - 2, 2); }
            return Regex.IsMatch(orderKey, @"^[A-Za-z0-9]+$") && Regex.IsMatch(shipmentKey, @"^[0-9]{2}$");
        }
        private List<BlStatusLookup> BuildValidBlStatusLookups(
            List<BlStatusCandidate> candidates,
            ref int invalid)
        {
            var result = new List<BlStatusLookup>();

            foreach (BlStatusCandidate candidate in
                candidates ?? new List<BlStatusCandidate>())
            {
                if (!TrySplitBlNumber(
                        candidate.BlNumber,
                        out string orderKey,
                        out string shipmentKey))
                {
                    invalid++;
                    WriteLog(
                        "       Invalid BL number" +
                        " - ID : " + candidate.BlId +
                        " - BL : " + candidate.BlNumber);
                    continue;
                }

                result.Add(new BlStatusLookup
                {
                    Candidate = candidate,
                    OrderKey = orderKey,
                    ShipmentKey = shipmentKey
                });
            }

            return result;
        }

        private static string BuildBlStatusKey(
            string orderKey,
            string shipmentKey)
        {
            return (orderKey ?? "").Trim().ToUpperInvariant() +
                   "|" +
                   (shipmentKey ?? "").Trim().ToUpperInvariant();
        }

        private static string BuildDssInvoiceKey(
            string customerNumber,
            string orderKey)
        {
            return (customerNumber ?? "").Trim().ToUpperInvariant() +
                   "|" +
                   (orderKey ?? "").Trim().ToUpperInvariant();
        }

        private static Dictionary<string, string>
            LoadBlStatusesFromOracle(
                OracleConnection connection,
                List<BlStatusLookup> candidates)
        {
            var result = new Dictionary<string, string>(
                StringComparer.OrdinalIgnoreCase);

            List<BlStatusLookup> uniqueCandidates =
                (candidates ?? new List<BlStatusLookup>())
                    .GroupBy(
                        candidate => BuildBlStatusKey(
                            candidate.OrderKey,
                            candidate.ShipmentKey),
                        StringComparer.OrdinalIgnoreCase)
                    .Select(group => group.First())
                    .ToList();

            const int batchSize = 100;

            for (int offset = 0;
                 offset < uniqueCandidates.Count;
                 offset += batchSize)
            {
                List<BlStatusLookup> batch =
                    uniqueCandidates
                        .Skip(offset)
                        .Take(batchSize)
                        .ToList();

                var requestedRows = new List<string>();

                using (var command = connection.CreateCommand())
                {
                    command.BindByName = true;
                    command.CommandTimeout = 300;

                    command.Parameters.Add(
                        "BILL_CUST_NBR",
                        OracleDbType.Varchar2,
                        "203530",
                        ParameterDirection.Input);
                    command.Parameters.Add(
                        "BR_NBR",
                        OracleDbType.Varchar2,
                        "21",
                        ParameterDirection.Input);

                    for (int index = 0;
                         index < batch.Count;
                         index++)
                    {
                        string orderParameter =
                            "ORDER_KEY_" + index;
                        string shipmentParameter =
                            "SHIPMENT_KEY_" + index;

                        requestedRows.Add(
                            "SELECT :" + orderParameter +
                            " AS ORDER_KEY, :" + shipmentParameter +
                            " AS SHIPMENT_KEY FROM DUAL");

                        command.Parameters.Add(
                            orderParameter,
                            OracleDbType.Varchar2,
                            batch[index].OrderKey,
                            ParameterDirection.Input);
                        command.Parameters.Add(
                            shipmentParameter,
                            OracleDbType.Varchar2,
                            batch[index].ShipmentKey,
                            ParameterDirection.Input);
                    }

                    command.CommandText = @"
WITH REQUESTED_BL AS
(
    " + string.Join(
        Environment.NewLine + " UNION ALL " +
        Environment.NewLine,
        requestedRows) + @"
),
ODS_STATUS_RANKED AS
(
    SELECT TO_CHAR(OS.BR_NBR) ||
           TO_CHAR(OS.ORDR_NBR) AS ORDER_KEY,
           TO_CHAR(OS.DIST_NBR) ||
           TO_CHAR(OS.SHIP_NBR) AS SHIPMENT_KEY,
           OS.SHIP_STUS,
           ROW_NUMBER() OVER
           (
               PARTITION BY TO_CHAR(OS.BR_NBR) || TO_CHAR(OS.ORDR_NBR),
                            TO_CHAR(OS.DIST_NBR) || TO_CHAR(OS.SHIP_NBR)
               ORDER BY OH.ORDR_DATE DESC
           ) AS ROW_NUMBER_VALUE
    FROM DSSDATA.UVW_ORDERHEADFR OH
    INNER JOIN DSSDATA.UVW_ORDERSHIPFR OS
        ON OH.BR_NBR = OS.BR_NBR
       AND OH.ORDR_NBR = OS.ORDR_NBR
       AND OH.ORDR_DATE = OS.ORDR_DT
    INNER JOIN REQUESTED_BL REQUESTED
        ON REQUESTED.ORDER_KEY =
           TO_CHAR(OS.BR_NBR) || TO_CHAR(OS.ORDR_NBR)
       AND REQUESTED.SHIPMENT_KEY =
           TO_CHAR(OS.DIST_NBR) || TO_CHAR(OS.SHIP_NBR)
    WHERE TO_CHAR(OH.BILL_CUST_NBR) = :BILL_CUST_NBR
      AND TO_CHAR(OS.BR_NBR) = :BR_NBR
),
ODS_STATUS AS
(
    SELECT ORDER_KEY,
           SHIPMENT_KEY,
           SHIP_STUS
    FROM ODS_STATUS_RANKED
    WHERE ROW_NUMBER_VALUE = 1
)
SELECT REQUESTED.ORDER_KEY,
       REQUESTED.SHIPMENT_KEY,
       ODS_STATUS.SHIP_STUS
FROM REQUESTED_BL REQUESTED
LEFT JOIN ODS_STATUS
    ON ODS_STATUS.ORDER_KEY = REQUESTED.ORDER_KEY
   AND ODS_STATUS.SHIPMENT_KEY = REQUESTED.SHIPMENT_KEY";

                    using (OracleDataReader reader =
                        command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string orderKey =
                                Convert.ToString(
                                    reader["ORDER_KEY"]).Trim();
                            string shipmentKey =
                                Convert.ToString(
                                    reader["SHIPMENT_KEY"]).Trim();
                            string status =
                                reader["SHIP_STUS"] == DBNull.Value
                                    ? ""
                                    : Convert.ToString(
                                        reader["SHIP_STUS"]).Trim();

                            result[BuildBlStatusKey(
                                orderKey,
                                shipmentKey)] = status;
                        }
                    }
                }
            }

            return result;
        }

        private Dictionary<string, string>
            LoadLatestInvoicesFromDss(
                List<BlStatusLookup> candidates)
        {
            var result = new Dictionary<string, string>(
                StringComparer.OrdinalIgnoreCase);

            List<BlStatusLookup> uniqueCandidates =
                (candidates ?? new List<BlStatusLookup>())
                    .GroupBy(
                        candidate => BuildDssInvoiceKey(
                            candidate.Candidate.CustomerNumber,
                            candidate.OrderKey),
                        StringComparer.OrdinalIgnoreCase)
                    .Select(group => group.First())
                    .ToList();

            const int batchSize = 500;

            for (int offset = 0;
                 offset < uniqueCandidates.Count;
                 offset += batchSize)
            {
                List<BlStatusLookup> batch =
                    uniqueCandidates
                        .Skip(offset)
                        .Take(batchSize)
                        .ToList();

                var requestedRows = new List<string>();

                using (var connection = new SqlConnection(sqlDss))
                using (var command = connection.CreateCommand())
                {
                    command.CommandTimeout = 300;

                    for (int index = 0;
                         index < batch.Count;
                         index++)
                    {
                        string customerParameter =
                            "@CUSTOMER_" + index;
                        string orderParameter =
                            "@ORDER_" + index;

                        requestedRows.Add(
                            "SELECT " + customerParameter +
                            " AS CUSTOMER_NUMBER, " + orderParameter +
                            " AS ORDER_KEY");

                        command.Parameters.Add(P(
                            customerParameter,
                            SqlDbType.VarChar,
                            batch[index].Candidate.CustomerNumber,
                            50));
                        command.Parameters.Add(P(
                            orderParameter,
                            SqlDbType.VarChar,
                            batch[index].OrderKey,
                            50));
                    }

                    command.CommandText = @"
WITH REQUESTED AS
(
    " + string.Join(
        Environment.NewLine + " UNION ALL " +
        Environment.NewLine,
        requestedRows) + @"
),
LATEST_INVOICE AS
(
    SELECT REQUESTED.CUSTOMER_NUMBER,
           REQUESTED.ORDER_KEY,
           ISNULL(HEADER.invoice_nbr, '') AS invoice_nbr,
           ROW_NUMBER() OVER
           (
               PARTITION BY REQUESTED.CUSTOMER_NUMBER,
                            REQUESTED.ORDER_KEY
               ORDER BY HEADER.invoice_date DESC,
                        HEADER.invoice_nbr DESC
           ) AS row_number
    FROM REQUESTED
    LEFT JOIN dbo.shipment_header HEADER
        ON HEADER.order_nbr = REQUESTED.ORDER_KEY
       AND HEADER.branch_customer_nbr =
           REQUESTED.CUSTOMER_NUMBER
       AND HEADER.invoice_date > DATEADD(year, -2, GETDATE())
)
SELECT CUSTOMER_NUMBER,
       ORDER_KEY,
       invoice_nbr
FROM LATEST_INVOICE
WHERE row_number = 1;";

                    connection.Open();

                    using (SqlDataReader reader =
                        command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string customerNumber =
                                Convert.ToString(
                                    reader["CUSTOMER_NUMBER"]).Trim();
                            string orderKey =
                                Convert.ToString(
                                    reader["ORDER_KEY"]).Trim();
                            string invoice =
                                Convert.ToString(
                                    reader["invoice_nbr"]).Trim();

                            if (!string.IsNullOrWhiteSpace(invoice))
                            {
                                result[BuildDssInvoiceKey(
                                    customerNumber,
                                    orderKey)] = invoice;
                            }
                        }
                    }
                }
            }

            return result;
        }
        private void UpdateBlStatus(BlStatusCandidate c, string status)
        { ExecuteNonQuery(sqlGestion, "UPDATE dbo.T_suivi_commande_end_user_BL SET status_bl=@S WHERE id_bl=@ID AND cust_po=@PO AND num_bl=@BL;", P("@S", SqlDbType.VarChar, status, 100), new SqlParameter("@ID", SqlDbType.Int) { Value = c.BlId }, P("@PO", SqlDbType.VarChar, c.CustomerPo, 500), P("@BL", SqlDbType.VarChar, c.BlNumber, 500)); }

        public void Search_BL_Numbers(string sql_con, string logs, string tmp_folder, string session_name)
        {
            imcaConnection = sql_con; string root = GetServicePath(); logsFolder = Path.Combine(root, logs ?? ""); tempFolder = Path.Combine(root, tmp_folder ?? ""); sessionName = session_name ?? "";
            Directory.CreateDirectory(logsFolder); Directory.CreateDirectory(tempFolder);
            JsonFile cfg = JsonConvert.DeserializeObject<JsonFile>(GetParameter(sql_con, ApplicationName) ?? "");
            if (cfg?.countries == null || cfg.countries.Count == 0) throw new InvalidOperationException(ApplicationName + " parameters are empty or invalid");
            foreach (Country item in cfg.countries)
            {
                Apply(item, sql_con); if (!IsTrue(item.active)) continue; ValidateSearchBlConfiguration(); OracleConnection oracle = null;
                try
                {
                    List<PurchaseOrderCandidate> candidates = LoadPurchaseOrdersForBlSearch();
                    int poFound = 0, poMissing = 0, ok = 0, errors = 0, noBl = 0, inserted = 0, updated = 0, unchanged = 0, invalidClosed = 0;
                    WriteLog("   Search BL candidate(s) : " + candidates.Count + " - Refresh delay : " + delaiRefreshBlTracking + " day(s)");
                    foreach (var c in candidates)
                    {
                        try
                        {
                            string po = c.Po;
                            if (string.IsNullOrWhiteSpace(po) && c.Customer == "21203530")
                            {
                                if (oracle == null) { oracle = new OracleConnection(oracleOdsConnection); oracle.Open(); WriteLog("   Direct Oracle ODS connection established for Customer PO lookup"); }
                                po = FindCustomerPoFromOracle(oracle, c.EndUserRef);
                                if (string.IsNullOrWhiteSpace(po)) { poMissing++; continue; }
                                ExecuteNonQuery(sqlGestion, "UPDATE dbo.T_suivi_commande_end_user_cust_po SET cust_po=@P WHERE id_cust_po=@I;", P("@P", SqlDbType.VarChar, po, 500), new SqlParameter("@I", SqlDbType.Int) { Value = c.Id }); poFound++;
                            }
                            if (string.IsNullOrWhiteSpace(po)) { poMissing++; continue; }
                            int requestId = InsertImXmlRequest(c.Customer, "Order Status : recherche num BL", po);
                            ImXmlOrderStatusResult r;
                            try { r = ExecuteImXmlOrderStatus(po, c.Customer, requestId); } catch (Exception ex) { r = new ImXmlOrderStatusResult { Message = "IMXML technical error : " + Details(ex) }; }
                            UpdateImXmlRequest(requestId, r.Message);
                            if (!r.Success) { errors++; WriteLog("   IMXML Order Status error - Customer PO : " + po + " - Message : " + r.Message); if (IsTrue(debug) && !string.IsNullOrWhiteSpace(r.DebugResponse)) WriteLog("   IMXML response - Customer PO : " + po + " - Content : " + r.DebugResponse); if ((r.Message ?? "").IndexOf("Customer Purchase Order is invalid", StringComparison.OrdinalIgnoreCase) >= 0) { CloseCustomerPo(c.Id); invalidClosed++; } continue; }
                            ok++; if (r.BlList.Count == 0) { noBl++; continue; }
                            foreach (var bl in r.BlList) { var u = UpsertBlFromImXml(po, c.Customer, bl); if (u == BlUpsertResult.Inserted) inserted++; else if (u == BlUpsertResult.Updated) updated++; else unchanged++; }
                        }
                        catch (Exception ex) { errors++; WriteLog("   Search BL error - Customer PO ID : " + c.Id + " - " + Details(ex)); }
                    }
                    WriteLog(
                        "   Search_BL_Numbers summary" +
                        " - Candidates : " + candidates.Count +
                        " - Customer PO found in ODS : " + poFound +
                        " - Customer PO not found : " + poMissing +
                        " - IMXML success : " + ok +
                        " - IMXML errors : " + errors +
                        " - No BL : " + noBl);
                    WriteLog(
                        "   Search_BL_Numbers BL summary" +
                        " - Inserted : " + inserted +
                        " - Updated : " + updated +
                        " - Unchanged : " + unchanged +
                        " - Invalid PO closed : " + invalidClosed);
                }
                catch (Exception ex) { WriteLog("   Search_BL_Numbers error - " + Details(ex)); SendTechnicalAlert(nameof(Search_BL_Numbers), Details(ex)); }
                finally { oracle?.Dispose(); }
            }
        }
        private void ValidateSearchBlConfiguration() { var m = new List<string>(); if (string.IsNullOrWhiteSpace(sqlGestion)) m.Add("sqlGestion"); if (string.IsNullOrWhiteSpace(oracleOdsConnection)) m.Add("oracleOdsConnection"); if (string.IsNullOrWhiteSpace(imXmlOrderStatusUrl)) m.Add("imXmlOrderStatusUrl"); if (string.IsNullOrWhiteSpace(imXmlOrderStatusTemplate)) m.Add("imXmlOrderStatusTemplate"); if (m.Count > 0) throw new InvalidOperationException("Search_BL_Numbers configuration incomplete : " + string.Join(", ", m)); imXmlOrderStatusUrl = ValidateHttpUrl(imXmlOrderStatusUrl, imXmlOrderStatusUrlParameterGlobal); var d = new XmlDocument(); d.LoadXml(imXmlOrderStatusTemplate); string[] required = { "//*[local-name()='OrderStatusRequest']/*[local-name()='Version']", "//*[local-name()='TransactionHeader']/*[local-name()='SenderID']", "//*[local-name()='TransactionHeader']/*[local-name()='ReceiverID']", "//*[local-name()='TransactionHeader']/*[local-name()='CountryCode']", "//*[local-name()='TransactionHeader']/*[local-name()='LoginID']", "//*[local-name()='TransactionHeader']/*[local-name()='Password']", "//*[local-name()='TransactionHeader']/*[local-name()='TransactionID']", "//*[local-name()='OrderHeaderInfo']/*[local-name()='CustomerPO']" }; foreach (string xpath in required) if (d.SelectSingleNode(xpath) == null) throw new InvalidOperationException("Required node missing from IMXML template : " + xpath); }
        private List<PurchaseOrderCandidate> LoadPurchaseOrdersForBlSearch()
        {
            const string q = @"
SELECT DISTINCT
       cp.id_cust_po,
       ISNULL(cp.cust_po, '') AS cust_po,
       ISNULL(cp.end_user_order_ref, '') AS end_user_order_ref,
       ISNULL(m.cust_nbr, '') AS cust_nbr
FROM dbo.T_suivi_commande_end_user_cust_po cp
INNER JOIN dbo.T_suivi_commande_end_user_fichier f
    ON f.id_fichier = cp.id_fichier
INNER JOIN dbo.T_suivi_commande_end_user_mail m
    ON m.id_mail = f.id_mail
WHERE DATEDIFF(DAY, m.date_mail, GETDATE()) <= @D
  AND cp.ouvert = 1
  AND ISNULL(cp.cde_annule, 0) <> 1
  AND
  (
      NOT EXISTS
      (
          SELECT 1
          FROM dbo.T_suivi_commande_end_user_BL bl
          WHERE bl.cust_po = cp.cust_po
      )
      OR EXISTS
      (
          SELECT 1
          FROM dbo.T_suivi_commande_end_user_BL bl
          WHERE bl.cust_po = cp.cust_po
            AND
            (
                bl.status_bl IS NULL
                OR bl.status_bl NOT IN
                (
                    'INVOICED',
                    'FORCED BILLED',
                    'BILLING ERROR',
                    'VOIDED'
                )
            )
      )
      OR NOT EXISTS
      (
          SELECT 1
          FROM dbo.T_suivi_commande_end_user_BL bl_billed
          WHERE bl_billed.cust_po = cp.cust_po
            AND bl_billed.status_bl IN
            (
                'INVOICED',
                'FORCED BILLED'
            )
      )
  )
ORDER BY cp.id_cust_po;";

            var candidates = new List<PurchaseOrderCandidate>();

            using (var connection = new SqlConnection(sqlGestion))
            using (var command = new SqlCommand(q, connection))
            {
                command.CommandTimeout = 300;
                command.Parameters.Add("@D", SqlDbType.Int).Value = delaiRefreshBlTracking;
                connection.Open();

                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        candidates.Add(new PurchaseOrderCandidate
                        {
                            Id = Convert.ToInt32(reader["id_cust_po"]),
                            Po = Convert.ToString(reader["cust_po"]).Trim(),
                            EndUserRef = Convert.ToString(reader["end_user_order_ref"]).Trim(),
                            Customer = Convert.ToString(reader["cust_nbr"]).Trim()
                        });
                    }
                }
            }

            return candidates;
        }
        private static string FindCustomerPoFromOracle(OracleConnection cn, string e) { if (string.IsNullOrWhiteSpace(e)) return ""; const string q = @"SELECT CUST_ORDR_NBR FROM (SELECT TRIM(CUST_ORDR_NBR) CUST_ORDR_NBR FROM DSSDATA.UVW_ORDERHEADFR WHERE TRIM(EU_PO_NBR)=:E AND TO_CHAR(BILL_CUST_NBR)=:C AND CUST_ORDR_NBR IS NOT NULL ORDER BY CUST_ORDR_NBR DESC) WHERE ROWNUM=1"; using (var cmd = new OracleCommand(q, cn)) { cmd.BindByName = true; cmd.Parameters.Add("E", OracleDbType.Varchar2, e, ParameterDirection.Input); cmd.Parameters.Add("C", OracleDbType.Varchar2, "203530", ParameterDirection.Input); object v = cmd.ExecuteScalar(); return v == null || v == DBNull.Value ? "" : Convert.ToString(v).Trim(); } }
        private ImXmlCredentials LoadImXmlCredentials(string c) { using (var cn = new SqlConnection(sqlGestion)) using (var cmd = new SqlCommand("SELECT TOP(1) ISNULL(imXml_login_user,''),ISNULL(imXml_mdp,'') FROM dbo.T_suivi_commande_end_user_client_config WHERE cust_nbr=@C", cn)) { cmd.Parameters.Add(P("@C", SqlDbType.VarChar, c, 100)); cn.Open(); using (var r = cmd.ExecuteReader()) { if (!r.Read()) throw new InvalidOperationException("No IMXML configuration found for customer " + c); return new ImXmlCredentials { Login = Convert.ToString(r[0]).Trim(), Password = Convert.ToString(r[1]).Trim() }; } } }
        private ImXmlOrderStatusResult ExecuteImXmlOrderStatus(string po, string customer, int requestId) { var request = BuildImXmlOrderStatusRequest(po, requestId, LoadImXmlCredentials(customer)); if (IsTrue(debug)) WriteLog("   IMXML Order Status request - Customer PO : " + po); return ParseImXmlOrderStatusResponse(PostImXmlRequest(request, imXmlOrderStatusUrl)); }
        private XmlDocument BuildImXmlOrderStatusRequest(string po, int id, ImXmlCredentials c) { var d = new XmlDocument(); d.LoadXml(imXmlOrderStatusTemplate); SetNode(d, "//*[local-name()='LoginID']", c.Login); SetNode(d, "//*[local-name()='Password']", c.Password); SetNode(d, "//*[local-name()='TransactionID']", id.ToString(CultureInfo.InvariantCulture)); SetNode(d, "//*[local-name()='CustomerPO']", po); return d; }
        private static void SetNode(XmlDocument d, string x, string v) { XmlNode n = d.SelectSingleNode(x); if (n == null) throw new InvalidOperationException("Required IMXML template node not found : " + x); n.InnerText = v ?? ""; }
        private string PostImXmlRequest(XmlDocument d, string requestUrl) { Encoding encoding = Encoding.GetEncoding("iso-8859-1"); byte[] b = encoding.GetBytes(d.OuterXml); ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12; var h = WebRequest.CreateHttp(requestUrl); h.Method = "POST"; h.ContentType = "text/xml; charset=iso-8859-1"; h.ContentLength = b.Length; h.Timeout = 300000; h.ReadWriteTimeout = 300000; using (var o = h.GetRequestStream()) o.Write(b, 0, b.Length); using (var r = (HttpWebResponse)h.GetResponse()) using (var i = r.GetResponseStream()) using (var sr = new StreamReader(i ?? Stream.Null, encoding, true)) return sr.ReadToEnd(); }
        private static ImXmlOrderStatusResult ParseImXmlOrderStatusResponse(string text) { var z = new ImXmlOrderStatusResult(); if (string.IsNullOrWhiteSpace(text)) { z.Message = "Empty IMXML response"; return z; } var d = new XmlDocument(); try { d.LoadXml(text); } catch (Exception ex) { z.Message = "Invalid IMXML response : " + Truncate(ex.Message, 150); z.DebugResponse = Truncate(Regex.Replace(text ?? "", @"\s+", " ").Trim(), 500); return z; } XmlNode root = d.SelectSingleNode("//*[local-name()='OrderStatusResponse']"); if (root == null) { string rootName = d.DocumentElement == null ? "none" : d.DocumentElement.LocalName; z.Message = "Unexpected IMXML response - Root : " + rootName; z.DebugResponse = Truncate(Regex.Replace(d.OuterXml ?? "", @"\s+", " ").Trim(), 500); return z; } string er = (root.SelectSingleNode(".//*[local-name()='ErrorStatus']")?.InnerText ?? "").Trim(); if (er != "") { z.Message = Truncate(er, 250); return z; } foreach (XmlNode o in root.SelectNodes(".//*[local-name()='OrderInfo']")) { string n = (o.SelectSingleNode("./*[local-name()='BranchOrderNumber']")?.InnerText ?? "").Trim(); foreach (XmlNode x in o.SelectNodes("./*[local-name()='OrderSuffix']")) { if (!string.IsNullOrWhiteSpace(x.SelectSingleNode("./*[local-name()='InvoiceDate']")?.InnerText)) continue; string id = x.Attributes?["ID"]?.Value ?? ""; if (id != "") z.BlList.Add(new BlInformation { Number = n + "-" + id, Status = (x.SelectSingleNode("./*[local-name()='OrderStatus']")?.InnerText ?? "").Trim() }); } } z.Success = true; z.Message = z.BlList.Count == 0 ? "Requête OK - aucun BL trouvé" : "Requête OK - " + z.BlList.Count + " BL trouvé(s)"; return z; }
        private int InsertImXmlRequest(string c, string t, string p) => ScalarInt(sqlGestion, "INSERT INTO dbo.T_suivi_commande_end_user_imXml_request(cust_nbr,request_type,request_date,cust_po) VALUES(@C,@T,GETDATE(),@P);SELECT CONVERT(int,SCOPE_IDENTITY());", P("@C", SqlDbType.VarChar, c, 100), P("@T", SqlDbType.VarChar, t, 500), P("@P", SqlDbType.VarChar, p, 500));
        private void UpdateImXmlRequest(int id, string m) => ExecuteNonQuery(sqlGestion, "UPDATE dbo.T_suivi_commande_end_user_imXml_request SET cmt=@M WHERE id_request=@I", P("@M", SqlDbType.VarChar, Truncate(m, 250), 250), new SqlParameter("@I", SqlDbType.Int) { Value = id });
        private static bool IsFinalBlStatus(string s) => new[] { "INVOICED", "FORCED BILLED", "BILLING ERROR", "VOIDED" }.Contains((s ?? "").Trim(), StringComparer.OrdinalIgnoreCase);
        private BlUpsertResult UpsertBlFromImXml(string po, string customer, BlInformation bl) { using (var cn = new SqlConnection(sqlGestion)) { cn.Open(); using (var tx = cn.BeginTransaction(IsolationLevel.Serializable)) { int id = 0; string current = ""; using (var cmd = new SqlCommand("SELECT TOP(1) id_bl,ISNULL(status_bl,'') FROM dbo.T_suivi_commande_end_user_BL WITH(UPDLOCK,HOLDLOCK) WHERE cust_po=@P AND num_bl=@B", cn, tx)) { cmd.Parameters.Add(P("@P", SqlDbType.VarChar, po, 500)); cmd.Parameters.Add(P("@B", SqlDbType.VarChar, bl.Number, 100)); using (var r = cmd.ExecuteReader()) if (r.Read()) { id = Convert.ToInt32(r[0]); current = Convert.ToString(r[1]); } } if (id == 0) { using (var cmd = new SqlCommand("INSERT INTO dbo.T_suivi_commande_end_user_BL(cust_po,num_bl,status_bl,cust_nbr) VALUES(@P,@B,@S,@C)", cn, tx)) { cmd.Parameters.Add(P("@P", SqlDbType.VarChar, po, 500)); cmd.Parameters.Add(P("@B", SqlDbType.VarChar, bl.Number, 100)); cmd.Parameters.Add(P("@S", SqlDbType.VarChar, bl.Status, 100)); cmd.Parameters.Add(P("@C", SqlDbType.VarChar, customer, 100)); cmd.ExecuteNonQuery(); } tx.Commit(); return BlUpsertResult.Inserted; } if (!IsFinalBlStatus(current) && !string.Equals(current, bl.Status, StringComparison.OrdinalIgnoreCase)) { using (var cmd = new SqlCommand("UPDATE dbo.T_suivi_commande_end_user_BL SET status_bl=@S WHERE id_bl=@I", cn, tx)) { cmd.Parameters.Add(P("@S", SqlDbType.VarChar, bl.Status, 100)); cmd.Parameters.Add("@I", SqlDbType.Int).Value = id; cmd.ExecuteNonQuery(); } tx.Commit(); return BlUpsertResult.Updated; } tx.Commit(); return BlUpsertResult.Unchanged; } } }
        public void Close_Completed_Orders(
            string sql_con,
            string logs,
            string tmp_folder,
            string session_name)
        {
            imcaConnection = sql_con;
            string root = GetServicePath();
            logsFolder = Path.Combine(root, logs ?? "");
            tempFolder = Path.Combine(root, tmp_folder ?? "");
            sessionName = session_name ?? "";

            Directory.CreateDirectory(logsFolder);
            Directory.CreateDirectory(tempFolder);

            JsonFile configuration =
                JsonConvert.DeserializeObject<JsonFile>(
                    GetParameter(sql_con, ApplicationName) ?? "");

            if (configuration?.countries == null ||
                configuration.countries.Count == 0)
            {
                throw new InvalidOperationException(
                    ApplicationName +
                    " parameters are empty or invalid");
            }

            foreach (Country item in configuration.countries)
            {
                Apply(item, sql_con);

                if (!IsTrue(item.active))
                {
                    continue;
                }

                try
                {
                    ValidateCloseOrdersConfiguration();

                    int examined;
                    int closed = CloseCompletedCustomerPos(
                        out examined);

                    WriteLog(
                        "   Close_Completed_Orders summary" +
                        " - Open orders examined : " + examined +
                        " - Orders closed : " + closed +
                        " - Orders left open : " +
                        (examined - closed));
                }
                catch (Exception ex)
                {
                    WriteLog(
                        "   Close_Completed_Orders error - " +
                        Details(ex));

                    SendTechnicalAlert(
                        nameof(Close_Completed_Orders),
                        Details(ex));
                }
            }
        }

        private void ValidateCloseOrdersConfiguration()
        {
            if (string.IsNullOrWhiteSpace(sqlGestion))
            {
                throw new InvalidOperationException(
                    "Close_Completed_Orders configuration " +
                    "incomplete : sqlGestion");
            }
        }

        private int CloseCompletedCustomerPos(out int examined)
        {
            const string selectSql = @"
SELECT DISTINCT
       cp.id_cust_po,
       ISNULL(cp.cust_po, '') AS cust_po
FROM dbo.T_suivi_commande_end_user_cust_po cp
INNER JOIN dbo.T_suivi_commande_end_user_fichier f
    ON f.id_fichier = cp.id_fichier
INNER JOIN dbo.T_suivi_commande_end_user_mail m
    ON m.id_mail = f.id_mail
WHERE cp.ouvert = 1
  AND ISNULL(cp.cde_annule, 0) = 0
  AND cp.cust_po IS NOT NULL
  AND LTRIM(RTRIM(cp.cust_po)) <> ''
  AND DATEDIFF(DAY, m.date_mail, GETDATE()) <= @DELAI
ORDER BY cp.id_cust_po;";

            var candidates = new List<Tuple<int, string>>();

            using (var connection = new SqlConnection(sqlGestion))
            using (var command = new SqlCommand(
                selectSql,
                connection))
            {
                command.CommandTimeout = 300;
                command.Parameters.Add(
                    "@DELAI",
                    SqlDbType.Int).Value =
                    delaiRefreshBlTracking;

                connection.Open();

                using (SqlDataReader reader =
                    command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        candidates.Add(
                            Tuple.Create(
                                Convert.ToInt32(
                                    reader["id_cust_po"]),
                                Convert.ToString(
                                    reader["cust_po"]).Trim()));
                    }
                }
            }

            examined = candidates.Count;
            int closed = 0;

            foreach (Tuple<int, string> candidate in candidates)
            {
                if (!CanCloseCustomerPo(
                        candidate.Item1,
                        candidate.Item2))
                {
                    continue;
                }

                CloseCustomerPo(candidate.Item1);
                closed++;

                WriteLog(
                    "       Customer PO closed" +
                    " - ID : " + candidate.Item1 +
                    " - Customer PO : " + candidate.Item2);
            }

            return closed;
        }

        private bool CanCloseCustomerPo(int id, string po)
        {
            int incomplete = ScalarInt(
                sqlGestion,
                @"SELECT COUNT(*)
FROM dbo.T_suivi_commande_end_user_cust_po_line
WHERE id_cust_po = @I
  AND
  (
      (serial_number IS NULL OR LTRIM(RTRIM(serial_number)) = '')
      AND
      (IMEI1 IS NULL OR LTRIM(RTRIM(IMEI1)) = '')
  );",
                new SqlParameter("@I", SqlDbType.Int) { Value = id });

            if (incomplete > 0)
                return false;

            const string sql = @"
SELECT
    COUNT(*) AS total_bl,
    SUM(
        CASE
            WHEN status_bl IN ('INVOICED', 'FORCED BILLED') THEN 1
            ELSE 0
        END
    ) AS billed_bl,
    SUM(
        CASE
            WHEN ISNULL(status_bl, '') NOT IN
                 ('INVOICED', 'FORCED BILLED', 'BILLING ERROR', 'VOIDED') THEN 1
            ELSE 0
        END
    ) AS non_final_bl
FROM dbo.T_suivi_commande_end_user_BL
WHERE cust_po = @P;";

            using (var connection = new SqlConnection(sqlGestion))
            using (var command = new SqlCommand(sql, connection))
            {
                command.CommandTimeout = 300;
                command.Parameters.Add(P("@P", SqlDbType.VarChar, po, 500));
                connection.Open();

                using (var reader = command.ExecuteReader())
                {
                    if (!reader.Read())
                        return false;

                    int totalBl = reader["total_bl"] == DBNull.Value
                        ? 0
                        : Convert.ToInt32(reader["total_bl"]);
                    int billedBl = reader["billed_bl"] == DBNull.Value
                        ? 0
                        : Convert.ToInt32(reader["billed_bl"]);
                    int nonFinalBl = reader["non_final_bl"] == DBNull.Value
                        ? 0
                        : Convert.ToInt32(reader["non_final_bl"]);

                    return totalBl > 0 && billedBl > 0 && nonFinalBl == 0;
                }
            }
        }
        private void CloseCustomerPo(int id) => ExecuteNonQuery(sqlGestion, "UPDATE dbo.T_suivi_commande_end_user_cust_po SET ouvert=0 WHERE id_cust_po=@I", new SqlParameter("@I", SqlDbType.Int) { Value = id });

        public void Update_BL_Country_Codes(string sql_con, string logs, string tmp_folder, string session_name)
        {
            imcaConnection = sql_con; string root = GetServicePath();
            logsFolder = Path.Combine(root, logs ?? ""); tempFolder = Path.Combine(root, tmp_folder ?? ""); sessionName = session_name ?? "";
            Directory.CreateDirectory(logsFolder); Directory.CreateDirectory(tempFolder);
            JsonFile cfg = JsonConvert.DeserializeObject<JsonFile>(GetParameter(sql_con, ApplicationName) ?? "");
            if (cfg?.countries == null || cfg.countries.Count == 0) throw new InvalidOperationException(ApplicationName + " parameters are empty or invalid");
            foreach (Country item in cfg.countries)
            {
                Apply(item, sql_con); if (!IsTrue(item.active)) continue; ValidateCountryCodeConfiguration();
                try
                {
                    List<BlCountryCandidate> candidates = LoadBlCountryCandidates();
                    int imXmlSuccess = 0, dssSuccess = 0, notFound = 0, errors = 0;
                    WriteLog("   BL country code candidate(s) : " + candidates.Count + " - Refresh delay : " + delaiRefreshBlTracking + " day(s)");
                    foreach (BlCountryCandidate candidate in candidates)
                    {
                        try
                        {
                            if (!TryGetOrderNumber(candidate.BlNumber, out string orderNumber)) { errors++; WriteLog("       Invalid BL number - ID : " + candidate.BlId + " - BL : " + candidate.BlNumber); continue; }
                            int requestId = InsertImXmlRequest(candidate.CustomerNumber, "Order Detail : recherche code pays", candidate.CustomerPo);
                            ImXmlOrderDetailResult result;
                            try { result = ExecuteImXmlOrderDetail(candidate.CustomerPo, candidate.CustomerNumber, requestId, orderNumber); }
                            catch (Exception ex) { result = new ImXmlOrderDetailResult { Message = "IMXML technical error : " + Details(ex) }; }
                            UpdateImXmlRequest(requestId, result.Message);
                            if (result.Success && !string.IsNullOrWhiteSpace(result.CountryCode))
                            {
                                UpdateBlCountryCode(candidate.BlId, result.CountryCode, result.ZipCode); imXmlSuccess++;
                                WriteLog("       BL country code updated from IMXML - BL : " + candidate.BlNumber + " - Country : " + result.CountryCode + " - Zip code : " + result.ZipCode); continue;
                            }
                            if (result.OrderNotFound)
                            {
                                string dssCountryCode = FindCountryCodeInDss(candidate.CustomerNumber, orderNumber, candidate.CustomerPo);
                                if (!string.IsNullOrWhiteSpace(dssCountryCode)) { UpdateBlCountryCode(candidate.BlId, dssCountryCode, null); dssSuccess++; WriteLog("       BL country code updated from DSS - BL : " + candidate.BlNumber + " - Country : " + dssCountryCode); }
                                else { notFound++; WriteLog("       BL country code not found in DSS - BL : " + candidate.BlNumber + " - Customer PO : " + candidate.CustomerPo); }
                                continue;
                            }
                            errors++; WriteLog("       IMXML Order Detail error - BL : " + candidate.BlNumber + " - Customer PO : " + candidate.CustomerPo + " - Message : " + result.Message);
                            if (IsTrue(debug) && !string.IsNullOrWhiteSpace(result.DebugResponse)) WriteLog("       IMXML Order Detail response - BL : " + candidate.BlNumber + " - Content : " + result.DebugResponse);
                        }
                        catch (Exception ex) { errors++; WriteLog("       BL country code processing error - BL ID : " + candidate.BlId + " - BL : " + candidate.BlNumber + " - " + Details(ex)); }
                    }
                    WriteLog("   Update_BL_Country_Codes summary - Candidates : " + candidates.Count + " - Updated from IMXML : " + imXmlSuccess + " - Updated from DSS : " + dssSuccess + " - Not found : " + notFound + " - Errors : " + errors);
                }
                catch (Exception ex) { WriteLog("   Update_BL_Country_Codes error - " + Details(ex)); SendTechnicalAlert(nameof(Update_BL_Country_Codes), Details(ex)); }
            }
        }
        private void ValidateCountryCodeConfiguration()
        {
            var missing = new List<string>();
            if (string.IsNullOrWhiteSpace(sqlGestion)) missing.Add("sqlGestion"); if (string.IsNullOrWhiteSpace(sqlDss)) missing.Add("sqlDss");
            if (string.IsNullOrWhiteSpace(imXmlOrderDetailUrl)) missing.Add("imXmlOrderDetailUrl"); if (string.IsNullOrWhiteSpace(imXmlOrderDetailTemplate)) missing.Add("imXmlOrderDetailTemplate");
            if (missing.Count > 0) throw new InvalidOperationException("Update_BL_Country_Codes configuration incomplete : " + string.Join(", ", missing));
            imXmlOrderDetailUrl = ValidateHttpUrl(imXmlOrderDetailUrl, imXmlOrderDetailUrlParameterGlobal);
            var d = new XmlDocument(); d.LoadXml(imXmlOrderDetailTemplate);
            string[] required = { "/*[local-name()='OrderDetailRequest']", "/*[local-name()='OrderDetailRequest']/*[local-name()='Version']", "//*[local-name()='TransactionHeader']/*[local-name()='SenderID']", "//*[local-name()='TransactionHeader']/*[local-name()='ReceiverID']", "//*[local-name()='TransactionHeader']/*[local-name()='CountryCode']", "//*[local-name()='TransactionHeader']/*[local-name()='LoginID']", "//*[local-name()='TransactionHeader']/*[local-name()='Password']", "//*[local-name()='TransactionHeader']/*[local-name()='TransactionID']", "//*[local-name()='OrderHeaderInfo']/*[local-name()='BranchOrderNumber']", "//*[local-name()='OrderHeaderInfo']/*[local-name()='CustomerPO']" };
            foreach (string xpath in required) if (d.SelectSingleNode(xpath) == null) throw new InvalidOperationException("Required node missing from IMXML Order Detail template : " + xpath);
        }
        private List<BlCountryCandidate> LoadBlCountryCandidates()
        {
            const string sql = @"SELECT DISTINCT bl.id_bl,ISNULL(bl.num_bl,'') num_bl,ISNULL(bl.cust_po,'') cust_po,ISNULL(bl.cust_nbr,ISNULL(m.cust_nbr,'')) cust_nbr
FROM dbo.T_suivi_commande_end_user_BL bl JOIN dbo.T_suivi_commande_end_user_cust_po cp ON cp.cust_po=bl.cust_po
JOIN dbo.T_suivi_commande_end_user_fichier f ON f.id_fichier=cp.id_fichier JOIN dbo.T_suivi_commande_end_user_mail m ON m.id_mail=f.id_mail
WHERE (bl.code_pays IS NULL OR LTRIM(RTRIM(bl.code_pays))='') AND ISNULL(cp.cde_annule,0)=0 AND bl.num_bl IS NOT NULL AND LTRIM(RTRIM(bl.num_bl))<>''
AND cp.cust_po IS NOT NULL AND LTRIM(RTRIM(cp.cust_po))<>'' AND DATEDIFF(DAY,m.date_mail,GETDATE())<=@DELAI ORDER BY bl.id_bl;";
            var result = new List<BlCountryCandidate>();
            using (var cn = new SqlConnection(sqlGestion)) using (var cmd = new SqlCommand(sql, cn)) { cmd.CommandTimeout = 300; cmd.Parameters.Add("@DELAI", SqlDbType.Int).Value = delaiRefreshBlTracking; cn.Open(); using (var r = cmd.ExecuteReader()) while (r.Read()) result.Add(new BlCountryCandidate { BlId = Convert.ToInt32(r["id_bl"]), BlNumber = Convert.ToString(r["num_bl"]).Trim(), CustomerPo = Convert.ToString(r["cust_po"]).Trim(), CustomerNumber = Convert.ToString(r["cust_nbr"]).Trim() }); }
            return result;
        }
        private static bool TryGetOrderNumber(string blNumber, out string orderNumber)
        {
            orderNumber = ""; string normalized = (blNumber ?? "").Trim().Replace(" ", ""); if (string.IsNullOrWhiteSpace(normalized)) return false;
            string[] parts = normalized.Split(new[] { '-' }, StringSplitOptions.RemoveEmptyEntries); orderNumber = parts.Length > 0 ? parts[0].Trim() : normalized;
            return Regex.IsMatch(orderNumber, @"^[A-Za-z0-9]+$");
        }
        private ImXmlOrderDetailResult ExecuteImXmlOrderDetail(string customerPo, string customerNumber, int requestId, string branchOrderNumber)
        {
            XmlDocument request = BuildImXmlOrderDetailRequest(customerPo, requestId, branchOrderNumber, LoadImXmlCredentials(customerNumber));
            if (IsTrue(debug)) WriteLog("       IMXML Order Detail request - Customer PO : " + customerPo + " - Branch order number : " + branchOrderNumber);
            return ParseImXmlOrderDetailResponse(PostImXmlRequest(request, imXmlOrderDetailUrl));
        }
        private XmlDocument BuildImXmlOrderDetailRequest(string customerPo, int requestId, string branchOrderNumber, ImXmlCredentials credentials)
        {
            var d = new XmlDocument(); d.LoadXml(imXmlOrderDetailTemplate); SetNode(d, "//*[local-name()='LoginID']", credentials.Login); SetNode(d, "//*[local-name()='Password']", credentials.Password);
            SetNode(d, "//*[local-name()='TransactionID']", requestId.ToString(CultureInfo.InvariantCulture)); SetNode(d, "//*[local-name()='BranchOrderNumber']", branchOrderNumber); SetNode(d, "//*[local-name()='CustomerPO']", customerPo); return d;
        }
        private static ImXmlOrderDetailResult ParseImXmlOrderDetailResponse(string text)
        {
            var z = new ImXmlOrderDetailResult(); if (string.IsNullOrWhiteSpace(text)) { z.Message = "Empty IMXML response"; return z; }
            var d = new XmlDocument(); try { d.LoadXml(text); } catch (Exception ex) { z.Message = "Invalid IMXML response : " + Truncate(ex.Message, 150); z.DebugResponse = Truncate(Regex.Replace(text ?? "", @"\s+", " ").Trim(), 500); return z; }
            XmlNode root = d.SelectSingleNode("//*[local-name()='OrderDetailResponse']");
            if (root == null) { string rootName = d.DocumentElement == null ? "none" : d.DocumentElement.LocalName; z.Message = "Unexpected IMXML response - Root : " + rootName; z.DebugResponse = Truncate(Regex.Replace(d.OuterXml ?? "", @"\s+", " ").Trim(), 500); return z; }
            string error = (root.SelectSingleNode(".//*[local-name()='ErrorStatus']")?.InnerText ?? "").Trim();
            if (error != "") { z.Message = Truncate(error, 250); z.OrderNotFound = error.IndexOf("Order Not Found", StringComparison.OrdinalIgnoreCase) >= 0; return z; }
            XmlNode suffix = root.SelectSingleNode(".//*[local-name()='OrderInformation']/*[local-name()='OrderSuffix']"); if (suffix == null) { z.Message = "No OrderSuffix found in IMXML response"; return z; }
            z.CountryCode = (suffix.SelectSingleNode(".//*[local-name()='ShipToInformation']/*[local-name()='ShipToCountryCode']")?.InnerText ?? "").Trim();
            z.ZipCode = (suffix.SelectSingleNode(".//*[local-name()='ShipToInformation']/*[local-name()='ShipToPostalCode']")?.InnerText ?? "").Trim();
            if (string.IsNullOrWhiteSpace(z.CountryCode)) { z.Message = "Country code missing from IMXML response"; return z; }
            z.Success = true; z.Message = "Requête OK"; return z;
        }
        private void UpdateBlCountryCode(int blId, string countryCode, string zipCode)
        {
            var zip = new SqlParameter("@ZIP", SqlDbType.VarChar, 50) { Value = string.IsNullOrWhiteSpace(zipCode) ? (object)DBNull.Value : zipCode.Trim() };
            ExecuteNonQuery(sqlGestion, "UPDATE dbo.T_suivi_commande_end_user_BL SET code_pays=@COUNTRY,zip_code=CASE WHEN @ZIP IS NULL THEN zip_code ELSE @ZIP END WHERE id_bl=@ID;", P("@COUNTRY", SqlDbType.VarChar, countryCode?.Trim() ?? "", 10), zip, new SqlParameter("@ID", SqlDbType.Int) { Value = blId });
        }
        private string FindCountryCodeInDss(string customerNumber, string orderNumber, string customerPo)
        {
            const string sql = @"SELECT TOP(1) ISNULL(a.Dest_Country,'') FROM dbo.Shipment_Address a WHERE a.Br_Invoice_Nbr=(SELECT TOP(1) h.invoice_nbr FROM dbo.shipment_header h WHERE h.order_nbr=@ORDER AND h.branch_customer_nbr=@CUSTOMER AND LTRIM(RTRIM(h.CustPoNbr))=@PO ORDER BY h.invoice_date DESC);";
            return ScalarString(sqlDss, sql, P("@ORDER", SqlDbType.VarChar, orderNumber ?? "", 50), P("@CUSTOMER", SqlDbType.VarChar, customerNumber ?? "", 50), P("@PO", SqlDbType.VarChar, customerPo ?? "", 500));
        }
        public void Update_BL_Tracking(string sql_con, string logs, string tmp_folder, string session_name)
        {
            imcaConnection = sql_con;
            string root = GetServicePath();
            logsFolder = Path.Combine(root, logs ?? "");
            tempFolder = Path.Combine(root, tmp_folder ?? "");
            sessionName = session_name ?? "";
            Directory.CreateDirectory(logsFolder);
            Directory.CreateDirectory(tempFolder);

            JsonFile cfg = JsonConvert.DeserializeObject<JsonFile>(
                GetParameter(sql_con, ApplicationName) ?? "");
            if (cfg?.countries == null || cfg.countries.Count == 0)
                throw new InvalidOperationException(
                    ApplicationName + " parameters are empty or invalid");

            foreach (Country item in cfg.countries)
            {
                Apply(item, sql_con);
                if (!IsTrue(item.active)) continue;

                try
                {
                    ValidateTrackingConfiguration();
                    List<BlTrackingCandidate> candidates = LoadBlTrackingCandidates();
                    int updated = 0;
                    int notFound = 0;
                    int noPackage = 0;
                    int invalidResponse = 0;
                    int imXmlErrors = 0;
                    int invalidShipDate = 0;
                    int technicalErrors = 0;
                    int processingErrors = 0;

                    WriteLog(
                        "   BL tracking candidate(s) : " + candidates.Count +
                        " - Refresh delay : " + delaiRefreshBlTracking + " day(s)");

                    foreach (BlTrackingCandidate candidate in candidates)
                    {
                        try
                        {
                            if (!TrySplitTrackingBlNumber(
                                    candidate.BlNumber,
                                    out string branchOrderNumber,
                                    out string suffixNumber))
                            {
                                processingErrors++;
                                WriteLog(
                                    "       Invalid BL number for tracking" +
                                    " - BL ID : " + candidate.BlId +
                                    " - BL : " + candidate.BlNumber);
                                continue;
                            }

                            string invoiceNumber = FindTrackingInvoiceInDss(
                                branchOrderNumber,
                                candidate.CustomerPo,
                                candidate.CustomerNumber);

                            int requestId = InsertImXmlRequest(
                                candidate.CustomerNumber,
                                "Order Tracking",
                                candidate.CustomerPo);

                            ImXmlTrackingResult result;
                            try
                            {
                                result = ExecuteImXmlOrderTracking(
                                    candidate.CustomerPo,
                                    candidate.CustomerNumber,
                                    requestId,
                                    branchOrderNumber,
                                    suffixNumber,
                                    invoiceNumber);
                            }
                            catch (Exception ex)
                            {
                                result = new ImXmlTrackingResult
                                {
                                    Outcome = TrackingOutcome.TechnicalError,
                                    Message = "IMXML technical error : " + Details(ex)
                                };
                            }

                            UpdateImXmlRequest(requestId, result.Message);

                            if (!result.Success)
                            {
                                switch (result.Outcome)
                                {
                                    case TrackingOutcome.NoPackage:
                                        noPackage++;
                                        break;
                                    case TrackingOutcome.InvalidResponse:
                                        invalidResponse++;
                                        break;
                                    case TrackingOutcome.ImXmlError:
                                        imXmlErrors++;
                                        break;
                                    case TrackingOutcome.InvalidShipDate:
                                        invalidShipDate++;
                                        break;
                                    case TrackingOutcome.TechnicalError:
                                        technicalErrors++;
                                        break;
                                    default:
                                        notFound++;
                                        break;
                                }

                                WriteLog(
                                    "       BL tracking not updated" +
                                    " - Outcome : " + result.Outcome +
                                    " - BL : " + candidate.BlNumber +
                                    " - Invoice : " + invoiceNumber +
                                    " - Message : " + result.Message);
                                if (IsTrue(debug) &&
                                    !string.IsNullOrWhiteSpace(result.DebugResponse))
                                {
                                    WriteLog(
                                        "       IMXML Order Tracking response" +
                                        " - BL : " + candidate.BlNumber +
                                        " - Content : " + result.DebugResponse);
                                }
                                continue;
                            }

                            CarrierMapping mapping = FindCarrierMapping(
                                result.CarrierCode);

                            UpdateBlTrackingAndFlagFile(
                                candidate.BlId,
                                candidate.FileId,
                                result,
                                mapping);
                            updated++;

                            WriteLog(
                                "       BL tracking updated" +
                                " - BL : " + candidate.BlNumber +
                                " - Tracking : " + result.TrackingNumber +
                                " - Carrier impulse : " + result.CarrierCode +
                                " - Carrier mapped : " + mapping.CarrierCode +
                                " - Carrier name : " + mapping.CarrierName);
                        }
                        catch (Exception ex)
                        {
                            processingErrors++;
                            WriteLog(
                                "       BL tracking processing error" +
                                " - BL ID : " + candidate.BlId +
                                " - BL : " + candidate.BlNumber +
                                " - " + Details(ex));
                        }
                    }

                    WriteLog(
                        "   Update_BL_Tracking summary" +
                        " - Candidates : " + candidates.Count +
                        " - Updated : " + updated +
                        " - Not found : " + notFound +
                        " - No package : " + noPackage +
                        " - Invalid response : " + invalidResponse +
                        " - IMXML errors : " + imXmlErrors +
                        " - Invalid ship date : " + invalidShipDate +
                        " - Technical errors : " + technicalErrors +
                        " - Processing errors : " + processingErrors);
                }
                catch (Exception ex)
                {
                    WriteLog("   Update_BL_Tracking error - " + Details(ex));
                    SendTechnicalAlert(nameof(Update_BL_Tracking), Details(ex));
                }
            }
        }

        private void ValidateTrackingConfiguration()
        {
            var missing = new List<string>();
            if (string.IsNullOrWhiteSpace(sqlGestion)) missing.Add("sqlGestion");
            if (string.IsNullOrWhiteSpace(sqlDss)) missing.Add("sqlDss");
            if (string.IsNullOrWhiteSpace(imXmlOrderTrackingUrl)) missing.Add("imXmlOrderTrackingUrl");
            if (string.IsNullOrWhiteSpace(imXmlOrderTrackingTemplate)) missing.Add("imXmlOrderTrackingTemplate");
            if (missing.Count > 0)
                throw new InvalidOperationException(
                    "Update_BL_Tracking configuration incomplete : " +
                    string.Join(", ", missing));

            imXmlOrderTrackingUrl = ValidateHttpUrl(
                imXmlOrderTrackingUrl,
                imXmlOrderTrackingUrlParameterGlobal);

            var document = new XmlDocument();
            document.LoadXml(imXmlOrderTrackingTemplate);
            string[] required =
            {
                "/*[local-name()='OrderTrackingRequest']",
                "/*[local-name()='OrderTrackingRequest']/*[local-name()='Version']",
                "//*[local-name()='TransactionHeader']/*[local-name()='SenderID']",
                "//*[local-name()='TransactionHeader']/*[local-name()='ReceiverID']",
                "//*[local-name()='TransactionHeader']/*[local-name()='CountryCode']",
                "//*[local-name()='TransactionHeader']/*[local-name()='LoginID']",
                "//*[local-name()='TransactionHeader']/*[local-name()='Password']",
                "//*[local-name()='TransactionHeader']/*[local-name()='TransactionID']",
                "//*[local-name()='TrackingRequestHeader']/*[local-name()='CustomerPO']"
            };
            foreach (string xpath in required)
                if (document.SelectSingleNode(xpath) == null)
                    throw new InvalidOperationException(
                        "Required node missing from IMXML Order Tracking template : " + xpath);
        }

        private List<BlTrackingCandidate> LoadBlTrackingCandidates()
        {
            const string sql = @"
WITH TrackingSource AS
(
    SELECT
        bl.id_bl,
        cp.id_fichier,
        ISNULL(bl.num_bl, '') AS num_bl,
        ISNULL(bl.cust_po, '') AS cust_po,
        ISNULL(bl.cust_nbr, ISNULL(m.cust_nbr, '')) AS cust_nbr,
        RIGHT(REPLACE(LTRIM(RTRIM(bl.num_bl)), ' ', ''), 2) AS bl_suffix,
        ROW_NUMBER() OVER
        (
            PARTITION BY
                bl.cust_po,
                RIGHT(REPLACE(LTRIM(RTRIM(bl.num_bl)), ' ', ''), 2)
            ORDER BY
                bl.id_bl,
                cp.id_fichier DESC
        ) AS suffix_row_number
    FROM dbo.T_suivi_commande_end_user_BL bl
    INNER JOIN dbo.T_suivi_commande_end_user_cust_po cp
        ON cp.cust_po = bl.cust_po
    INNER JOIN dbo.T_suivi_commande_end_user_fichier f
        ON f.id_fichier = cp.id_fichier
    INNER JOIN dbo.T_suivi_commande_end_user_mail m
        ON m.id_mail = f.id_mail
    WHERE bl.status_bl IN
          (
              'INVOICED',
              'FORCED BILLED',
              'BILLING ERROR'
          )
      AND cp.ouvert = 1
      AND ISNULL(cp.cde_annule, 0) = 0
      AND DATEDIFF(DAY, m.date_mail, GETDATE()) <= @DELAI
      AND bl.num_bl IS NOT NULL
      AND LTRIM(RTRIM(bl.num_bl)) <> ''
      AND cp.cust_po IS NOT NULL
      AND LTRIM(RTRIM(cp.cust_po)) <> ''
      AND RIGHT(
              REPLACE(LTRIM(RTRIM(bl.num_bl)), ' ', ''),
              2
          ) IN ('11', '12')
      AND LTRIM(RTRIM(ISNULL(bl.track_id, ''))) = ''
      AND LTRIM(RTRIM(ISNULL(bl.num_suivi_transporteur, ''))) = ''
      AND NOT EXISTS
      (
          SELECT 1
          FROM dbo.T_suivi_commande_end_user_BL tracked_bl
          WHERE tracked_bl.cust_po = bl.cust_po
            AND tracked_bl.status_bl IN
                (
                    'INVOICED',
                    'FORCED BILLED',
                    'BILLING ERROR'
                )
            AND RIGHT(
                    REPLACE(
                        LTRIM(RTRIM(tracked_bl.num_bl)),
                        ' ',
                        ''
                    ),
                    2
                ) = RIGHT(
                    REPLACE(
                        LTRIM(RTRIM(bl.num_bl)),
                        ' ',
                        ''
                    ),
                    2
                )
            AND
            (
                LTRIM(RTRIM(ISNULL(tracked_bl.track_id, ''))) <> ''
                OR
                LTRIM(RTRIM(
                    ISNULL(tracked_bl.num_suivi_transporteur, '')
                )) <> ''
            )
      )
)
SELECT
    id_bl,
    id_fichier,
    num_bl,
    cust_po,
    cust_nbr
FROM TrackingSource
WHERE suffix_row_number = 1
ORDER BY
    cust_po,
    bl_suffix,
    id_bl;";

            var result = new List<BlTrackingCandidate>();

            using (var connection = new SqlConnection(sqlGestion))
            using (var command = new SqlCommand(sql, connection))
            {
                command.CommandTimeout = 300;
                command.Parameters.Add(
                    "@DELAI",
                    SqlDbType.Int).Value =
                    delaiRefreshBlTracking;

                connection.Open();

                using (SqlDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        result.Add(new BlTrackingCandidate
                        {
                            BlId = Convert.ToInt32(reader["id_bl"]),
                            FileId = Convert.ToInt32(reader["id_fichier"]),
                            BlNumber = Convert.ToString(reader["num_bl"]).Trim(),
                            CustomerPo = Convert.ToString(reader["cust_po"]).Trim(),
                            CustomerNumber = Convert.ToString(reader["cust_nbr"]).Trim()
                        });
                    }
                }
            }

            return result;
        }

        private static bool TrySplitTrackingBlNumber(
            string blNumber,
            out string branchOrderNumber,
            out string suffixNumber)
        {
            branchOrderNumber = "";
            suffixNumber = "";
            string normalized = (blNumber ?? "").Trim().Replace(" ", "");
            if (string.IsNullOrWhiteSpace(normalized)) return false;
            string[] parts = normalized.Split(
                new[] { '-' },
                StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2) return false;
            branchOrderNumber = parts[0].Trim();
            suffixNumber = parts[parts.Length - 1].Trim();
            return Regex.IsMatch(branchOrderNumber, @"^[A-Za-z0-9]+$") &&
                   Regex.IsMatch(suffixNumber, @"^[A-Za-z0-9]+$");
        }

        private string FindTrackingInvoiceInDss(
            string orderNumber,
            string customerPo,
            string customerNumber)
        {
            const string sql = @"
SELECT TOP (1) ISNULL(invoice_nbr, '')
FROM dbo.shipment_header
WHERE order_nbr = @ORDER
  AND LTRIM(RTRIM(ISNULL(CustPoNbr, ''))) = @PO
  AND branch_customer_nbr = @CUSTOMER
ORDER BY invoice_date DESC;";
            return ScalarString(
                sqlDss,
                sql,
                P("@ORDER", SqlDbType.VarChar, orderNumber, 50),
                P("@PO", SqlDbType.VarChar, customerPo, 500),
                P("@CUSTOMER", SqlDbType.VarChar, customerNumber, 50));
        }

        private ImXmlTrackingResult ExecuteImXmlOrderTracking(
            string customerPo,
            string customerNumber,
            int requestId,
            string branchOrderNumber,
            string suffixNumber,
            string invoiceNumber)
        {
            XmlDocument request = BuildImXmlOrderTrackingRequest(
                customerPo,
                requestId,
                LoadImXmlCredentials(customerNumber));
            if (IsTrue(debug))
                WriteLog(
                    "       IMXML Order Tracking request" +
                    " - Customer PO : " + customerPo +
                    " - Branch order number : " + branchOrderNumber +
                    " - Suffix : " + suffixNumber +
                    " - Invoice : " + invoiceNumber);
            string response = PostImXmlRequest(
                request,
                imXmlOrderTrackingUrl);
            return ParseImXmlTrackingResponse(
                response,
                branchOrderNumber,
                suffixNumber,
                invoiceNumber);
        }

        private XmlDocument BuildImXmlOrderTrackingRequest(
            string customerPo,
            int requestId,
            ImXmlCredentials credentials)
        {
            var document = new XmlDocument();
            document.LoadXml(imXmlOrderTrackingTemplate);
            SetNode(document, "//*[local-name()='LoginID']", credentials.Login);
            SetNode(document, "//*[local-name()='Password']", credentials.Password);
            SetNode(document, "//*[local-name()='TransactionID']",
                requestId.ToString(CultureInfo.InvariantCulture));
            SetNode(document, "//*[local-name()='TrackingRequestHeader']/*[local-name()='CustomerPO']", customerPo);
            return document;
        }

        private static ImXmlTrackingResult ParseImXmlTrackingResponse(
            string text,
            string branchOrderNumber,
            string suffixNumber,
            string invoiceNumber)
        {
            var result = new ImXmlTrackingResult();
            if (string.IsNullOrWhiteSpace(text))
            {
                result.Outcome = TrackingOutcome.InvalidResponse;
                result.Message = "Empty IMXML response";
                return result;
            }

            var document = new XmlDocument();
            try { document.LoadXml(text); }
            catch (Exception ex)
            {
                result.Outcome = TrackingOutcome.InvalidResponse;
                result.Message = "Invalid IMXML response : " +
                    Truncate(ex.Message, 150);
                result.DebugResponse = Truncate(
                    Regex.Replace(text ?? "", @"\s+", " ").Trim(), 500);
                return result;
            }

            XmlNode root = document.SelectSingleNode(
                "//*[local-name()='OrderTrackingResponse']");
            if (root == null)
            {
                result.Outcome = TrackingOutcome.InvalidResponse;
                result.Message = "Unexpected IMXML tracking response";
                result.DebugResponse = Truncate(
                    Regex.Replace(document.OuterXml ?? "", @"\s+", " ").Trim(),
                    500);
                return result;
            }

            string error = (root.SelectSingleNode(
                ".//*[local-name()='TransactionHeader']/*[local-name()='ErrorStatus']")
                ?.InnerText ?? "").Trim();
            if (!string.IsNullOrWhiteSpace(error))
            {
                result.Outcome = TrackingOutcome.ImXmlError;
                result.Message = Truncate(error, 250);
                return result;
            }

            foreach (XmlNode order in root.SelectNodes(
                ".//*[local-name()='Order']"))
            {
                string responseOrder = (order.SelectSingleNode(
                    "./*[local-name()='BranchOrderNumber']")?.InnerText ?? "").Trim();
                foreach (XmlNode suffix in order.SelectNodes(
                    "./*[local-name()='Suffix']"))
                {
                    string responseSuffix =
                        (suffix.Attributes?["SuffixNumber"]?.Value ?? "").Trim();
                    bool blMatch =
                        string.Equals(responseOrder, branchOrderNumber,
                            StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(responseSuffix, suffixNumber,
                            StringComparison.OrdinalIgnoreCase);
                    bool invoiceMatch =
                        !string.IsNullOrWhiteSpace(invoiceNumber) &&
                        string.Equals(responseOrder + responseSuffix,
                            invoiceNumber.Replace("-", "").Replace(" ", ""),
                            StringComparison.OrdinalIgnoreCase);
                    if (!blMatch && !invoiceMatch) continue;

                    XmlNode package = suffix.SelectSingleNode(
                        "./*[local-name()='Package']");
                    if (package == null)
                    {
                        result.Outcome = TrackingOutcome.NoPackage;
                        result.Message = "Tracking package not found";
                        return result;
                    }

                    string tracking =
                        (package.Attributes?["ID"]?.Value ?? "").Trim();
                    string shipDateText = (package.SelectSingleNode(
                        "./*[local-name()='ShipDate']")?.InnerText ?? "").Trim();
                    string carrier =
                        (suffix.SelectSingleNode("./*[local-name()='Carrier']")
                            ?.Attributes?["Code"]?.Value ?? "").Trim();

                    if (string.IsNullOrWhiteSpace(tracking))
                    {
                        result.Outcome = TrackingOutcome.NoPackage;
                        result.Message = "Tracking number not found";
                        return result;
                    }
                    if (!DateTime.TryParseExact(
                            shipDateText,
                            "yyyy-MM-dd",
                            CultureInfo.InvariantCulture,
                            DateTimeStyles.None,
                            out DateTime shipDate))
                    {
                        result.Outcome = TrackingOutcome.InvalidShipDate;
                        result.Message = "Invalid or missing tracking ship date";
                        return result;
                    }

                    result.Success = true;
                    result.Outcome = TrackingOutcome.Success;
                    result.Message = "Requête OK";
                    result.TrackingNumber = tracking;
                    result.CarrierCode = carrier;
                    result.ShipDate = shipDate;
                    return result;
                }
            }

            result.Outcome = TrackingOutcome.NotFound;
            result.Message = "BL or invoice not found in tracking response";
            return result;
        }

        private CarrierMapping FindCarrierMapping(string impulseCarrierCode)
        {
            var result = new CarrierMapping();
            if (string.IsNullOrWhiteSpace(impulseCarrierCode)) return result;

            const string sql = @"
SELECT TOP (1)
       ISNULL(cc.carrier_code, '') AS carrier_code,
       ISNULL(t.nom_transporteur, '') AS nom_transporteur
FROM dbo.T_suivi_commande_end_user_carrier_code cc
LEFT JOIN dbo.T_suivi_commande_end_user_transporteur t
    ON t.carrier_code = cc.carrier_code
WHERE cc.carrier_impulse LIKE @IMPULSE
ORDER BY CASE WHEN t.nom_transporteur IS NULL THEN 1 ELSE 0 END,
         t.nom_transporteur;";

            using (var connection = new SqlConnection(sqlGestion))
            using (var command = new SqlCommand(sql, connection))
            {
                command.CommandTimeout = 300;
                command.Parameters.Add(P(
                    "@IMPULSE",
                    SqlDbType.VarChar,
                    impulseCarrierCode.Trim() + "%",
                    100));
                connection.Open();
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        result.CarrierCode = Convert.ToString(
                            reader["carrier_code"]).Trim();
                        result.CarrierName = Convert.ToString(
                            reader["nom_transporteur"]).Trim();
                    }
                }
            }
            return result;
        }

        private void UpdateBlTrackingAndFlagFile(
            int blId,
            int fileId,
            ImXmlTrackingResult tracking,
            CarrierMapping mapping)
        {
            const string updateBlSql = @"
UPDATE dbo.T_suivi_commande_end_user_BL
SET track_id = @TRACK,
    ship_date = @SHIP_DATE,
    carrier_code =
        CASE
            WHEN @IMPULSE_CARRIER = '' THEN carrier_code
            ELSE @IMPULSE_CARRIER
        END,
    num_suivi_transporteur =
        CASE
            WHEN @MAPPED_CARRIER <> ''
             AND @MAPPED_CARRIER <> 'OT' THEN @TRACK
            ELSE num_suivi_transporteur
        END,
    nom_transporteur =
        CASE
            WHEN @MAPPED_CARRIER <> ''
             AND @MAPPED_CARRIER <> 'OT'
             AND @CARRIER_NAME <> '' THEN @CARRIER_NAME
            ELSE nom_transporteur
        END
WHERE id_bl = @ID;";

            const string updateFileSql = @"
UPDATE dbo.T_suivi_commande_end_user_fichier
SET a_envoyer_o_n = 'O',
    info_tracking_updated_o_n = 'O'
WHERE id_fichier = @FILE_ID;";

            using (var connection = new SqlConnection(sqlGestion))
            {
                connection.Open();
                using (SqlTransaction transaction = connection.BeginTransaction())
                {
                    try
                    {
                        using (var command = new SqlCommand(
                            updateBlSql,
                            connection,
                            transaction))
                        {
                            command.CommandTimeout = 300;
                            command.Parameters.Add(P(
                                "@TRACK",
                                SqlDbType.VarChar,
                                tracking.TrackingNumber,
                                500));
                            command.Parameters.Add(
                                new SqlParameter("@SHIP_DATE", SqlDbType.Date)
                                {
                                    Value = tracking.ShipDate.HasValue
                                        ? (object)tracking.ShipDate.Value.Date
                                        : DBNull.Value
                                });
                            command.Parameters.Add(P(
                                "@IMPULSE_CARRIER",
                                SqlDbType.VarChar,
                                tracking.CarrierCode,
                                100));
                            command.Parameters.Add(P(
                                "@MAPPED_CARRIER",
                                SqlDbType.VarChar,
                                mapping.CarrierCode,
                                100));
                            command.Parameters.Add(P(
                                "@CARRIER_NAME",
                                SqlDbType.VarChar,
                                mapping.CarrierName,
                                500));
                            command.Parameters.Add(
                                new SqlParameter("@ID", SqlDbType.Int)
                                {
                                    Value = blId
                                });

                            if (command.ExecuteNonQuery() != 1)
                            {
                                throw new InvalidOperationException(
                                    "BL tracking update affected an unexpected number of rows" +
                                    " - BL ID : " + blId);
                            }
                        }

                        using (var command = new SqlCommand(
                            updateFileSql,
                            connection,
                            transaction))
                        {
                            command.CommandTimeout = 300;
                            command.Parameters.Add(
                                new SqlParameter("@FILE_ID", SqlDbType.Int)
                                {
                                    Value = fileId
                                });

                            if (command.ExecuteNonQuery() != 1)
                            {
                                throw new InvalidOperationException(
                                    "Tracking file flag update affected an unexpected number of rows" +
                                    " - File ID : " + fileId);
                            }
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

        public void Send_Files_To_Customers(
            string sql_con,
            string logs,
            string tmp_folder,
            string session_name)
        {
            imcaConnection = sql_con;
            string root = GetServicePath();
            logsFolder = Path.Combine(root, logs ?? "");
            tempFolder = Path.Combine(root, tmp_folder ?? "");
            sessionName = session_name ?? "";
            Directory.CreateDirectory(logsFolder);
            Directory.CreateDirectory(tempFolder);
            JsonFile cfg = JsonConvert.DeserializeObject<JsonFile>(
                GetParameter(sql_con, ApplicationName) ?? "");
            if (cfg?.countries == null || cfg.countries.Count == 0)
                throw new InvalidOperationException(ApplicationName + " parameters are empty or invalid");
            foreach (Country item in cfg.countries)
            {
                Apply(item, sql_con);
                if (!IsTrue(item.active)) continue;

                if (!IsCustomerFileSendHour(DateTime.Now))
                {
                    WriteLog(
                        "   Customer file sending skipped outside configured hours" +
                        " - Current time : " +
                        DateTime.Now.ToString("HH:mm") +
                        " - Configured hours : " +
                        string.Join(", ", sendFilesHours));
                    continue;
                }

                Validate();
                if (string.IsNullOrWhiteSpace(econocomSendAs))
                    throw new InvalidOperationException("fr_econocom_graph_send_as is empty");
                graph = ConnectGraph();
                imeiSources = LoadImeiSources(sql_con, item.country);
                WriteLog("   IMEI source table(s) loaded : " + imeiSources.Count);
                List<PendingCustomerFile> pendingFiles = LoadPendingCustomerFiles();
                WriteLog("   Customer file(s) to send : " + pendingFiles.Count);
                foreach (PendingCustomerFile pending in pendingFiles)
                {
                    string filePath = Path.Combine(
                        tempFolder,
                        "ECONOCOM_" + pending.FileId + "_" + Path.GetFileName(pending.FileName));
                    try
                    {
                        GenerateCustomerCsv(pending, filePath);
                        string subject = BuildCustomerSubject(pending);
                        SendCustomerFileWithGraph(pending, subject, filePath);
                        MarkCustomerFileAsSent(pending.FileId, pending.SendType);
                        WriteLog("       Customer file sent - ID : " + pending.FileId + " - Type : " + pending.SendType);
                    }
                    catch (Exception ex)
                    {
                        WriteLog("       Customer file send error - ID : " + pending.FileId + " - " + Details(ex));
                        SendTechnicalAlert(nameof(Send_Files_To_Customers), "File ID " + pending.FileId + " - " + Details(ex));
                    }
                    finally
                    {
                        try { if (File.Exists(filePath)) File.Delete(filePath); } catch { }
                    }
                }
            }
        }

        private List<PendingCustomerFile> LoadPendingCustomerFiles()
        {
            const string sql = @"
SELECT f.id_fichier, x.SendType, f.id_mail, f.nom_fichier, m.cust_nbr,
       ISNULL(c.en_tete_fichier_csv,''), ISNULL(c.dest_mail,''),
       ISNULL(c.mail_copy_to,''), ISNULL(c.subject_mail,''), ISNULL(c.body_mail,'')
FROM dbo.T_suivi_commande_end_user_fichier f
INNER JOIN dbo.T_suivi_commande_end_user_mail m ON m.id_mail=f.id_mail
INNER JOIN dbo.T_suivi_commande_end_user_client_config c ON c.cust_nbr=m.cust_nbr
CROSS APPLY
(
    SELECT 0 AS SendType
    WHERE f.a_envoyer_o_n='O' AND f.sn_imei_updated_o_n='O'
    UNION ALL
    SELECT 1 AS SendType
    WHERE f.a_envoyer_o_n='O'
      AND f.info_tracking_updated_from_tool_o_n='O'
      AND f.info_tracking_updated_o_n='O'
) x
ORDER BY f.id_fichier, x.SendType;";
            var result = new List<PendingCustomerFile>();
            using (var con = new SqlConnection(sqlGestion))
            using (var cmd = new SqlCommand(sql, con))
            {
                cmd.CommandTimeout = 300; con.Open();
                using (SqlDataReader r = cmd.ExecuteReader())
                    while (r.Read()) result.Add(new PendingCustomerFile
                    {
                        FileId = Convert.ToInt32(r[0]),
                        SendType = Convert.ToInt32(r[1]),
                        MailId = Convert.ToInt32(r[2]),
                        FileName = Convert.ToString(r[3]),
                        CustomerNumber = Convert.ToString(r[4]),
                        Header = Convert.ToString(r[5]),
                        To = Convert.ToString(r[6]),
                        Cc = Convert.ToString(r[7]),
                        Subject = Convert.ToString(r[8]),
                        Body = Convert.ToString(r[9])
                    });
            }
            return result;
        }

        private List<ImeiSource> LoadImeiSources(string sqlCon, string currentCountry)
        {
            string json = GetParameter(sqlCon, imeiConfigurationParameterGlobal);
            ImeiJsonFile cfg = JsonConvert.DeserializeObject<ImeiJsonFile>(json ?? "");
            var tables = new List<string>();
            if (cfg?.countries != null)
                tables.AddRange(cfg.countries
                    .Where(c => IsTrue(c.active) && string.Equals(c.country, currentCountry, StringComparison.OrdinalIgnoreCase))
                    .SelectMany(c => c.sharedmailboxes ?? new List<ImeiMailbox>())
                    .Where(m => IsTrue(m.active) && !string.IsNullOrWhiteSpace(m.table_name))
                    .Select(m => m.table_name));
            if (!string.IsNullOrWhiteSpace(appleImeiTableName)) tables.Add(appleImeiTableName);
            return tables.Distinct(StringComparer.OrdinalIgnoreCase).Select(BuildImeiSource).ToList();
        }

        private ImeiSource BuildImeiSource(string configuredTable)
        {
            string[] parts = (configuredTable ?? "").Split('.');
            string schema = parts.Length == 2 ? parts[0] : "dbo", table = parts.Length == 2 ? parts[1] : parts[0];
            if (!Regex.IsMatch(schema, @"^[A-Za-z_][A-Za-z0-9_]*$") || !Regex.IsMatch(table, @"^[A-Za-z_][A-Za-z0-9_]*$"))
                throw new InvalidOperationException("Invalid IMEI table name : " + configuredTable);
            var columns = new List<string>();
            using (var con = new SqlConnection(sqlGestion)) using (var cmd = new SqlCommand(@"
SELECT c.name FROM sys.columns c JOIN sys.tables t ON t.object_id=c.object_id
JOIN sys.schemas s ON s.schema_id=t.schema_id WHERE t.name=@T AND s.name=@S;", con))
            { cmd.Parameters.Add(P("@T", SqlDbType.NVarChar, table, 128)); cmd.Parameters.Add(P("@S", SqlDbType.NVarChar, schema, 128)); con.Open(); using (var r = cmd.ExecuteReader()) while (r.Read()) columns.Add(Convert.ToString(r[0])); }
            if (columns.Count == 0) throw new InvalidOperationException("IMEI table not found : " + configuredTable);
            var source = new ImeiSource { TableName = configuredTable, QualifiedTableName = "[" + schema + "]" + ".[" + table + "]", IsApple = table.Equals("IMEI_APPLE", StringComparison.OrdinalIgnoreCase) };
            source.LookupColumn = FindColumn(columns, "IMEI1", "IMEI", "IMEI Number", "IMEI Number ");
            source.SerialColumn = FindColumn(columns, "serial_number", "Serial Number", "SN");
            source.Imei2Column = FindColumn(columns, "IMEI2", "IMEIbis", "IMEI 2");
            source.EidColumn = FindColumn(columns, "eID", "EID");
            return source;
        }

        private static string FindColumn(IEnumerable<string> columns, params string[] aliases)
        {
            foreach (string alias in aliases)
            { string found = columns.FirstOrDefault(c => NormalizeName(c) == NormalizeName(alias)); if (found != null) return found; }
            return "";
        }
        private static string NormalizeName(string value) => Regex.Replace(value ?? "", "[^A-Za-z0-9]", "").ToUpperInvariant();

        private ImeiInformation FindImeiInformation(string imei1, string existingSerial)
        {
            var info = new ImeiInformation { SerialNumber = existingSerial ?? "" };
            foreach (ImeiSource source in imeiSources)
            {
                string lookupValue = !string.IsNullOrWhiteSpace(source.LookupColumn) ? imei1 : existingSerial;
                if (string.IsNullOrWhiteSpace(lookupValue) || string.IsNullOrWhiteSpace(source.LookupColumn)) continue;
                string sql = "SELECT TOP(1) " + SqlColumnOrEmpty(source.SerialColumn, "SerialNumber") + "," +
                    SqlColumnOrEmpty(source.Imei2Column, "Imei2") + "," + SqlColumnOrEmpty(source.EidColumn, "Eid") +
                    " FROM " + source.QualifiedTableName + " WHERE LTRIM(RTRIM([" + source.LookupColumn.Replace("]", "]]") + "]))=@V;";
                using (var con = new SqlConnection(sqlGestion)) using (var cmd = new SqlCommand(sql, con))
                {
                    cmd.Parameters.Add(P("@V", SqlDbType.VarChar, lookupValue, 200)); con.Open(); using (var r = cmd.ExecuteReader()) if (r.Read())
                    {
                        string serial = Convert.ToString(r[0]).Trim(), imei2 = Convert.ToString(r[1]).Trim(), eid = Convert.ToString(r[2]).Trim();
                        if (source.IsApple && !string.IsNullOrWhiteSpace(serial) && !serial.StartsWith("S", StringComparison.OrdinalIgnoreCase)) serial = "S" + serial;
                        if (string.IsNullOrWhiteSpace(info.SerialNumber)) info.SerialNumber = serial;
                        if (string.IsNullOrWhiteSpace(info.Imei2)) info.Imei2 = imei2;
                        if (string.IsNullOrWhiteSpace(info.Eid)) info.Eid = eid;
                    }
                }
                if (!string.IsNullOrWhiteSpace(info.SerialNumber) && !string.IsNullOrWhiteSpace(info.Imei2) && !string.IsNullOrWhiteSpace(info.Eid)) break;
            }
            return info;
        }
        private static string SqlColumnOrEmpty(string column, string alias) => string.IsNullOrWhiteSpace(column) ? "CAST('' AS varchar(500)) AS [" + alias + "]" : "ISNULL(CONVERT(varchar(500),[" + column.Replace("]", "]]") + "]), '') AS [" + alias + "]";

        private void GenerateCustomerCsv(PendingCustomerFile pending, string filePath)
        {
            string[] headers = SplitConfiguredCsvHeader(pending.Header);
            if (headers.Length == 0) throw new InvalidOperationException("en_tete_fichier_csv is empty for customer " + pending.CustomerNumber);
            using (var writer = new StreamWriter(filePath, false, Encoding.UTF8))
            {
                writer.WriteLine(string.Join(",", headers.Select(CsvEscape)));
                using (var con = new SqlConnection(sqlGestion)) using (var cmd = new SqlCommand(GetCustomerRowsSql(), con))
                {
                    cmd.CommandTimeout = 300; cmd.Parameters.Add("@FILE", SqlDbType.Int).Value = pending.FileId; con.Open(); using (var r = cmd.ExecuteReader()) while (r.Read())
                    { var values = BuildCustomerOutputValues(r, pending.CustomerNumber); writer.WriteLine(string.Join(",", headers.Select(h => CsvEscape(GetOutputValue(values, h))))); }
                }
            }
        }

        private static string GetCustomerRowsSql() => @"
SELECT po.end_user_order_ref, CASE WHEN @FILE IS NOT NULL THEN ISNULL(po.cust_po_transmis,po.cust_po) END order_reference,
bl.ship_date, ISNULL(trans.vat_number,ISNULL(bl.nom_transporteur,'')) carrier_vat_id,
bl.num_suivi_transporteur, bl.delivery_date, trans.url_suivi_tracking,
line.bundle, line.supplier_product_ref, line.serial_number, line.IMEI1,
line.end_user_first_name,line.end_user_last_name,line.end_user_mail,line.end_user_id,line.label_content,line.comment
FROM dbo.T_suivi_commande_end_user_cust_po_line line
JOIN dbo.T_suivi_commande_end_user_cust_po po ON po.id_cust_po=line.id_cust_po
LEFT JOIN dbo.T_suivi_commande_end_user_BL bl ON bl.id_bl=line.id_bl
LEFT JOIN dbo.T_suivi_commande_end_user_carrier_code cc ON cc.carrier_impulse LIKE CONCAT(bl.carrier_code,'%') AND bl.carrier_code IS NOT NULL
LEFT JOIN dbo.T_suivi_commande_end_user_transporteur trans ON trans.carrier_code=cc.carrier_code AND trans.nom_transporteur=bl.nom_transporteur
WHERE po.id_fichier=@FILE;";

        private Dictionary<string, string> BuildCustomerOutputValues(SqlDataReader r, string customer)
        {
            string imei1 = Convert.ToString(r[10]).Trim(), existingSerial = Convert.ToString(r[9]).Trim();
            ImeiInformation imei = FindImeiInformation(imei1, existingSerial);
            string tracking = Convert.ToString(r[4]).Trim(), url = Convert.ToString(r[6]).Trim();
            string pod = string.IsNullOrWhiteSpace(tracking) ? "" : url.Replace("[num_suivi]", tracking).Replace("[track_id]", tracking);
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["Order_Number"] = Convert.ToString(r[0]),
                ["Order_Reference"] = Convert.ToString(r[1]),
                ["Dispatch_Date"] = FormatCsvDate(r[2]),
                ["Carrier_Vat_ID"] = Convert.ToString(r[3]),
                ["Delivery_Tracking_Number"] = tracking,
                ["Delivery_Date"] = FormatCsvDate(r[5]),
                ["POD_ID"] = pod,
                ["Product_Type"] = Convert.ToString(r[7]),
                ["Product_Code"] = Convert.ToString(r[8]),
                ["Bundle"] = Convert.ToString(r[7]),
                ["Serial_Number"] = imei.SerialNumber,
                ["IMEI"] = imei1,
                ["IMEI 2"] = imei.Imei2,
                ["IMEI2"] = imei.Imei2,
                ["eID"] = imei.Eid,
                ["EndUser_First_Name"] = Convert.ToString(r[11]),
                ["EndUser_First_Last_Name"] = Convert.ToString(r[12]),
                ["EndUser_Last_Name"] = Convert.ToString(r[12]),
                ["EndUser_eMail_Address"] = Convert.ToString(r[13]),
                ["UserID"] = Convert.ToString(r[14]),
                ["Label_Content"] = Convert.ToString(r[15]),
                ["Comment"] = Convert.ToString(r[16])
            };
        }
        private static string GetOutputValue(Dictionary<string, string> values, string header)
        { string key = values.Keys.FirstOrDefault(k => NormalizeName(k) == NormalizeName(header)); return key == null ? "" : values[key] ?? ""; }
        private static string FormatCsvDate(object value) => value == DBNull.Value ? "" : Convert.ToDateTime(value).ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
        private static string CsvEscape(string value)
        { value = value ?? ""; return value.IndexOfAny(new[] { ',', '"', '\r', '\n' }) >= 0 ? "\"" + value.Replace("\"", "\"\"") + "\"" : value; }

        private string BuildCustomerSubject(PendingCustomerFile p)
        {
            string baseName = Path.GetFileNameWithoutExtension(p.FileName);
            if (p.SendType == 1) return (p.Subject + " " + baseName).Trim();
            string po = ScalarString(sqlGestion, "SELECT TOP(1) ISNULL(cust_po,'') FROM dbo.T_suivi_commande_end_user_cust_po WHERE id_fichier=@ID;", new SqlParameter("@ID", SqlDbType.Int) { Value = p.FileId });
            if (string.IsNullOrWhiteSpace(po)) throw new InvalidOperationException("Cust PO missing for file ID " + p.FileId);
            return "AIRBUS - IMEI - INGRAM > " + baseName + " / " + po;
        }

        private void SendCustomerFileWithGraph(PendingCustomerFile p, string subject, string filePath)
        {
            string to = p.SendType == 1 ? p.To : p.Cc, cc = p.SendType == 1 ? p.Cc : "";
            List<Recipient> recipients = Recipients(to); if (recipients.Count == 0) throw new InvalidOperationException("No valid customer recipient for file ID " + p.FileId);
            var sender = new Recipient { EmailAddress = new EmailAddress { Address = econocomSendAs, Name = "Ingram Micro" } };
            var message = new Message { Subject = subject, From = sender, Sender = sender, Body = new ItemBody { ContentType = BodyType.Html, Content = p.Body ?? "" }, ToRecipients = recipients, CcRecipients = Recipients(cc), Attachments = new List<Attachment> { new FileAttachment { OdataType = "#microsoft.graph.fileAttachment", Name = Path.GetFileName(p.FileName), ContentType = "text/csv", ContentBytes = File.ReadAllBytes(filePath) } } };
            Retry(
                () =>
                {
                    graph.Users[econocomSendAs]
                        .SendMail
                        .PostAsync(
                            new Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody
                            {
                                Message = message,
                                SaveToSentItems = true
                            })
                        .GetAwaiter()
                        .GetResult();
                    return true;
                },
                "Send customer file with Graph - File ID : " + p.FileId);
        }
        private void MarkCustomerFileAsSent(int fileId, int sendType)
        {
            string sql = sendType == 0 ? "UPDATE dbo.T_suivi_commande_end_user_fichier SET a_envoyer_o_n='N',sn_imei_updated_o_n='T' WHERE id_fichier=@ID;" : "UPDATE dbo.T_suivi_commande_end_user_fichier SET a_envoyer_o_n='N',info_tracking_updated_o_n='T',info_tracking_updated_from_tool_o_n='T' WHERE id_fichier=@ID;";
            ExecuteNonQuery(sqlGestion, sql, new SqlParameter("@ID", SqlDbType.Int) { Value = fileId });
        }

        private void ProcessMailbox(SharedMailbox mailbox, DateTime filterDate)
        {
            MailFolder inbox = GetFolder(mailbox.Address, "inbox", folderIn);
            MailFolder archives = GetFolder(mailbox.Address, inbox.Id, folderArchive);
            MailFolder duplicates = GetFolder(mailbox.Address, inbox.Id, folderDuplicate);
            MailFolder errorsFolder = GetFolder(mailbox.Address, inbox.Id, folderError);
            List<Message> messages = GetMessages(mailbox.Address, inbox.Id, filterDate);
            int integrated = 0, duplicateCount = 0, waiting = 0, errors = 0;
            WriteLog("   Processing mailbox Econocom Airbus : " + mailbox.Address + " - Found : " + messages.Count);

            foreach (Message summary in messages)
            {
                Message email = null; List<CsvFileValidation> filesToDelete = null;
                try
                {
                    ProcessingState state = GetState(mailbox.Address, summary.Id);
                    if (state != null && state.Status == "S")
                    {
                        string finalizedFolderId =
                            string.Equals(state.DestinationFolder, folderDuplicate, StringComparison.OrdinalIgnoreCase)
                                ? duplicates.Id
                                : string.Equals(state.DestinationFolder, folderError, StringComparison.OrdinalIgnoreCase)
                                    ? errorsFolder.Id
                                    : archives.Id;

                        FinalizeGraph(
                            mailbox.Address,
                            summary.Id,
                            finalizedFolderId);
                        continue;
                    }
                    email = GetMessage(mailbox.Address, summary.Id);
                    ClientImportConfiguration clientConfiguration =
                        GetClientImportConfiguration(
                            GetSender(email));

                    string customer =
                        clientConfiguration?.CustomerNumber ?? "";

                    CsvValidationResult validation =
                        ValidateCsvAttachments(
                            mailbox.Address,
                            email,
                            clientConfiguration);
                    filesToDelete = validation.Files;

                    if (!validation.IsValid)
                    {
                        bool alertAlreadySent =
                            SaveValidationError(
                                mailbox.Address,
                                email,
                                validation);

                        if (!alertAlreadySent)
                        {
                            SendValidationAlert(
                                mailbox.Address,
                                email,
                                validation);

                            MarkAlertSent(
                                mailbox.Address,
                                email.Id);
                        }

                        if (validation.ErrorType.Equals(
                                "CSV_STRUCTURE",
                                StringComparison.OrdinalIgnoreCase))
                        {
                            SetRetryableErrorStatus(
                                mailbox.Address,
                                email.Id,
                                folderError,
                                validation.ErrorType,
                                validation.Message);

                            MoveRetryableError(
                                mailbox.Address,
                                email.Id,
                                errorsFolder.Id);

                            errors++;

                            WriteLog(
                                "       Invalid CSV moved to error folder" +
                                " - Subject : " +
                                (email.Subject ?? "<no subject>") +
                                " - Details : " +
                                validation.Message);
                        }
                        else if (validation.ErrorType.Equals(
                                     "MISSING_BUNDLE",
                                     StringComparison.OrdinalIgnoreCase) &&
                                 IsBundleUpdateDelayExpired(
                                     mailbox.Address,
                                     email.Id))
                        {
                            SendBundleDelayExpiredAlert(
                                mailbox.Address,
                                email,
                                validation);

                            SetRetryableErrorStatus(
                                mailbox.Address,
                                email.Id,
                                folderError,
                                validation.ErrorType,
                                validation.Message);

                            MoveRetryableError(
                                mailbox.Address,
                                email.Id,
                                errorsFolder.Id);

                            errors++;

                            WriteLog(
                                "       Missing bundle delay expired; " +
                                "mail moved to error folder" +
                                " - Delay : " +
                                bundleUpdateDelayDays +
                                " day(s)" +
                                " - Subject : " +
                                (email.Subject ?? "<no subject>") +
                                " - Details : " +
                                validation.Message);
                        }
                        else
                        {
                            waiting++;

                            WriteLog(
                                "       Mail left pending" +
                                " - Type : " +
                                validation.ErrorType +
                                " - Delay : " +
                                bundleUpdateDelayDays +
                                " day(s)" +
                                " - Details : " +
                                validation.Message);
                        }

                        continue;
                    }

                    if (validation.UnexpectedColumns.Count > 0)
                    {
                        SendCsvAdditionalColumnsAlert(
                            mailbox.Address,
                            email,
                            validation);
                    }

                    if (validation.IsDuplicate)
                    {
                        int duplicateMailId = InsertMailAndTrack(
                            mailbox.Address,
                            email,
                            GetMime(mailbox.Address, email.Id),
                            customer);

                        SendDuplicateAlert(
                            mailbox.Address,
                            email,
                            validation.Message);

                        SetSuccess(
                            mailbox.Address,
                            email.Id,
                            duplicateMailId,
                            folderDuplicate);

                        MarkProcessed(
                            mailbox.Address,
                            email.Id);

                        Move(
                            mailbox.Address,
                            email.Id,
                            duplicates.Id);

                        duplicateCount++;
                        continue;
                    }

                    int mailId = InsertMailAndTrack(mailbox.Address, email, GetMime(mailbox.Address, email.Id), customer);
                    ImportValidatedFiles(validation.Files, mailId, customer);
                    SetSuccess(mailbox.Address, email.Id, mailId, folderArchive);
                    MarkProcessed(mailbox.Address, email.Id); Move(mailbox.Address, email.Id, archives.Id);
                    integrated++;
                }
                catch (Exception ex)
                {
                    errors++; UpsertTechnicalError(mailbox.Address, email ?? summary, ex);
                    WriteLog("       Mail processing error - " + (summary.Subject ?? "<no subject>") + " - " + Details(ex));
                    SendTechnicalAlert("Read_Email_with_Graph", Details(ex));
                }
                finally { DeleteTempFiles(filesToDelete); }
            }
            WriteLog("   Mailbox summary - Imported : " + integrated + " - Duplicates : " + duplicateCount + " - Waiting : " + waiting + " - Errors : " + errors);
        }

        private CsvValidationResult ValidateCsvAttachments(
            string mailbox,
            Message email,
            ClientImportConfiguration clientConfiguration)
        {
            var result = new CsvValidationResult();
            string customer =
                clientConfiguration?.CustomerNumber ?? "";

            if (clientConfiguration == null ||
                string.IsNullOrWhiteSpace(customer))
            {
                result.ErrorType = "CSV_STRUCTURE";
                result.Message =
                    "Aucune configuration client n'a été trouvée pour " +
                    "l'expéditeur du message";
                return result;
            }

            if (string.IsNullOrWhiteSpace(
                    clientConfiguration.ExpectedCsvHeader))
            {
                result.ErrorType = "CSV_STRUCTURE";
                result.Message =
                    "Le champ en_tete_fichier_csv est vide pour le client " +
                    customer;
                return result;
            }

            string[] expectedHeaderFields =
                SplitConfiguredCsvHeader(
                    clientConfiguration.ExpectedCsvHeader);

            int expectedColumnCount =
                expectedHeaderFields.Length;

            if (expectedColumnCount == 0)
            {
                result.ErrorType = "CSV_STRUCTURE";
                result.Message =
                    "Le champ en_tete_fichier_csv ne contient aucune " +
                    "colonne exploitable pour le client " + customer;
                return result;
            }

            List<FileAttachment> attachments = GetCsvAttachments(mailbox, email.Id);
            if (attachments.Count == 0)
            {
                result.ErrorType = "CSV_STRUCTURE"; result.Message = "Aucune pièce jointe CSV trouvée"; return result;
            }
            var allBundles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            string firstOrderReference = "";
            foreach (FileAttachment attachment in attachments)
            {
                string safeName = Path.GetFileName(attachment.Name ?? "attachment.csv");
                string temp = Path.Combine(tempFolder, Guid.NewGuid().ToString("N") + "_" + safeName);
                File.WriteAllBytes(temp, attachment.ContentBytes ?? Array.Empty<byte>());
                var file = new CsvFileValidation { Name = safeName, TempPath = temp };
                result.Files.Add(file);
                try
                {
                    using (var parser = new TextFieldParser(temp, Encoding.Default))
                    {
                        parser.TextFieldType = FieldType.Delimited; parser.SetDelimiters(","); parser.HasFieldsEnclosedInQuotes = true;
                        bool header = true;
                        while (!parser.EndOfData)
                        {
                            string[] fields = parser.ReadFields() ?? Array.Empty<string>();
                            if (header)
                            {
                                file.ColumnIndexes =
                                    BuildCsvColumnIndexes(fields);

                                List<string> unexpectedColumns =
                                    GetUnexpectedCsvColumns(
                                        fields,
                                        expectedHeaderFields);

                                foreach (string unexpectedColumn in unexpectedColumns)
                                {
                                    if (!result.UnexpectedColumns.Contains(
                                            unexpectedColumn,
                                            StringComparer.OrdinalIgnoreCase))
                                    {
                                        result.UnexpectedColumns.Add(unexpectedColumn);
                                    }
                                }

                                List<string> missingBusinessColumns =
                                    GetMissingBusinessColumns(
                                        file.ColumnIndexes);

                                if (missingBusinessColumns.Count > 0)
                                {
                                    result.ErrorType = "CSV_STRUCTURE";
                                    result.Message =
                                        "Le fichier " + safeName +
                                        " ne contient pas toutes les colonnes " +
                                        "métier requises. Colonne(s) absente(s) : " +
                                        string.Join(", ", missingBusinessColumns) +
                                        ". En-tête configuré : " +
                                        clientConfiguration.ExpectedCsvHeader;
                                    return result;
                                }

                                header = false;
                                continue;
                            }

                            int minimumColumnCount =
                                file.ColumnIndexes.Count == 0
                                    ? 0
                                    : file.ColumnIndexes.Values.Max() + 1;

                            if (fields.Length < minimumColumnCount)
                            {
                                result.ErrorType = "CSV_STRUCTURE";
                                result.Message =
                                    "Le fichier " + safeName +
                                    " contient une ligne incomplète avec " +
                                    fields.Length +
                                    " colonne(s). Le nombre minimum requis " +
                                    "pour les colonnes métier détectées est " +
                                    minimumColumnCount + ".";
                                return result;
                            }
                            if (string.IsNullOrWhiteSpace(firstOrderReference))
                            {
                                firstOrderReference =
                                    GetCsvValue(
                                        fields,
                                        file.ColumnIndexes,
                                        "order_number");
                            }

                            file.Rows.Add(fields);

                            string bundle =
                                GetCsvValue(
                                    fields,
                                    file.ColumnIndexes,
                                    "product_type");

                            if (!string.IsNullOrWhiteSpace(bundle))
                            {
                                allBundles.Add(bundle.Trim());
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    result.ErrorType = "CSV_STRUCTURE"; result.Message = "Le fichier " + safeName + " est illisible : " + ex.Message; return result;
                }
                if (file.Rows.Count == 0)
                {
                    result.ErrorType = "CSV_STRUCTURE"; result.Message = "Le fichier " + safeName + " ne contient aucune ligne de commande"; return result;
                }
            }

            foreach (string bundle in allBundles)
                if (string.IsNullOrWhiteSpace(GetSku(bundle))) result.MissingBundles.Add(bundle);
            if (result.MissingBundles.Count > 0)
            {
                result.ErrorType = "MISSING_BUNDLE";
                result.Message = "Bundle(s) à créer : " + string.Join(", ", result.MissingBundles.OrderBy(v => v));
                return result;
            }
            if (!string.IsNullOrWhiteSpace(firstOrderReference) && IsDuplicate(customer, firstOrderReference))
            {
                result.IsValid = true; result.IsDuplicate = true; result.Message = "Commande déjà traitée : " + firstOrderReference; return result;
            }
            result.IsValid = true; return result;
        }

        private bool SaveValidationError(string mailbox, Message email, CsvValidationResult validation)
        {
            ProcessingState old = GetState(mailbox, email.Id);
            bool alreadySent = old?.AlertSentAtUtc.HasValue == true && string.Equals(old.ErrorType, validation.ErrorType, StringComparison.OrdinalIgnoreCase);
            ExecuteNonQuery(sqlGestion, @"
MERGE dbo.T_suivi_commande_end_user_MailProcessing t
USING(SELECT @MAILBOX MailboxAddress,@GRAPH GraphMessageId)s
ON t.MailboxAddress=s.MailboxAddress AND t.GraphMessageId=s.GraphMessageId
WHEN MATCHED THEN UPDATE SET InternetMessageId=@INTERNET,MessageSubject=@SUBJECT,ReceivedDateTimeUtc=@RECEIVED,
ProcessingStatus='E',ErrorType=@TYPE,MissingBundles=@BUNDLES,ErrorMessage=@ERROR,
AlertSentAtUtc=CASE WHEN ISNULL(t.ErrorType,'')<>@TYPE OR ISNULL(t.MissingBundles,'')<>@BUNDLES THEN NULL ELSE t.AlertSentAtUtc END,
AttemptCount=t.AttemptCount+1,LastAttemptAtUtc=SYSUTCDATETIME()
WHEN NOT MATCHED THEN INSERT(MailboxAddress,GraphMessageId,InternetMessageId,MessageSubject,ReceivedDateTimeUtc,
ProcessingStatus,ErrorType,MissingBundles,ErrorMessage,AttemptCount)
VALUES(@MAILBOX,@GRAPH,@INTERNET,@SUBJECT,@RECEIVED,'E',@TYPE,@BUNDLES,@ERROR,1);",
                P("@MAILBOX", SqlDbType.NVarChar, mailbox, 320), P("@GRAPH", SqlDbType.NVarChar, email.Id, 500),
                P("@INTERNET", SqlDbType.NVarChar, email.InternetMessageId ?? "", 500), P("@SUBJECT", SqlDbType.NVarChar, Truncate(email.Subject, 1000), 1000),
                new SqlParameter("@RECEIVED", SqlDbType.DateTime2) { Value = email.ReceivedDateTime?.UtcDateTime ?? DateTime.UtcNow },
                P("@TYPE", SqlDbType.VarChar, validation.ErrorType, 50),
                P("@BUNDLES", SqlDbType.NVarChar, Truncate(string.Join(";", validation.MissingBundles.OrderBy(v => v)), 2000), 2000),
                P("@ERROR", SqlDbType.NVarChar, Truncate(validation.Message, 2000), 2000));
            return alreadySent;
        }

        private static List<string> GetUnexpectedCsvColumns(
            string[] receivedHeaderFields,
            string[] configuredHeaderFields)
        {
            var knownColumns = new HashSet<string>(
                CsvColumnAliases.Values
                    .SelectMany(aliases => aliases)
                    .Concat(
                        configuredHeaderFields ??
                        Array.Empty<string>())
                    .Select(NormalizeCsvColumnName)
                    .Where(value =>
                        !string.IsNullOrWhiteSpace(value)),
                StringComparer.OrdinalIgnoreCase);

            return (receivedHeaderFields ??
                    Array.Empty<string>())
                .Select(value =>
                    (value ?? "").Trim())
                .Where(value =>
                    !string.IsNullOrWhiteSpace(value))
                .Where(value =>
                    !knownColumns.Contains(
                        NormalizeCsvColumnName(value)))
                .Distinct(
                    StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private void SendCsvAdditionalColumnsAlert(
            string mailbox,
            Message email,
            CsvValidationResult validation)
        {
            string text =
                "Commande Econocom/Airbus importée malgré la présence " +
                "de nouvelle(s) colonne(s) non utilisée(s). " +
                "Boîte : " + mailbox +
                ". Sujet : " + (email.Subject ?? "") +
                ". Colonne(s) supplémentaire(s) : " +
                string.Join(", ", validation.UnexpectedColumns) +
                ". Aucune donnée de ces colonnes n'a été importée.";

            SendAlertToRecipients(
                "Nouvelle colonne CSV détectée",
                text,
                MergeRecipients(
                    businessAlertRecipients,
                    alertRecipients));
        }

        private static string MergeRecipients(params string[] recipientLists)
        {
            return string.Join(
                ";",
                (recipientLists ?? Array.Empty<string>())
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .SelectMany(value => value.Split(
                        new[] { ';', ',' },
                        StringSplitOptions.RemoveEmptyEntries))
                    .Select(value => value.Trim())
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Distinct(StringComparer.OrdinalIgnoreCase));
        }

        private void SendValidationAlert(string mailbox, Message email, CsvValidationResult validation)
        {
            string text = validation.ErrorType == "MISSING_BUNDLE"
                ? "Commande Econocom/Airbus laissée en attente. Boîte : " + mailbox + ". Sujet : " + (email.Subject ?? "") +
                  ". Bundle(s) à créer dans T_suivi_commande_end_user_ref_sku : " + string.Join(", ", validation.MissingBundles) +
                  ". La commande sera reprise automatiquement après création."
                : "Commande Econocom/Airbus laissée en attente. Boîte : " + mailbox + ". Sujet : " + (email.Subject ?? "") + ". " + validation.Message;
            SendAlertToRecipients(
                "Validation CSV",
                text,
                MergeRecipients(
                    businessAlertRecipients,
                    alertRecipients));
        }

        private bool IsBundleUpdateDelayExpired(
            string mailbox,
            string graphMessageId)
        {
            ProcessingState state =
                GetState(mailbox, graphMessageId);

            if (state?.FirstSeenAtUtc == null)
            {
                return false;
            }

            DateTime expirationDateUtc =
                state.FirstSeenAtUtc.Value.AddDays(
                    bundleUpdateDelayDays);

            return DateTime.UtcNow >= expirationDateUtc;
        }

        private void SendBundleDelayExpiredAlert(
            string mailbox,
            Message email,
            CsvValidationResult validation)
        {
            string text =
                "Commande Econocom/Airbus déplacée dans le dossier " +
                folderError +
                " car le ou les bundles n'ont pas été créés dans le " +
                "délai configuré de " +
                bundleUpdateDelayDays +
                " jour(s). Boîte : " + mailbox +
                ". Sujet : " + (email.Subject ?? "") +
                ". Bundle(s) toujours absent(s) : " +
                string.Join(", ", validation.MissingBundles) +
                ".";

            SendBusinessAlert(
                "Délai de création du bundle dépassé",
                text);
        }

        private void SendDuplicateAlert(
            string mailbox,
            Message email,
            string duplicateDetails)
        {
            string text =
                "Commande Econocom/Airbus détectée en doublon. " +
                "Boîte : " + mailbox + ". " +
                "Sujet : " + (email.Subject ?? "") + ". " +
                duplicateDetails + ". " +
                "Le mail a été classé dans le dossier " +
                folderDuplicate + ".";

            SendBusinessAlert(
                "Commande en doublon",
                text);
        }

        private void SetRetryableErrorStatus(
            string mailbox,
            string graphId,
            string destinationFolder,
            string errorType,
            string errorMessage)
        {
            ExecuteNonQuery(
                sqlGestion,
                @"UPDATE dbo.T_suivi_commande_end_user_MailProcessing
SET ProcessingStatus='E',
    DestinationFolder=@FOLDER,
    ErrorType=@TYPE,
    ErrorMessage=@ERROR,
    ProcessedAtUtc=NULL,
    LastAttemptAtUtc=SYSUTCDATETIME()
WHERE MailboxAddress=@MAILBOX
  AND GraphMessageId=@GRAPH;",
                P("@FOLDER", SqlDbType.NVarChar, destinationFolder, 100),
                P("@TYPE", SqlDbType.VarChar, errorType, 50),
                P("@ERROR", SqlDbType.NVarChar, Truncate(errorMessage, 2000), 2000),
                P("@MAILBOX", SqlDbType.NVarChar, mailbox, 320),
                P("@GRAPH", SqlDbType.NVarChar, graphId, 500));
        }


        private void MarkAlertSent(string mailbox, string graphId)
        {
            ExecuteNonQuery(sqlGestion, @"UPDATE dbo.T_suivi_commande_end_user_MailProcessing SET AlertSentAtUtc=SYSUTCDATETIME()
WHERE MailboxAddress=@MAILBOX AND GraphMessageId=@GRAPH;",
                P("@MAILBOX", SqlDbType.NVarChar, mailbox, 320), P("@GRAPH", SqlDbType.NVarChar, graphId, 500));
        }

        private void ImportValidatedFiles(List<CsvFileValidation> files, int mailId, string customer)
        {
            foreach (CsvFileValidation file in files)
            {
                int fileId = ScalarInt(sqlGestion, @"INSERT INTO dbo.T_suivi_commande_end_user_fichier(id_mail,nom_fichier)
VALUES(@MAIL,@NAME); SELECT CONVERT(int,SCOPE_IDENTITY());",
                    new SqlParameter("@MAIL", SqlDbType.Int) { Value = mailId }, P("@NAME", SqlDbType.VarChar, file.Name, 500));
                var poIds = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                foreach (string[] row in file.Rows)
                {
                    string transmittedPo =
                        GetCsvValue(
                            row,
                            file.ColumnIndexes,
                            "order_reference");

                    string orderNumber =
                        GetCsvValue(
                            row,
                            file.ColumnIndexes,
                            "order_number");

                    if (!poIds.TryGetValue(
                            transmittedPo,
                            out int poId))
                    {
                        poId = InsertPo(
                            fileId,
                            customer,
                            transmittedPo,
                            orderNumber);

                        poIds.Add(
                            transmittedPo,
                            poId);
                    }

                    InsertLine(
                        poId,
                        row,
                        file.ColumnIndexes);
                }
            }
        }

        private int InsertMailAndTrack(string mailbox, Message email, byte[] mime, string customer)
        {
            using (var con = new SqlConnection(sqlGestion))
            {
                con.Open(); using (var tx = con.BeginTransaction())
                {
                    try
                    {
                        using (var cmd = new SqlCommand("dbo.USP_Insert_mail_suivi_cde_end_user", con, tx))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.Add("@date_mail", SqlDbType.DateTime).Value = email.ReceivedDateTime?.LocalDateTime ?? DateTime.Now;
                            cmd.Parameters.Add("@fichier", SqlDbType.Image).Value = mime;
                            cmd.Parameters.Add("@expediteur", SqlDbType.VarChar).Value = GetSender(email);
                            cmd.Parameters.Add("@extension", SqlDbType.VarChar).Value = ".eml";
                            cmd.Parameters.Add("@cust_nbr", SqlDbType.VarChar).Value = customer ?? "";
                            cmd.Parameters.Add("@date_trt", SqlDbType.DateTime).Value = DateTime.Now; cmd.ExecuteNonQuery();
                        }
                        int id; using (var cmd = new SqlCommand("SELECT MAX(id_mail) FROM dbo.T_suivi_commande_end_user_mail;", con, tx)) id = Convert.ToInt32(cmd.ExecuteScalar());
                        using (var cmd = new SqlCommand(@"MERGE dbo.T_suivi_commande_end_user_MailProcessing t
USING(SELECT @MAILBOX MailboxAddress,@GRAPH GraphMessageId)s
ON t.MailboxAddress=s.MailboxAddress AND t.GraphMessageId=s.GraphMessageId
WHEN MATCHED THEN UPDATE SET InternetMessageId=@INTERNET,MessageSubject=@SUBJECT,
ReceivedDateTimeUtc=@RECEIVED,MailDatabaseId=@ID,ProcessingStatus='P',ErrorType=NULL,
MissingBundles=NULL,ErrorMessage=NULL,AlertSentAtUtc=NULL,
AttemptCount=t.AttemptCount+1,LastAttemptAtUtc=SYSUTCDATETIME()
WHEN NOT MATCHED THEN INSERT(MailboxAddress,GraphMessageId,InternetMessageId,
MessageSubject,ReceivedDateTimeUtc,MailDatabaseId,ProcessingStatus,AttemptCount)
VALUES(@MAILBOX,@GRAPH,@INTERNET,@SUBJECT,@RECEIVED,@ID,'P',1);", con, tx))
                        {
                            cmd.Parameters.Add("@ID", SqlDbType.Int).Value = id;
                            cmd.Parameters.Add(P("@MAILBOX", SqlDbType.NVarChar, mailbox, 320));
                            cmd.Parameters.Add(P("@GRAPH", SqlDbType.NVarChar, email.Id, 500));
                            cmd.Parameters.Add(P("@INTERNET", SqlDbType.NVarChar, email.InternetMessageId ?? "", 500));
                            cmd.Parameters.Add(P("@SUBJECT", SqlDbType.NVarChar, Truncate(email.Subject, 1000), 1000));
                            cmd.Parameters.Add("@RECEIVED", SqlDbType.DateTime2).Value =
                                email.ReceivedDateTime?.UtcDateTime ?? DateTime.UtcNow;
                            cmd.ExecuteNonQuery();
                        }
                        tx.Commit(); return id;
                    }
                    catch { tx.Rollback(); throw; }
                }
            }
        }

        private List<SharedMailbox> LoadMailboxes()
        {
            var list = new List<SharedMailbox>();
            using (var con = new SqlConnection(sqlGestion)) using (var cmd = new SqlCommand(@"SELECT id_mailboxe,nom_mailboxe,mailboxe,
dt_heure_filtre,raffraichissement_min,date_dernier_raf,actif FROM dbo.T_SharedMailboxes
WHERE typologie=@T ORDER BY ordre,id_mailboxe;", con))
            {
                cmd.Parameters.Add(P("@T", SqlDbType.VarChar, mailboxTypology, 100)); con.Open(); using (SqlDataReader r = cmd.ExecuteReader())
                    while (r.Read()) list.Add(new SharedMailbox
                    {
                        Id = Convert.ToInt32(r[0]),
                        Name = Convert.ToString(r[1]),
                        Address = Convert.ToString(r[2]),
                        FilterDate = r[3] == DBNull.Value ? (DateTime?)null : Convert.ToDateTime(r[3]),
                        RefreshMinutes = r[4] == DBNull.Value ? 0 : Convert.ToInt32(r[4]),
                        LastRefresh = r[5] == DBNull.Value ? (DateTime?)null : Convert.ToDateTime(r[5]),
                        Active = r[6] != DBNull.Value && Convert.ToInt32(r[6]) == 1
                    });
            }
            return list;
        }

        private ProcessingState GetState(string mailbox, string graphId)
        {
            using (var con = new SqlConnection(sqlGestion)) using (var cmd = new SqlCommand(@"SELECT ProcessingStatus,MailDatabaseId,
ISNULL(DestinationFolder,''),ISNULL(ErrorType,''),AlertSentAtUtc,FirstSeenAtUtc FROM dbo.T_suivi_commande_end_user_MailProcessing
WHERE MailboxAddress=@M AND GraphMessageId=@G;", con))
            {
                cmd.Parameters.Add(P("@M", SqlDbType.NVarChar, mailbox, 320)); cmd.Parameters.Add(P("@G", SqlDbType.NVarChar, graphId, 500)); con.Open();
                using (SqlDataReader r = cmd.ExecuteReader())
                {
                    if (!r.Read()) return null; return new ProcessingState
                    {
                        Status = Convert.ToString(r[0]),
                        MailId = r[1] == DBNull.Value ? (int?)null : Convert.ToInt32(r[1]),
                        DestinationFolder = Convert.ToString(r[2]),
                        ErrorType = Convert.ToString(r[3]),
                        AlertSentAtUtc = r[4] == DBNull.Value ? (DateTime?)null : Convert.ToDateTime(r[4]),
                        FirstSeenAtUtc = r[5] == DBNull.Value ? (DateTime?)null : Convert.ToDateTime(r[5])
                    };
                }
            }
        }

        private static Dictionary<string, int> BuildCsvColumnIndexes(
            string[] headerFields)
        {
            var normalizedHeaders =
                (headerFields ?? Array.Empty<string>())
                    .Select((value, index) => new
                    {
                        Name = NormalizeCsvColumnName(value),
                        Index = index
                    })
                    .Where(item =>
                        !string.IsNullOrWhiteSpace(item.Name))
                    .GroupBy(item =>
                        item.Name,
                        StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(
                        group => group.Key,
                        group => group.First().Index,
                        StringComparer.OrdinalIgnoreCase);

            var result =
                new Dictionary<string, int>(
                    StringComparer.OrdinalIgnoreCase);

            foreach (KeyValuePair<string, string[]> mapping in
                CsvColumnAliases)
            {
                foreach (string alias in mapping.Value)
                {
                    string normalizedAlias =
                        NormalizeCsvColumnName(alias);

                    if (normalizedHeaders.TryGetValue(
                            normalizedAlias,
                            out int index))
                    {
                        result[mapping.Key] = index;
                        break;
                    }
                }
            }

            return result;
        }

        private static List<string> GetMissingBusinessColumns(
            Dictionary<string, int> columnIndexes)
        {
            return CsvColumnAliases.Keys
                .Where(key =>
                    columnIndexes == null ||
                    !columnIndexes.ContainsKey(key))
                .Select(key =>
                    CsvColumnAliases[key][0])
                .ToList();
        }

        private static string GetCsvValue(
            string[] row,
            Dictionary<string, int> columnIndexes,
            string businessColumn)
        {
            if (row == null ||
                columnIndexes == null ||
                !columnIndexes.TryGetValue(
                    businessColumn,
                    out int index) ||
                index < 0 ||
                index >= row.Length)
            {
                return "";
            }

            return (row[index] ?? "").Trim();
        }

        private static string NormalizeCsvColumnName(string value)
        {
            return Regex.Replace(
                NormalizeCsvHeaderValue(value),
                "[^A-Za-z0-9]",
                "")
                .ToUpperInvariant();
        }

        private ClientImportConfiguration GetClientImportConfiguration(
            string sender)
        {
            string normalizedSender =
                (sender ?? "").Trim().ToLowerInvariant();

            string domain = normalizedSender.Contains("@")
                ? normalizedSender.Split('@')[1]
                : "";

            const string sql = @"
SELECT TOP (1)
       ISNULL(cust_nbr,''),
       ISNULL(en_tete_fichier_csv,'')
FROM dbo.T_suivi_commande_end_user_client_config
WHERE CHARINDEX(@MAIL,LOWER(ISNULL(mail_from,''))) > 0
   OR CHARINDEX(@DOMAIN,LOWER(ISNULL(mail_from_domaine,''))) > 0
ORDER BY CASE
             WHEN CHARINDEX(@MAIL,LOWER(ISNULL(mail_from,''))) > 0
                 THEN 0
             ELSE 1
         END,
         cust_nbr;";

            using (var connection = new SqlConnection(sqlGestion))
            using (var command = new SqlCommand(sql, connection))
            {
                command.CommandTimeout = 300;
                command.Parameters.Add(
                    P(
                        "@MAIL",
                        SqlDbType.VarChar,
                        normalizedSender,
                        500));
                command.Parameters.Add(
                    P(
                        "@DOMAIN",
                        SqlDbType.VarChar,
                        domain,
                        255));

                connection.Open();

                using (SqlDataReader reader = command.ExecuteReader())
                {
                    if (!reader.Read())
                    {
                        return null;
                    }

                    return new ClientImportConfiguration
                    {
                        CustomerNumber =
                            Convert.ToString(reader[0]).Trim(),
                        ExpectedCsvHeader =
                            Convert.ToString(reader[1]).Trim()
                    };
                }
            }
        }


        private static string[] SplitConfiguredCsvHeader(
            string configuredHeader)
        {
            return (configuredHeader ?? "")
                .Split(
                    new[] { ',' },
                    StringSplitOptions.None)
                .Select(field =>
                    NormalizeCsvHeaderValue(field))
                .Where(field =>
                    !string.IsNullOrWhiteSpace(field))
                .ToArray();
        }

        private static string NormalizeCsvHeaderValue(string value)
        {
            return (value ?? "")
                .Trim()
                .Trim('﻿')
                .Replace("\r", "")
                .Replace("\n", "");
        }

        private bool IsDuplicate(string customer, string orderRef) => ScalarInt(sqlGestion, @"SELECT COUNT(*) FROM dbo.T_suivi_commande_end_user_cust_po p
JOIN dbo.T_suivi_commande_end_user_fichier f ON p.id_fichier=f.id_fichier JOIN dbo.T_suivi_commande_end_user_mail m ON f.id_mail=m.id_mail
WHERE m.cust_nbr=@C AND m.date_mail>DATEADD(month,-1,GETDATE()) AND p.end_user_order_ref=@R AND p.cde_annule<>1;",
            P("@C", SqlDbType.VarChar, customer, 100), P("@R", SqlDbType.VarChar, orderRef, 500)) > 0;
        private string GetSku(string bundle) => ScalarString(sqlGestion,
            "SELECT TOP(1) ISNULL(sku,'') FROM dbo.T_suivi_commande_end_user_ref_sku WHERE UPPER(bundle)=UPPER(@B);",
            P("@B", SqlDbType.VarChar, bundle.Trim(), 500));
        private int InsertPo(int fileId, string customer, string po, string orderRef)
        {
            if (customer == "21203530") return ScalarInt(sqlGestion, @"INSERT INTO dbo.T_suivi_commande_end_user_cust_po
(id_fichier,cust_po_transmis,end_user_order_ref,ouvert) VALUES(@F,@P,@R,1); SELECT CONVERT(int,SCOPE_IDENTITY());",
                new SqlParameter("@F", SqlDbType.Int) { Value = fileId }, P("@P", SqlDbType.VarChar, po, 500), P("@R", SqlDbType.VarChar, orderRef, 500));
            return ScalarInt(sqlGestion, @"INSERT INTO dbo.T_suivi_commande_end_user_cust_po(id_fichier,cust_po,ouvert)
VALUES(@F,@P,1); SELECT CONVERT(int,SCOPE_IDENTITY());", new SqlParameter("@F", SqlDbType.Int) { Value = fileId }, P("@P", SqlDbType.VarChar, po, 500));
        }
        private void InsertLine(
            int poId,
            string[] row,
            Dictionary<string, int> columnIndexes)
        {
            string bundle =
                GetCsvValue(row, columnIndexes, "product_type");

            string sku = GetSku(bundle);

            string supplier = ScalarString(
                sqlDss,
                "SELECT TOP(1) ISNULL(mfr_part_nbr,'') " +
                "FROM dbo.product WHERE sku=@S;",
                P("@S", SqlDbType.VarChar, sku, 100));

            ExecuteNonQuery(
                sqlGestion,
                @"INSERT INTO dbo.T_suivi_commande_end_user_cust_po_line
(id_cust_po,order_reference,bundle,sku,supplier_product_ref,end_user_id,end_user_first_name,end_user_last_name,end_user_mail,label_content,comment)
VALUES(@PO,@O,@B,@S,@SUP,@ID,@F,@L,@M,@LAB,@C);",
                new SqlParameter("@PO", SqlDbType.Int) { Value = poId },
                P("@O", SqlDbType.VarChar,
                    GetCsvValue(row, columnIndexes, "order_reference"), 500),
                P("@B", SqlDbType.VarChar, bundle, 500),
                P("@S", SqlDbType.VarChar, sku, 100),
                P("@SUP", SqlDbType.VarChar, supplier, 500),
                P("@ID", SqlDbType.VarChar,
                    GetCsvValue(row, columnIndexes, "end_user_id"), 500),
                P("@F", SqlDbType.VarChar,
                    GetCsvValue(row, columnIndexes, "end_user_first_name"), 500),
                P("@L", SqlDbType.VarChar,
                    GetCsvValue(row, columnIndexes, "end_user_last_name"), 500),
                P("@M", SqlDbType.VarChar,
                    GetCsvValue(row, columnIndexes, "end_user_email"), 500),
                P("@LAB", SqlDbType.VarChar,
                    GetCsvValue(row, columnIndexes, "label_content"), 2000),
                P("@C", SqlDbType.VarChar,
                    GetCsvValue(row, columnIndexes, "comment"), 2000));
        }

        private void SetSuccess(string mailbox, string graphId, int mailId, string folder) => ExecuteNonQuery(sqlGestion, @"UPDATE dbo.T_suivi_commande_end_user_MailProcessing
SET ProcessingStatus='S',MailDatabaseId=@ID,DestinationFolder=@F,ErrorType=NULL,MissingBundles=NULL,ErrorMessage=NULL,AlertSentAtUtc=NULL,ProcessedAtUtc=SYSUTCDATETIME()
WHERE MailboxAddress=@M AND GraphMessageId=@G;", new SqlParameter("@ID", SqlDbType.Int) { Value = mailId }, P("@F", SqlDbType.NVarChar, folder, 100), P("@M", SqlDbType.NVarChar, mailbox, 320), P("@G", SqlDbType.NVarChar, graphId, 500));
        private void UpsertTechnicalError(string mailbox, Message email, Exception ex)
        {
            if (email == null || string.IsNullOrWhiteSpace(email.Id)) return;
            ExecuteNonQuery(sqlGestion, @"MERGE dbo.T_suivi_commande_end_user_MailProcessing t USING(SELECT @M MailboxAddress,@G GraphMessageId)s
ON t.MailboxAddress=s.MailboxAddress AND t.GraphMessageId=s.GraphMessageId WHEN MATCHED THEN UPDATE SET ProcessingStatus='E',ErrorType='TECHNICAL',ErrorMessage=@E,
AttemptCount=t.AttemptCount+1,LastAttemptAtUtc=SYSUTCDATETIME() WHEN NOT MATCHED THEN INSERT(MailboxAddress,GraphMessageId,ProcessingStatus,ErrorType,ErrorMessage,AttemptCount)
VALUES(@M,@G,'E','TECHNICAL',@E,1);", P("@M", SqlDbType.NVarChar, mailbox, 320), P("@G", SqlDbType.NVarChar, email.Id, 500), P("@E", SqlDbType.NVarChar, Truncate(Details(ex), 2000), 2000));
        }

        private List<FileAttachment> GetCsvAttachments(string mailbox, string messageId)
        {
            AttachmentCollectionResponse response = Retry(() => graph.Users[mailbox].Messages[messageId].Attachments.GetAsync(c => Header(c.Headers)).GetAwaiter().GetResult(), "attachments");
            var files = new List<FileAttachment>();
            foreach (Attachment a in response?.Value ?? new List<Attachment>()) if (a is FileAttachment f && string.Equals(Path.GetExtension(f.Name), ".csv", StringComparison.OrdinalIgnoreCase))
            { FileAttachment full = Retry(() => graph.Users[mailbox].Messages[messageId].Attachments[f.Id].GetAsync(c => Header(c.Headers)).GetAwaiter().GetResult() as FileAttachment, "attachment"); if (full != null) files.Add(full); }
            return files;
        }
        private List<Message> GetMessages(string mailbox, string folderId, DateTime filter)
        {
            int top = ParseInt(numberOfMails, 20); MessageCollectionResponse r = Retry(() => graph.Users[mailbox].MailFolders[folderId].Messages.GetAsync(c =>
            {
                Header(c.Headers); c.QueryParameters.Top = top; c.QueryParameters.Orderby = new[] { "receivedDateTime asc" }; c.QueryParameters.Filter =
                "isRead eq false and receivedDateTime gt " + filter.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ", CultureInfo.InvariantCulture);
                c.QueryParameters.Select = new[] { "id", "subject", "receivedDateTime", "from", "sender", "internetMessageId" };
            }).GetAwaiter().GetResult(), "messages");
            return r?.Value ?? new List<Message>();
        }
        private Message GetMessage(string mailbox, string id) => Retry(() => graph.Users[mailbox].Messages[id].GetAsync(c => {
            Header(c.Headers); c.QueryParameters.Select =
            new[] { "id", "subject", "receivedDateTime", "from", "sender", "internetMessageId", "hasAttachments" };
        }).GetAwaiter().GetResult(), "message");
        private byte[] GetMime(string mailbox, string id) { using (Stream i = Retry(() => graph.Users[mailbox].Messages[id].Content.GetAsync(c => Header(c.Headers)).GetAwaiter().GetResult(), "mime")) using (var o = new MemoryStream()) { i.CopyTo(o); return o.ToArray(); } }
        private void MarkProcessed(string mailbox, string id) => Retry(() => {
            graph.Users[mailbox].Messages[id].PatchAsync(new Message
            {
                IsRead = true,
                SingleValueExtendedProperties = new List<SingleValueLegacyExtendedProperty> { new SingleValueLegacyExtendedProperty { Id = ProcessedPropertyId, Value = DateTime.UtcNow.ToString("O") } }
            }, c => Header(c.Headers)).GetAwaiter().GetResult(); return true;
        }, "mark");
        private void MoveRetryableError(
            string mailbox,
            string messageId,
            string errorFolderId)
        {
            Retry(() =>
            {
                graph.Users[mailbox]
                    .Messages[messageId]
                    .PatchAsync(
                        new Message { IsRead = true },
                        configuration => Header(configuration.Headers))
                    .GetAwaiter()
                    .GetResult();
                return true;
            }, "mark retryable error as read");

            Move(mailbox, messageId, errorFolderId);
        }

        private void Move(string mailbox, string id, string folderId) => Retry(() => { graph.Users[mailbox].Messages[id].Move.PostAsync(new Microsoft.Graph.Users.Item.Messages.Item.Move.MovePostRequestBody { DestinationId = folderId }, c => Header(c.Headers)).GetAwaiter().GetResult(); return true; }, "move");
        private void FinalizeGraph(string mailbox, string id, string folderId) { try { MarkProcessed(mailbox, id); Move(mailbox, id, folderId); } catch { } }
        private MailFolder GetFolder(string mailbox, string parent, string display)
        {
            if (display.Equals("Inbox", StringComparison.OrdinalIgnoreCase)) return Retry(() => graph.Users[mailbox].MailFolders["inbox"].GetAsync(c => Header(c.Headers)).GetAwaiter().GetResult(), "inbox");
            MailFolderCollectionResponse r = Retry(() => graph.Users[mailbox].MailFolders[parent].ChildFolders.GetAsync(c => { Header(c.Headers); c.QueryParameters.Top = 100; c.QueryParameters.Filter = "displayName eq '" + display.Replace("'", "''") + "'"; }).GetAwaiter().GetResult(), "folder");
            MailFolder f = r?.Value?.FirstOrDefault(x => string.Equals(x.DisplayName, display, StringComparison.OrdinalIgnoreCase)); if (f == null) throw new DirectoryNotFoundException(display); return f;
        }

        private void Apply(Country i, string imca)
        {
            country = i.country ?? ""; skValid = i.sk_valid ?? ""; name = i.name ?? ""; active = i.active ?? ""; debug = i.debug ?? ""; numberOfMails = i.number_of_mails ?? "20"; startDateScan = i.start_date_scan ?? "";
            mailboxTypology = string.IsNullOrWhiteSpace(i.sharedmailbox_typologie) ? "ECONOCOM" : i.sharedmailbox_typologie;
            bundleUpdateDelayDays = i.delai_update_bundle > 0 ? i.delai_update_bundle : 3;
            folderIn = i.sharedmailbox_folder_in ?? "Inbox"; folderArchive = i.sharedmailbox_folder_archive ?? "archives";
            folderDuplicate = i.sharedmailbox_folder_duplicate ?? "doublons"; folderError = i.sharedmailbox_folder_error ?? "erreur"; sqlGestionParameter = i.sql_connexion_gestion_cdes_parameter_global ?? "";
            sqlDssParameter = i.sql_connexion_dss_parameter_global ?? "";
            alertParameter = i.email_in_case_of_technical_issue_parameter_global ?? "";
            businessAlertRecipients = i.business_alert_recipients ?? "";
            sendAsParameter = i.fr_graph_send_as_parameter_global ?? "";
            econocomSendAs = i.fr_econocom_graph_send_as ?? "";
            imeiConfigurationParameterGlobal = i.imei_configuration_parameter_global ?? "IMEI_MANAGEMENT_FR";
            appleImeiTableName = i.apple_imei_table_name ?? "dbo.IMEI_APPLE";
            sendFilesHours = i.send_files_hours ?? new List<string>();
            oracleOdsParameterGlobal = string.IsNullOrWhiteSpace(i.oracle_ods_parameter_global) ? "DWHP_IMT" : i.oracle_ods_parameter_global;
            imXmlOrderStatusUrlParameterGlobal = string.IsNullOrWhiteSpace(i.imxml_order_status_url_parameter_global) ? "ECONOCOM_IMXML_ORDER_STATUS_URL_FR" : i.imxml_order_status_url_parameter_global;
            imXmlOrderStatusTemplateParameterGlobal = string.IsNullOrWhiteSpace(i.imxml_order_status_template_parameter_global) ? "ECONOCOM_IMXML_ORDER_STATUS_TEMPLATE_FR" : i.imxml_order_status_template_parameter_global;
            imXmlOrderDetailUrlParameterGlobal = string.IsNullOrWhiteSpace(i.imxml_order_detail_url_parameter_global) ? "ECONOCOM_IMXML_ORDER_DETAIL_URL_FR" : i.imxml_order_detail_url_parameter_global;
            imXmlOrderDetailTemplateParameterGlobal = string.IsNullOrWhiteSpace(i.imxml_order_detail_template_parameter_global) ? "ECONOCOM_IMXML_ORDER_DETAIL_TEMPLATE_FR" : i.imxml_order_detail_template_parameter_global;
            imXmlOrderTrackingUrlParameterGlobal = string.IsNullOrWhiteSpace(i.imxml_order_tracking_url_parameter_global) ? "ECONOCOM_IMXML_ORDER_TRACKING_URL_FR" : i.imxml_order_tracking_url_parameter_global;
            imXmlOrderTrackingTemplateParameterGlobal = string.IsNullOrWhiteSpace(i.imxml_order_tracking_template_parameter_global) ? "ECONOCOM_IMXML_ORDER_TRACKING_TEMPLATE_FR" : i.imxml_order_tracking_template_parameter_global;
            delaiRefreshBlTracking = i.Delai_Refresh_BL_Tracking > 0 ? i.Delai_Refresh_BL_Tracking : 180;
            sqlGestion = GetParameter(imca, sqlGestionParameter);
            oracleOdsConnection = GetParameter(imca, oracleOdsParameterGlobal);
            imXmlOrderStatusUrl = GetParameter(imca, imXmlOrderStatusUrlParameterGlobal);
            imXmlOrderStatusTemplate = GetParameter(imca, imXmlOrderStatusTemplateParameterGlobal);
            imXmlOrderDetailUrl = GetParameter(imca, imXmlOrderDetailUrlParameterGlobal);
            imXmlOrderDetailTemplate = GetParameter(imca, imXmlOrderDetailTemplateParameterGlobal);
            imXmlOrderTrackingUrl = GetParameter(imca, imXmlOrderTrackingUrlParameterGlobal);
            imXmlOrderTrackingTemplate = GetParameter(imca, imXmlOrderTrackingTemplateParameterGlobal);
            sqlDss = GetParameter(imca, sqlDssParameter);
            alertRecipients = GetParameter(imca, alertParameter);
            sendAs = GetParameter(imca, sendAsParameter);
        }
        private bool IsCustomerFileSendHour(DateTime localDateTime)
        {
            if (sendFilesHours == null || sendFilesHours.Count == 0)
            {
                return true;
            }

            foreach (string configuredValue in sendFilesHours)
            {
                string value = (configuredValue ?? "").Trim();

                if (string.IsNullOrWhiteSpace(value))
                {
                    continue;
                }

                if (TimeSpan.TryParseExact(
                        value,
                        new[]
                        {
                            "h",
                            "hh",
                            @"h\:mm",
                            @"hh\:mm"
                        },
                        CultureInfo.InvariantCulture,
                        out TimeSpan configuredTime) &&
                    configuredTime.Hours == localDateTime.Hour)
                {
                    return true;
                }
            }

            return false;
        }

        private void Validate()
        {
            if (new[]
                {
                    sqlGestion,
                    sqlDss,
                    alertRecipients,
                    businessAlertRecipients,
                    sendAs,
                    mailboxTypology
                }.Any(string.IsNullOrWhiteSpace))
            {
                throw new InvalidOperationException(
                    "Configuration incomplete");
            }
        }
        private GraphServiceClient ConnectGraph() { ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12; return new class_dev_tools.Ews_Modern_Auth().Get_Graph_Service(); }
        private T Retry<T>(Func<T> action, string operation)
        {
            const int maximumAttempts = 3;
            Exception lastException = null;

            for (int attempt = 1; attempt <= maximumAttempts; attempt++)
            {
                try
                {
                    return action();
                }
                catch (Exception ex) when (IsTransientGraphError(ex))
                {
                    lastException = ex;
                    if (attempt >= maximumAttempts)
                    {
                        break;
                    }

                    int delayMilliseconds = attempt * 5000;
                    WriteLog(
                        "       Temporary Graph error during " + operation +
                        ". Application attempt " + attempt + "/" +
                        maximumAttempts + ". Retry in " +
                        delayMilliseconds + " ms. Details : " +
                        Details(ex));
                    System.Threading.Thread.Sleep(delayMilliseconds);
                }
            }

            throw new InvalidOperationException(
                "Graph operation failed after " + maximumAttempts +
                " application attempt(s) : " + operation + " - " +
                Details(lastException),
                lastException);
        }

        private static bool IsTransientGraphError(Exception exception)
        {
            Exception current = exception;
            while (current != null)
            {
                if (current is HttpRequestException ||
                    current is TimeoutException ||
                    current is TaskCanceledException)
                {
                    return true;
                }

                if (current is ApiException apiException)
                {
                    int status = apiException.ResponseStatusCode;
                    if (status == 408 || status == 429 || status == 500 ||
                        status == 502 || status == 503 || status == 504)
                    {
                        return true;
                    }
                }

                ODataError graphError = current as ODataError;
                string graphCode = graphError?.Error?.Code ?? "";
                if (graphCode.Equals("TooManyRequests", StringComparison.OrdinalIgnoreCase) ||
                    graphCode.Equals("ErrorServerBusy", StringComparison.OrdinalIgnoreCase) ||
                    graphCode.Equals("ServiceUnavailable", StringComparison.OrdinalIgnoreCase) ||
                    graphCode.Equals("ApplicationThrottled", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                string text = current.Message ?? "";
                if (text.IndexOf("Too many retries performed", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    text.IndexOf("More than 3 retries encountered", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    text.IndexOf("TooManyRequests", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    text.IndexOf("ApplicationThrottled", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    text.IndexOf("temporarily unavailable", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    text.IndexOf("service unavailable", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    text.IndexOf("12002", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    text.IndexOf("timed out", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    text.IndexOf("operation was canceled", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    text.IndexOf("An error occurred while sending the request", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
                current = current.InnerException;
            }
            return false;
        }
        private void SendTechnicalAlert(
            string method,
            string text)
        {
            SendAlertToRecipients(
                method,
                text,
                alertRecipients);
        }

        private void SendBusinessAlert(
            string method,
            string text)
        {
            SendAlertToRecipients(
                method,
                text,
                businessAlertRecipients);
        }

        private void SendAlertToRecipients(
            string method,
            string text,
            string recipientsValue)
        {
            try
            {
                if (graph == null)
                {
                    graph = ConnectGraph();
                }

                List<Recipient> recipients =
                    Recipients(recipientsValue);

                if (recipients.Count == 0)
                {
                    throw new InvalidOperationException(
                        "No valid alert recipient configured for " +
                        method);
                }

                var sender = new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = sendAs,
                        Name =
                            "Ingram Micro - Gestion des commandes " +
                            "Econocom/Airbus"
                    }
                };

                Retry(
                    () =>
                    {
                        graph.Users[sendAs].SendMail.PostAsync(
                            new Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody
                            {
                                Message = new Message
                                {
                                    Subject =
                                        "Gestion des commandes Econocom/Airbus - " +
                                        method,
                                    From = sender,
                                    Sender = sender,
                                    Body = new ItemBody
                                    {
                                        ContentType = BodyType.Text,
                                        Content = text
                                    },
                                    ToRecipients = recipients
                                },
                                SaveToSentItems = true
                            }).GetAwaiter().GetResult();
                        return true;
                    },
                    "Send Graph alert : " + method);
            }
            catch (Exception ex)
            {
                WriteLog(
                    "Alert error: " + Details(ex));
            }
        }

        private static string ValidateHttpUrl(string value, string name) { string u = (value ?? "").Trim(); if (!Uri.TryCreate(u, UriKind.Absolute, out Uri uri) || (uri.Scheme != "https" && uri.Scheme != "http")) throw new InvalidOperationException("Invalid URL configured for " + name); return uri.AbsoluteUri; }
        private string GetParameter(string c, string p) => string.IsNullOrWhiteSpace(p) ? "" : ScalarString(c, "SELECT ISNULL(VALUE,'') FROM PCM_TAB_IMCA_PARAMETER_GLOBAL WHERE SK_VALID=0 AND PARAMETER=@P;", P("@P", SqlDbType.NVarChar, p, 255));
        private static void ExecuteNonQuery(string c, string s, params SqlParameter[] p) { using (var cn = new SqlConnection(c)) using (var cmd = new SqlCommand(s, cn)) { if (p != null) cmd.Parameters.AddRange(p); cn.Open(); cmd.ExecuteNonQuery(); } }
        private static string ScalarString(string c, string s, params SqlParameter[] p) { using (var cn = new SqlConnection(c)) using (var cmd = new SqlCommand(s, cn)) { if (p != null) cmd.Parameters.AddRange(p); cn.Open(); return Convert.ToString(cmd.ExecuteScalar()).Trim(); } }
        private static int ScalarInt(string c, string s, params SqlParameter[] p) { using (var cn = new SqlConnection(c)) using (var cmd = new SqlCommand(s, cn)) { if (p != null) cmd.Parameters.AddRange(p); cn.Open(); return Convert.ToInt32(cmd.ExecuteScalar()); } }
        private static SqlParameter P(string n, SqlDbType t, object v, int size) { var p = new SqlParameter(n, t, size); p.Value = v ?? ""; return p; }
        private bool IsDue(SharedMailbox m) => !m.LastRefresh.HasValue || m.RefreshMinutes <= 0 || m.LastRefresh.Value.AddMinutes(m.RefreshMinutes) <= DateTime.Now;
        private void DeleteTempFiles(List<CsvFileValidation> files) { foreach (var f in files ?? new List<CsvFileValidation>()) try { if (File.Exists(f.TempPath)) File.Delete(f.TempPath); } catch { } }
        private void WriteLog(string m) { try { string p = Path.Combine(logsFolder, "IMCA_" + sessionName + "_" + DateTime.Now.ToString("dd_MM_yyyy") + "_" + country + "_" + ApplicationName + ".txt"); lock (LogLock) File.AppendAllText(p, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - " + m + Environment.NewLine, Encoding.UTF8); } catch { } }
        private static string Details(Exception exception)
        {
            if (exception == null) return "Unknown error";
            var details = new List<string>();
            Exception current = exception;
            while (current != null)
            {
                details.Add("Type=" + current.GetType().FullName);
                if (!string.IsNullOrWhiteSpace(current.Message))
                    details.Add("Message=" + current.Message);
                ODataError graphError = current as ODataError;
                if (!string.IsNullOrWhiteSpace(graphError?.Error?.Code))
                    details.Add("GraphCode=" + graphError.Error.Code);
                if (!string.IsNullOrWhiteSpace(graphError?.Error?.Message))
                    details.Add("GraphMessage=" + graphError.Error.Message);
                if (current is ApiException apiException)
                    details.Add("HttpStatus=" + apiException.ResponseStatusCode);
                current = current.InnerException;
            }
            return string.Join(" | ", details
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase));
        }
        private DateTime ParseStartDate() => DateTime.TryParseExact(startDateScan, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime d) ? d : new DateTime(1900, 1, 1);
        private static int ParseInt(string v, int f) => int.TryParse(v, out int n) && n > 0 ? n : f;
        private static List<Recipient> Recipients(string v) => (v ?? "").Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).Where(x => Regex.IsMatch(x, @"^[^\s@]+@[^\s@]+\.[^\s@]+$")).Distinct(StringComparer.OrdinalIgnoreCase).Select(x => new Recipient { EmailAddress = new EmailAddress { Address = x } }).ToList();
        private static string GetSender(Message m) => m?.From?.EmailAddress?.Address ?? m?.Sender?.EmailAddress?.Address ?? "";
        private static string Truncate(string v, int n) { v = v ?? ""; return v.Length <= n ? v : v.Substring(0, n); }
        private static bool IsTrue(string v) => string.Equals(v?.Trim(), "TRUE", StringComparison.OrdinalIgnoreCase);
        private static void Header(RequestHeaders h) => h.Add("Prefer", ImmutablePreference);
        private static string GetServicePath() { string l = System.Reflection.Assembly.GetEntryAssembly()?.Location; return string.IsNullOrWhiteSpace(l) ? AppDomain.CurrentDomain.BaseDirectory : Path.GetDirectoryName(l); }
    }
}
