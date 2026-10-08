using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using Microsoft.VisualBasic.FileIO;
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
using System.Text;
using System.Text.RegularExpressions;

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
            public List<CsvFileValidation> Files = new List<CsvFileValidation>();
        }

        private static readonly Dictionary<string, string[]> CsvColumnAliases =
            new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
            {
                { "order_number", new[] { "Order_Number" } },
                { "order_reference", new[] { "Order_Reference" } },
                { "bundle", new[] { "Bundle", "Product_Bundle" } },
                { "end_user_first_name", new[] { "EndUser_First_Name" } },
                { "end_user_last_name", new[] { "EndUser_First_Last_Name", "EndUser_Last_Name" } },
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
            graph.Users[econocomSendAs].SendMail.PostAsync(new Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody { Message = message, SaveToSentItems = true }).GetAwaiter().GetResult();
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
                            SetValidationTerminalStatus(
                                mailbox.Address,
                                email.Id,
                                folderError);

                            MarkProcessed(
                                mailbox.Address,
                                email.Id);

                            Move(
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

                            SetValidationTerminalStatus(
                                mailbox.Address,
                                email.Id,
                                folderError);

                            MarkProcessed(
                                mailbox.Address,
                                email.Id);

                            Move(
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
                                string actualHeader =
                                    NormalizeCsvHeader(fields);

                                string expectedHeader =
                                    NormalizeCsvHeader(
                                        expectedHeaderFields);

                                if (!actualHeader.Equals(
                                        expectedHeader,
                                        StringComparison.OrdinalIgnoreCase))
                                {
                                    result.ErrorType = "CSV_STRUCTURE";
                                    result.Message =
                                        "Le fichier " + safeName +
                                        " ne correspond pas à l'en-tête configuré " +
                                        "pour le client " + customer +
                                        ". En-tête attendu : " +
                                        clientConfiguration.ExpectedCsvHeader +
                                        ". En-tête reçu : " +
                                        string.Join(",", fields);
                                    return result;
                                }

                                file.ColumnIndexes =
                                    BuildCsvColumnIndexes(fields);

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

                            if (fields.Length != expectedColumnCount)
                            {
                                result.ErrorType = "CSV_STRUCTURE";
                                result.Message =
                                    "Le fichier " + safeName +
                                    " contient une ligne avec " +
                                    fields.Length +
                                    " colonne(s), alors que l'en-tête configuré " +
                                    "en attend " + expectedColumnCount;
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
                                    "bundle");

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

        private void SendValidationAlert(string mailbox, Message email, CsvValidationResult validation)
        {
            string text = validation.ErrorType == "MISSING_BUNDLE"
                ? "Commande Econocom/Airbus laissée en attente. Boîte : " + mailbox + ". Sujet : " + (email.Subject ?? "") +
                  ". Bundle(s) à créer dans T_suivi_commande_end_user_ref_sku : " + string.Join(", ", validation.MissingBundles) +
                  ". La commande sera reprise automatiquement après création."
                : "Commande Econocom/Airbus laissée en attente. Boîte : " + mailbox + ". Sujet : " + (email.Subject ?? "") + ". " + validation.Message;
            SendBusinessAlert(
                "Validation CSV",
                text);
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

        private void SetValidationTerminalStatus(
            string mailbox,
            string graphId,
            string destinationFolder)
        {
            ExecuteNonQuery(
                sqlGestion,
                @"UPDATE dbo.T_suivi_commande_end_user_MailProcessing
SET ProcessingStatus='S',
    DestinationFolder=@FOLDER,
    ProcessedAtUtc=SYSUTCDATETIME(),
    LastAttemptAtUtc=SYSUTCDATETIME()
WHERE MailboxAddress=@MAILBOX
  AND GraphMessageId=@GRAPH;",
                P(
                    "@FOLDER",
                    SqlDbType.NVarChar,
                    destinationFolder,
                    100),
                P(
                    "@MAILBOX",
                    SqlDbType.NVarChar,
                    mailbox,
                    320),
                P(
                    "@GRAPH",
                    SqlDbType.NVarChar,
                    graphId,
                    500));
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

        private static string NormalizeCsvHeader(string[] fields)
        {
            return string.Join(
                ",",
                (fields ?? Array.Empty<string>())
                    .Select(field =>
                        NormalizeCsvHeaderValue(field)));
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
                .Replace("", "")
                .Replace("", "");
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
                GetCsvValue(row, columnIndexes, "bundle");

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
            sqlGestion = GetParameter(imca, sqlGestionParameter);
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
        private T Retry<T>(Func<T> action, string op) { Exception last = null; for (int n = 1; n <= 3; n++) { try { return action(); } catch (Exception ex) when (ex is HttpRequestException || ex is TimeoutException) { last = ex; System.Threading.Thread.Sleep(n * 5000); } } throw new InvalidOperationException(op, last); }
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
                                ContentType = BodyType.Html,
                                Content = WebUtility.HtmlEncode(text)
                            },
                            ToRecipients = recipients
                        },
                        SaveToSentItems = true
                    }).GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                WriteLog(
                    "Alert error: " + Details(ex));
            }
        }

        private string GetParameter(string c, string p) => string.IsNullOrWhiteSpace(p) ? "" : ScalarString(c, "SELECT ISNULL(VALUE,'') FROM PCM_TAB_IMCA_PARAMETER_GLOBAL WHERE SK_VALID=0 AND PARAMETER=@P;", P("@P", SqlDbType.NVarChar, p, 255));
        private static void ExecuteNonQuery(string c, string s, params SqlParameter[] p) { using (var cn = new SqlConnection(c)) using (var cmd = new SqlCommand(s, cn)) { if (p != null) cmd.Parameters.AddRange(p); cn.Open(); cmd.ExecuteNonQuery(); } }
        private static string ScalarString(string c, string s, params SqlParameter[] p) { using (var cn = new SqlConnection(c)) using (var cmd = new SqlCommand(s, cn)) { if (p != null) cmd.Parameters.AddRange(p); cn.Open(); return Convert.ToString(cmd.ExecuteScalar()).Trim(); } }
        private static int ScalarInt(string c, string s, params SqlParameter[] p) { using (var cn = new SqlConnection(c)) using (var cmd = new SqlCommand(s, cn)) { if (p != null) cmd.Parameters.AddRange(p); cn.Open(); return Convert.ToInt32(cmd.ExecuteScalar()); } }
        private static SqlParameter P(string n, SqlDbType t, object v, int size) { var p = new SqlParameter(n, t, size); p.Value = v ?? ""; return p; }
        private bool IsDue(SharedMailbox m) => !m.LastRefresh.HasValue || m.RefreshMinutes <= 0 || m.LastRefresh.Value.AddMinutes(m.RefreshMinutes) <= DateTime.Now;
        private string GetMailCustomerNumber(int id) => ScalarString(sqlGestion, "SELECT ISNULL(cust_nbr,'') FROM dbo.T_suivi_commande_end_user_mail WHERE id_mail=@ID;", new SqlParameter("@ID", SqlDbType.Int) { Value = id });
        private void DeleteTempFiles(List<CsvFileValidation> files) { foreach (var f in files ?? new List<CsvFileValidation>()) try { if (File.Exists(f.TempPath)) File.Delete(f.TempPath); } catch { } }
        private void WriteLog(string m) { try { string p = Path.Combine(logsFolder, "IMCA_" + sessionName + "_" + DateTime.Now.ToString("dd_MM_yyyy") + "_" + country + "_" + ApplicationName + ".txt"); lock (LogLock) File.AppendAllText(p, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - " + m + Environment.NewLine, Encoding.UTF8); } catch { } }
        private static string Details(Exception e) => e == null ? "Unknown error" : e.GetType().FullName + " | " + e.Message + (e.InnerException == null ? "" : " | " + Details(e.InnerException));
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
