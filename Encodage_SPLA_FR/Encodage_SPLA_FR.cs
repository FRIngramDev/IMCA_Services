using ClosedXML.Excel;
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
using System.Linq;
using System.Text;
using System.Threading;


namespace ENCODAGE_SPLA_FR
{
    public class ENCODAGE_SPLA_FR
    {
        public ENCODAGE_SPLA_FR()
        {
        }

        // ----- Paramètres courants pour le pays / traitement -----
        private string country = "";                       // Code pays traité
        private string sk_valid = "";                      // Paramètre SK valid
        private string name = "";                          // Nom lisible du pays / instance
        private string active = "";                        // Indique si le traitement est actif ("TRUE")
        private string debug = "";                         // Mode debug (logs détaillés si TRUE)
        private string start_date_scan = "";               // Date de départ pour la recherche d'emails
        private string number_of_mails = "10";             // Nombre max d'emails à traiter

        // ----- Boîte partagée (shared mailbox) et dossiers -----
        private string sharedmailbox_name = "";            // Adresse de la boîte partagée
        private string sharedmailbox_folder_in = "";       // Dossier source pour la lecture
        private string sharedmailbox_folder_out = "";      // Dossier de destination après traitement

        // ----- Paramètres d'alerte / email technique -----
        private string email_in_case_of_technical_issue_parameter_global = ""; // Nom du param global
        private string email_in_case_of_technical_issue = "";                  // Adresse email réelle
        private string email_destinataire = "";                  // Adresse email destinataire 

        // ----- Fichiers / logs / session -----
        private string logs_folder = "";                   // Répertoire des logs
        private string temp_folder = "";                   // Répertoire temporaire pour fichiers
        private string global_session_name = "";           // Nom de session pour logs et opérations

        // ----- Constantes / noms -----
        private string global_application_name = "ENCODAGE_SPLA_FR"; // Nom de l'application utilisé dans les logs

        // ----- Connexions SQL (chaînes / paramètres) -----
        private string sql_connexion = "";                         // Connexion principale (résolue dynamiquement)
        private string sql_connexion_parameter_global = "";        // Nom du paramètre contenant la connexion
        private string sql_connexion_macros = "";                // Connexion macros
        private string sql_macros_parameter_global = "";         // Nom du param macros

        // ----- Client Microsoft Graph (utilisé pour lecture/envoi emails) -----
        private GraphServiceClient graphService = null;

        private string timer_attente_impulse = ""; // Timer pour attendre la fin de l'impulse


        /// <summary>
        /// Représentation du fichier JSON de configuration attendu.
        /// Contient le dossier de logs global et la liste des pays configurés.
        /// </summary>
        public class JSON_file
        {
            public string logs_folder { get; set; } = "";                 // Répertoire des logs défini dans le JSON
            public List<Country> countries { get; set; } = new List<Country>(); // Liste des entrées pays
        }

        public class Country
        {
            public string country { get; set; } = "";                              // Code pays
            public string sk_valid { get; set; } = "";                             // Paramètre sk_valid
            public string name { get; set; } = "";                                 // Nom affichable
            public string active { get; set; } = "";                               // Active (TRUE/FALSE)
            public string debug { get; set; } = "";                                // Mode debug
            public string start_date_scan { get; set; } = "";                      // Date de début du scan (format libre)
            public string number_of_mails { get; set; } = "10";                    // Nombre d'emails à traiter

            public string sharedmailbox_name { get; set; } = "";                   // Boîte partagée
            public string sharedmailbox_folder_in { get; set; } = "INBOX";         // Dossier source
            public string sharedmailbox_folder_out { get; set; } = "Archives";     // Dossier cible

            public string email_in_case_of_technical_issue_parameter_global { get; set; } = ""; // Param nom destinataire

            public string sql_connexion_parameter_global { get; set; } = "";        // Param nom chaine SQL principale

            public string sql_macros_parameter_global { get; set; } = "";           // Param nom chaine SQL macros

            public string sql_annuaire_parameter_global { get; set; } = "";         // Param nom chaine SQL annuaire

            public string sql_incentives_parameter_global { get; set; } = "";
            public string email_destinataire { get; set; } = "";       // Param nom chaine SQL incentives

            public string timer_attente_impulse { get; set; } = ""; // Timer pour attendre la fin de l'impulse
        }

        /// <summary>
        /// Etablit et retourne une instance de GraphServiceClient configurée
        /// pour accéder aux boîtes aux lettres via Microsoft Graph.
        /// </summary>
        private GraphServiceClient Connexion_Microsoft_Graph()
        {
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;

            // On utilise EWS API MANAGED REFERENCE 2.1
            GraphServiceClient GraphService;

            class_dev_tools.Ews_Modern_Auth Ews_Modern_Auth = new class_dev_tools.Ews_Modern_Auth();

            GraphService = Ews_Modern_Auth.Get_Graph_Service();

            return GraphService;
        }

        /// <summary>
        /// Point d'entrée principal pour la lecture des emails depuis la boîte partagée.
        /// - sql_con : chaîne de connexion pour retrouver les paramètres globaux
        /// - logs, tmp_folder : répertoires relatifs pour logs et fichiers temporaires
        /// - session_name : nom de session utilisé dans les logs
        /// </summary>
        public void Read_Email_with_Graph(string sql_con, string logs, string tmp_folder, string session_name)
        {
            string global_parameters = "";
            string service_path = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
            int nb_mail = 0;


            setlogs_folder(service_path + "\\" + logs);
            settemp_folder(service_path + "\\" + tmp_folder);
            global_session_name = session_name;

            try
            {
                // Get the global parameters of the ACTION
                global_parameters = get_IMCA_paramters(sql_con, global_application_name);

                if (string.IsNullOrWhiteSpace(global_parameters))
                {
                    WriteToFile("No parameters found for " + global_application_name);
                    return;
                }

                JSON_file param = JsonConvert.DeserializeObject<JSON_file>(global_parameters);

                if (param == null || param.countries == null || param.countries.Count == 0)
                {
                    WriteToFile(global_application_name + " parameters are empty or invalid");
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

                    setSharedMailboxName(p.sharedmailbox_name);
                    setsharedmailbox_folder_in(p.sharedmailbox_folder_in);
                    setsharedmailbox_folder_out(p.sharedmailbox_folder_out);

                    setNumber_of_mails(p.number_of_mails);
                    setEmailDestinataire(p.email_destinataire);


                    setSqlConnexionParam(p.sql_connexion_parameter_global);
                    setSqlConnexionMacrosParam(p.sql_macros_parameter_global);
                    // setSqlConnexionIncentivesParam(p.sql_incentives_parameter_global);

                    setEmailInCaseOfTechnicalIssueParam(p.email_in_case_of_technical_issue_parameter_global);

                    email_in_case_of_technical_issue = get_IMCA_paramters(sql_con, email_in_case_of_technical_issue_parameter_global);

                    sql_connexion = get_IMCA_paramters(sql_con, sql_connexion_parameter_global);
                    sql_connexion_macros = get_IMCA_paramters(sql_con, sql_macros_parameter_global);

                    setTimerAttenteImpulse(p.timer_attente_impulse);


                    if (active.ToUpper().Trim() != "TRUE")
                    {
                        continue;
                    }

                    try
                    {
                        WriteToFile(name.ToUpper() + "(" + country.ToUpper() + ")" + " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                        WriteToFile("   Debug Parameter is set to " + debug.ToUpper() + " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));

                        if (debug.ToUpper() == "TRUE")
                        {
                            WriteToFile("   Connexion to " + sharedmailbox_name + " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                        }

                        graphService = Connexion_Microsoft_Graph();


                        ValidateRequiredParameters();

                        if (!Directory.Exists(temp_folder))
                        {
                            Directory.CreateDirectory(temp_folder);
                        }

                        DeleteFiles(temp_folder, global_application_name);

                        if (debug.ToUpper() == "TRUE")
                        {
                            WriteToFile("   Sharedmailbox_folder_in  : " + sharedmailbox_folder_in);
                            WriteToFile("   Sharedmailbox_folder_out : " + sharedmailbox_folder_out);
                            WriteToFile("   Temp folder              : " + temp_folder);
                            WriteToFile("   Extracting the " + number_of_mails + " oldest messages");
                        }

                        MailFolder inputFolder = GetInputFolder(graphService);
                        MailFolder archiveFolder = GetChildFolderByName(graphService, sharedmailbox_name, sharedmailbox_folder_out);
                        MailFolder errorFolder = GetChildFolderByName(graphService, sharedmailbox_name, "Erreur");

                        MessageCollectionResponse messages = GetMessagesToProcess(graphService, inputFolder.Id);

                        nb_mail = 0;

                        if (messages != null && messages.Value != null && messages.Value.Count > 0)
                        {
                            if (debug.ToUpper() == "TRUE")
                            {
                                WriteToFile("   " + messages.Value.Count + " Email(s) found at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                            }

                            foreach (Message email in messages.Value)
                            {
                                try
                                {
                                    if (debug.ToUpper() == "TRUE")
                                    {
                                        WriteToFile("       Subject : " + email.Subject + " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                                    }


                                    if (email.HasAttachments == true)
                                    {
                                        List<string> attachments = DownloadExcelAttachmentsWithGraph(
                                            graphService,
                                            sharedmailbox_name,
                                            email.Id
                                        );

                                        if (attachments.Count == 0)
                                        {
                                            throw new Exception("Aucune pièce jointe Excel .xls trouvée dans le mail : " + email.Subject);
                                        }

                                        foreach (string attachment in attachments)
                                        {
                                            string resultFile = ProcessSplaWorkbook(attachment, email.Subject);

                                            EnvoiEmail_with_Graph(
                                                graphService,
                                                "Fichier Commandes SPLA traité",
                                                "Traitement du mail : '" + email.Subject + "' terminé.",
                                                email_destinataire,
                                                resultFile
                                            );
                                        }
                                    }
                                    else
                                    {
                                        throw new Exception("Le mail ne contient aucune pièce jointe : " + email.Subject);
                                    }

                                    MarkEmailAsRead(graphService, sharedmailbox_name, email.Id);
                                    MoveEmail(graphService, sharedmailbox_name, email.Id, archiveFolder.Id);

                                    nb_mail++;

                                    System.Threading.Thread.Sleep(500);
                                }
                                catch (Exception ex)
                                {
                                    WriteToFile("   Error processing email : " + ex.Message + " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));

                                    try
                                    {


                                        EnvoiEmail_with_Graph(
                                            graphService,
                                            "Encodage SPLA FR",
                                            "Encodage SPLA FR - Une erreur (" + ex.Message + ") est survenue lors du traitement d'un fichier. Le mail a été déplacé dans le dossier Erreur.",
                                            email_in_case_of_technical_issue
                                        );
                                    }
                                    catch (Exception mailEx)
                                    {
                                        WriteToFile("   Error sending alert email : " + mailEx.Message);
                                    }

                                    try
                                    {
                                        MoveEmail(graphService, sharedmailbox_name, email.Id, errorFolder.Id);
                                    }
                                    catch (Exception moveEx)
                                    {
                                        WriteToFile("   Error moving email to Erreur folder : " + moveEx.Message);
                                    }
                                }
                            }

                            if (debug.ToUpper() == "TRUE")
                            {
                                WriteToFile("   " + nb_mail + " Email(s) have been processed at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                            }
                        }
                        else
                        {
                            if (debug.ToUpper() == "TRUE")
                            {
                                WriteToFile("   No emails found");
                            }
                        }

                        messages = null;
                    }
                    catch (Exception e)
                    {
                        WriteToFile("   Error get emails : " + e.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteToFile("Global error Read_Email_with_Graph : " + ex.Message);
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }


        /// <summary>
        /// Récupère la collection de messages à traiter depuis le dossier spécifié.
        /// Le nombre d'emails récupérés est contrôlé par 'number_of_mails' et le filtre
        /// par la date de début si défini.
        /// </summary>
        private MessageCollectionResponse GetMessagesToProcess(GraphServiceClient graphService, string folderId)
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
                    config.QueryParameters.Orderby = new string[] { "receivedDateTime asc" };
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

        /// <summary>
        /// Construit la clause de filtre OData pour la récupération des messages
        /// en fonction de la date de début 'start_date_scan'.
        /// </summary>
        private string BuildMessageFilter()
        {
            if (string.IsNullOrWhiteSpace(start_date_scan))
            {
                return "";
            }

            DateTime searchDate;

            if (DateTime.TryParse(start_date_scan, out searchDate))
            {
                return "receivedDateTime ge " + searchDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ");
            }

            if (start_date_scan.Length == 8)
            {
                searchDate = new DateTime(
                    int.Parse(start_date_scan.Substring(0, 4)),
                    int.Parse(start_date_scan.Substring(4, 2)),
                    int.Parse(start_date_scan.Substring(6, 2))
                );

                return "receivedDateTime ge " + searchDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ");
            }

            return "";
        }

        /// <summary>
        /// Retourne l'objet MailFolder correspondant au dossier d'entrée configuré
        /// (INBOX ou un dossier enfant) de la boîte partagée.
        /// </summary>
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

            return GetChildFolderByName(graphService, sharedmailbox_name, sharedmailbox_folder_in);
        }

        private MailFolder GetChildFolderByName(GraphServiceClient graphService, string mailbox, string folderName)
        {
            string safeFolderName = EscapeODataString(folderName);

            MailFolderCollectionResponse folders = graphService.Users[mailbox]
                .MailFolders["inbox"]
                .ChildFolders
                .GetAsync(config =>
                {
                    config.QueryParameters.Filter = $"displayName eq '{safeFolderName}'";
                })
                .GetAwaiter()
                .GetResult();

            if (folders == null || folders.Value == null || folders.Value.Count == 0)
            {
                throw new Exception("Folder not found : " + folderName);
            }

            return folders.Value.First();
        }

        private List<string> DownloadExcelAttachmentsWithGraph(GraphServiceClient graphService, string mailbox, string messageId)
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
                    if (extension != ".xls")
                    {
                        continue;
                    }

                    if (fileAttachment.ContentBytes == null)
                    {
                        throw new Exception("Le contenu de la pièce jointe est vide : " + fileAttachment.Name);
                    }

                    string fileName = CleanFileName(global_application_name + "_" + fileAttachment.Name);
                    string filePath = Path.Combine(temp_folder, fileName);
                    File.WriteAllBytes(filePath, fileAttachment.ContentBytes);
                    downloadedFiles.Add(filePath);
                }
            }

            return downloadedFiles;
        }

        private string ProcessSplaWorkbook(string filePath, string subject)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                throw new FileNotFoundException("Fichier SPLA introuvable", filePath);
            }

            try
            {
                DataTable sourceSheet = ReadSplaSheet(filePath);
                ImportSplaRows(sourceSheet);
                StartImpulseAndWait();
                string resultFile = ExportSplaResultsToXlsx(sourceSheet, filePath);
                WriteToFile("       SPLA file processed : " + resultFile);
                return resultFile;
            }
            catch (Exception ex)
            {
                throw new Exception("Erreur pendant le traitement SPLA du mail '" + subject + "' : " + ex.Message, ex);
            }
        }

        private DataTable ReadSplaSheet(string filePath)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
            {
                DataSet dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = false
                    }
                });

                DataTable sheet = dataSet.Tables.Cast<DataTable>()
                    .FirstOrDefault(t => string.Equals(t.TableName.Trim(), "SPLA", StringComparison.OrdinalIgnoreCase));

                if (sheet == null)
                {
                    throw new Exception("L'onglet SPLA est introuvable dans le fichier Excel");
                }

                return sheet.Copy();
            }
        }

        private void ImportSplaRows(DataTable sheet)
        {
            const string truncateSql = "TRUNCATE TABLE T_SPLA_encodage";
            const string insertSql = @"INSERT INTO T_SPLA_encodage
                (ReportYear,ReportMonth,BranchCustomerNbr,AgreementNbr,ReportType,CustomerPo,
                 EnduserEnrollmentNbr,EndcustomerContact,EndcustomerEmail,EndcustomerName,
                 EndcustomerStreet,EndcustomerZIP,EndcustomerCity,EndcustomerCountry,SKU,
                 MfrPartNbr,Description,QTY,CustomerPricing,TotalPrice,top_traite)
                VALUES
                (@ReportYear,@ReportMonth,@BranchCustomerNbr,@AgreementNbr,@ReportType,@CustomerPo,
                 @EnduserEnrollmentNbr,@EndcustomerContact,@EndcustomerEmail,@EndcustomerName,
                 @EndcustomerStreet,@EndcustomerZIP,@EndcustomerCity,@EndcustomerCountry,@SKU,
                 @MfrPartNbr,@Description,@QTY,@CustomerPricing,@TotalPrice,'N')";

            try
            {
                using (SqlConnection connection = new SqlConnection(sql_connexion))
                {
                    connection.Open();
                    using (SqlTransaction transaction = connection.BeginTransaction())
                    {
                        ExecuteNonQuery(connection, transaction, truncateSql);

                        for (int rowIndex = 1; rowIndex < sheet.Rows.Count; rowIndex++)
                        {
                            DataRow row = sheet.Rows[rowIndex];
                            if (string.IsNullOrWhiteSpace(CellText(row, 0)))
                            {
                                break;
                            }

                            ExecuteNonQuery(connection, transaction, insertSql,
                                P("@ReportYear", CellText(row, 0)),
                                P("@ReportMonth", CellText(row, 1)),
                                P("@BranchCustomerNbr", CellText(row, 2)),
                                P("@AgreementNbr", CellText(row, 3)),
                                P("@ReportType", CellText(row, 4)),
                                P("@CustomerPo", CellText(row, 5)),
                                P("@EnduserEnrollmentNbr", CellText(row, 6)),
                                P("@EndcustomerContact", CellText(row, 7)),
                                P("@EndcustomerEmail", CellText(row, 8)),
                                P("@EndcustomerName", CellText(row, 9)),
                                P("@EndcustomerStreet", CellText(row, 10)),
                                P("@EndcustomerZIP", CellText(row, 11)),
                                P("@EndcustomerCity", CellText(row, 12)),
                                P("@EndcustomerCountry", CellText(row, 13)),
                                P("@SKU", CellText(row, 14)),
                                P("@MfrPartNbr", CellText(row, 15)),
                                P("@Description", CellText(row, 16)),
                                P("@QTY", CellText(row, 17)),
                                DecimalParameter("@CustomerPricing", CellValue(row, 18)),
                                DecimalParameter("@TotalPrice", CellValue(row, 19)));
                        }

                        transaction.Commit();
                    }
                }
            }
            catch (Exception ex)
            {
                SendSqlTechnicalIssueMail(nameof(ImportSplaRows), truncateSql + Environment.NewLine + insertSql, ex);
                throw;
            }
        }

        private void StartImpulseAndWait()
        {
            const string startSql =
                "UPDATE T_SPLA_impulse SET imp_en_cours = 1";

            const string statusSql =
                "SELECT TOP 1 ISNULL(imp_en_cours, 0) " +
                "FROM T_SPLA_impulse";

            ExecuteNonQuerySql(
                nameof(StartImpulseAndWait),
                sql_connexion,
                startSql);

            class_dev_tools.cls_ProcessActif_ala_Dde process =
                new class_dev_tools.cls_ProcessActif_ala_Dde(
                    sql_connexion_macros);

            process.Id_appli = 195;
            process.InsertionDemande();

            DateTime startTime = DateTime.Now;
            DateTime nextAlertTime = startTime.AddMinutes(Convert.ToInt32(timer_attente_impulse));

            while (Convert.ToBoolean(
                ExecuteScalarSql(
                    nameof(StartImpulseAndWait),
                    sql_connexion,
                    statusSql) ?? false))
            {
                TimeSpan elapsed = DateTime.Now - startTime;

                // Send another alert every XX ( timer_attente_impulse ) minutes while the macro is blocked.
                if (DateTime.Now >= nextAlertTime)
                {
                    SendImpulseRestartAlert(elapsed);
                    nextAlertTime = DateTime.Now.AddMinutes(Convert.ToInt32(timer_attente_impulse));
                }

                WriteToFile(
                    "       Waiting for macro 195. Elapsed time : " +
                    elapsed.ToString(@"hh\:mm\:ss"));

                Thread.Sleep(30000);
            }

            WriteToFile(
                "       Macro 195 completed after " +
                (DateTime.Now - startTime).ToString(@"hh\:mm\:ss"));
        }

        private void SendImpulseRestartAlert(TimeSpan elapsed)
        {
            try
            {
                string body =
                    "La macro Impulse 195 semble bloquée.<br/><br/>" +
                    "<b>Application :</b> " + global_application_name + "<br/>" +
                    "<b>Durée d'attente :</b> " +
                    elapsed.ToString(@"hh\:mm\:ss") + "<br/><br/>" +
                    "Merci de vérifier puis de relancer la macro 195.<br/>" +
                    "Le traitement SPLA continuera à attendre sa fin.";

                EnvoiEmail_with_Graph(
                    graphService,
                    global_application_name + " - Macro 195 à relancer",
                    body,
                    email_in_case_of_technical_issue);

                WriteToFile(
                    "       Macro 195 restart alert sent.");
            }
            catch (Exception ex)
            {
                WriteToFile(
                    "       Unable to send macro 195 restart alert : " +
                    ex.Message);
            }
        }

        private string ExportSplaResultsToXlsx(DataTable sourceSheet, string sourceFile)
        {
            const string selectSql = "SELECT * FROM T_SPLA_encodage ORDER BY id_ligne";
            DataTable results = ExecuteDataTableSql(nameof(ExportSplaResultsToXlsx), sql_connexion, selectSql);

            using (XLWorkbook workbook = new XLWorkbook())
            {
                IXLWorksheet sheet = workbook.Worksheets.Add("SPLA");
                CopyDataTableToWorksheet(sourceSheet, sheet);

                sheet.Cell(1, 21).Value = "Top Traité";
                sheet.Cell(1, 22).Value = "N° BL";
                sheet.Cell(1, 23).Value = "N° Ligne Impulse";
                sheet.Cell(1, 24).Value = "Commentaires";
                sheet.Range(1, 21, 1, 24).Style.Font.Bold = true;

                for (int rowIndex = 1; rowIndex < sourceSheet.Rows.Count; rowIndex++)
                {
                    DataRow sourceRow = sourceSheet.Rows[rowIndex];
                    if (string.IsNullOrWhiteSpace(CellText(sourceRow, 0)))
                    {
                        break;
                    }

                    DataRow result = FindSplaResult(results, sourceRow);
                    if (result == null)
                    {
                        continue;
                    }

                    int excelRow = rowIndex + 1;
                    sheet.Cell(excelRow, 21).Value = DbText(result, "top_traite");
                    sheet.Cell(excelRow, 22).Value = DbText(result, "N° BL");
                    sheet.Cell(excelRow, 23).Value = DbText(result, "num_ligne_impulse");
                    sheet.Cell(excelRow, 24).Value = DbText(result, "commentaires");
                }

                sheet.Columns(21, 24).AdjustToContents();
                string outputName = Path.GetFileNameWithoutExtension(sourceFile) + "_TRAITE.xlsx";
                string resultPath = Path.Combine(temp_folder,
                    global_application_name + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + "_" + outputName);
                workbook.SaveAs(resultPath);
                return resultPath;
            }
        }

        private static void CopyDataTableToWorksheet(DataTable source, IXLWorksheet target)
        {
            for (int rowIndex = 0; rowIndex < source.Rows.Count; rowIndex++)
            {
                for (int columnIndex = 0; columnIndex < source.Columns.Count; columnIndex++)
                {
                    object value = source.Rows[rowIndex][columnIndex];
                    if (value == null || value == DBNull.Value)
                    {
                        continue;
                    }

                    IXLCell cell = target.Cell(rowIndex + 1, columnIndex + 1);
                    if (value is DateTime date)
                    {
                        cell.Value = date;
                        cell.Style.DateFormat.Format = "dd/MM/yyyy";
                    }
                    else if (value is double number)
                    {
                        cell.Value = number;
                    }
                    else
                    {
                        cell.Value = Convert.ToString(value);
                    }
                }
            }
        }

        private static DataRow FindSplaResult(DataTable table, DataRow row)
        {
            string reportYear = CellText(row, 0);
            string reportMonth = CellText(row, 1);
            string branch = CellText(row, 2);
            string agreement = CellText(row, 3);
            string customerPo = CellText(row, 5);
            string email = CellText(row, 8);
            string sku = CellText(row, 14);
            string qty = CellText(row, 17);

            return table.AsEnumerable().FirstOrDefault(result =>
                Same(result, "ReportYear", reportYear) &&
                Same(result, "ReportMonth", reportMonth) &&
                Same(result, "BranchCustomerNbr", branch) &&
                Same(result, "SKU", sku) &&
                Same(result, "QTY", qty) &&
                (string.IsNullOrWhiteSpace(agreement) || Same(result, "AgreementNbr", agreement)) &&
                (string.IsNullOrWhiteSpace(customerPo)
                    ? Same(result, "EndcustomerEmail", email)
                    : Same(result, "CustomerPo", customerPo)));
        }

        private static bool Same(DataRow row, string column, string value)
        {
            return string.Equals(DbText(row, column), (value ?? "").Trim(), StringComparison.OrdinalIgnoreCase);
        }

        private static string DbText(DataRow row, string column)
        {
            return row.Table.Columns.Contains(column) && row[column] != DBNull.Value
                ? Convert.ToString(row[column]).Trim()
                : "";
        }

        private static string CellText(DataRow row, int columnIndex)
        {
            if (row == null || columnIndex < 0 || columnIndex >= row.Table.Columns.Count)
            {
                return "";
            }

            object value = row[columnIndex];
            if (value == null || value == DBNull.Value)
            {
                return "";
            }

            if (value is DateTime date)
            {
                return date.ToString("dd/MM/yyyy", CultureInfo.GetCultureInfo("fr-FR"));
            }

            return Convert.ToString(value, CultureInfo.GetCultureInfo("fr-FR")).Trim();
        }

        private static object CellValue(DataRow row, int columnIndex)
        {
            if (row == null ||
                columnIndex < 0 ||
                columnIndex >= row.Table.Columns.Count)
            {
                return null;
            }

            object value = row[columnIndex];

            return value == null || value == DBNull.Value
                ? null
                : value;
        }

        private static SqlParameter DecimalParameter(string name, object value)
        {
            SqlParameter parameter = new SqlParameter(name, SqlDbType.Decimal)
            {
                Precision = 18,
                Scale = 4

            };

            if (value == null || value == DBNull.Value)
            {
                parameter.Value = DBNull.Value;
                return parameter;
            }

            if (value is decimal decimalValue)
            {
                parameter.Value = decimalValue;
                return parameter;
            }

            if (value is double doubleValue)
            {
                parameter.Value = Convert.ToDecimal(doubleValue);
                return parameter;
            }

            string text = Convert.ToString(value)?.Trim() ?? "";

            if (decimal.TryParse(
                text,
                NumberStyles.Any,
                CultureInfo.GetCultureInfo("fr-FR"),
                out decimal result))
            {
                parameter.Value = result;
                return parameter;
            }

            if (decimal.TryParse(
                text,
                NumberStyles.Any,
                CultureInfo.InvariantCulture,
                out result))
            {
                parameter.Value = result;
                return parameter;
            }

            throw new FormatException(
                "Valeur décimale incorrecte pour " + name + " : " + text);
        }

        private int ExecuteNonQuerySql(string methodName, string connectionString, string sql,
            params SqlParameter[] parameters)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    return ExecuteNonQuery(connection, null, sql, parameters);
                }
            }
            catch (Exception ex)
            {
                SendSqlTechnicalIssueMail(methodName, sql, ex);
                throw;
            }
        }

        private object ExecuteScalarSql(string methodName, string connectionString, string sql,
            params SqlParameter[] parameters)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                using (SqlCommand command = new SqlCommand(sql, connection))
                {
                    command.CommandTimeout = 300;
                    if (parameters != null) command.Parameters.AddRange(parameters);
                    connection.Open();
                    return command.ExecuteScalar();
                }
            }
            catch (Exception ex)
            {
                SendSqlTechnicalIssueMail(methodName, sql, ex);
                throw;
            }
        }

        private DataTable ExecuteDataTableSql(string methodName, string connectionString, string sql,
            params SqlParameter[] parameters)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                using (SqlCommand command = new SqlCommand(sql, connection))
                using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                {
                    command.CommandTimeout = 300;
                    if (parameters != null) command.Parameters.AddRange(parameters);
                    DataTable table = new DataTable();
                    adapter.Fill(table);
                    return table;
                }
            }
            catch (Exception ex)
            {
                SendSqlTechnicalIssueMail(methodName, sql, ex);
                throw;
            }
        }

        private int ExecuteNonQuery(SqlConnection connection, SqlTransaction transaction, string sql,
            params SqlParameter[] parameters)
        {
            using (SqlCommand command = new SqlCommand(sql, connection, transaction))
            {
                command.CommandTimeout = 300;
                if (parameters != null) command.Parameters.AddRange(parameters);
                return command.ExecuteNonQuery();
            }
        }

        private static SqlParameter P(string name, object value)
        {
            return new SqlParameter(name, value ?? DBNull.Value);
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

        private void EnvoiEmail_with_Graph(GraphServiceClient graphService, string subject, string body, string recipient, string attachmentPath = "")
        {
            if (string.IsNullOrWhiteSpace(recipient) || !recipient.Contains("@"))
            {
                return;
            }

            List<Recipient> toRecipients = BuildRecipients(recipient);
            //  List<Recipient> ccRecipients = BuildRecipients(email_cc);

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

            /* if (ccRecipients.Count > 0)
             {
                 message.CcRecipients = ccRecipients;
             }*/

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

        private void ValidateRequiredParameters()
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

            if (string.IsNullOrWhiteSpace(temp_folder))
            {
                throw new Exception("temp_folder is empty");
            }

            if (string.IsNullOrWhiteSpace(sql_connexion))
            {
                throw new Exception("sql_connexion is empty. Parameter used : " + sql_connexion_parameter_global);
            }

            if (string.IsNullOrWhiteSpace(sql_connexion_macros))
            {
                throw new Exception("sql_connexion_macros is empty. Parameter used : " + sql_macros_parameter_global);
            }

            if (string.IsNullOrWhiteSpace(email_destinataire))
            {
                throw new Exception("email_destinataire is empty");
            }

            if (string.IsNullOrWhiteSpace(email_in_case_of_technical_issue))
            {
                throw new Exception("email_in_case_of_technical_issue is empty. Parameter used : " + email_in_case_of_technical_issue_parameter_global);
            }
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

        private string get_IMCA_paramters(string sql_con, string param_name)
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

                    cmd.Parameters.AddWithValue("@PARAMETER", param_name);

                    object result = cmd.ExecuteScalar();

                    if (result != null)
                    {
                        ret = result.ToString();
                    }
                }

                con.Close();
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
                "IMCA_" + global_session_name + "_" + DateTime.Now.ToString("dd_MM_yyyy") + "_" + country + "_" + global_application_name + ".txt"
            );

            File.AppendAllText(
                filePath,
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - " + message + Environment.NewLine
            );
        }

        private string CleanFileName(string fileName)
        {
            return string.Join("_", fileName.Split(Path.GetInvalidFileNameChars())).Trim();
        }

        private string EscapeODataString(string value)
        {
            if (value == null)
            {
                return "";
            }

            return value.Replace("'", "''");
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
                    if (Path.GetFileName(file).ToLower().StartsWith(prefixfile.ToLower()))
                    {
                        File.Delete(file);
                    }

                }
                catch
                {
                    // On ignore pour ne pas bloquer le service.
                }
            }
        }

        private void SendSqlTechnicalIssueMail(string methodName, string sql, Exception ex)
        {
            try
            {
                WriteToFile("SQL error in " + methodName + " : " + ex.Message);

                if (graphService == null)
                {
                    WriteToFile("Unable to send SQL technical issue email because GraphServiceClient is null");
                    return;
                }

                if (string.IsNullOrWhiteSpace(email_in_case_of_technical_issue) || !email_in_case_of_technical_issue.Contains("@"))
                {
                    WriteToFile("Unable to send SQL technical issue email because recipient is empty");
                    return;
                }

                string subject = global_application_name + " - Erreur SQL dans " + methodName;

                string body =
                    "Une erreur SQL est survenue dans " + global_application_name + ".<br/><br/>" +
                    "<b>Méthode :</b> " + methodName + "<br/>" +
                    "<b>Timeout configuré :</b> " + 300 + " secondes<br/>" +
                    "<b>Message :</b> " + ex.Message + "<br/><br/>" +
                    "<b>Requête SQL :</b><br/>" +
                    "<pre>" + sql + "</pre>";

                EnvoiEmail_with_Graph(
                    graphService,
                    subject,
                    body,
                    email_in_case_of_technical_issue
                );
            }
            catch (Exception mailEx)
            {
                WriteToFile("Error sending SQL technical issue email : " + mailEx.Message);
            }
        }

        public void setParamCountry(string country)
        {
            this.country = country ?? "";
        }

        public void setParamSK_Valid(string sk_valid)
        {
            this.sk_valid = sk_valid ?? "";
        }

        public void setParamName(string name)
        {
            this.name = name ?? "";
        }

        public void setParamActive(string active)
        {
            this.active = active ?? "";
        }

        public void setParamDebug(string debug)
        {
            this.debug = debug ?? "";
        }

        public void setStartDateScan(string start_date_scan)
        {
            this.start_date_scan = start_date_scan ?? "";
        }

        public void setNumber_of_mails(string number_of_mails)
        {
            this.number_of_mails = string.IsNullOrWhiteSpace(number_of_mails) ? "10" : number_of_mails;
        }

        public void setSharedMailboxName(string sharedmailbox_name)
        {
            this.sharedmailbox_name = sharedmailbox_name ?? "";
        }

        public void setsharedmailbox_folder_in(string sharedmailbox_folder_in)
        {
            this.sharedmailbox_folder_in = string.IsNullOrWhiteSpace(sharedmailbox_folder_in) ? "INBOX" : sharedmailbox_folder_in;
        }

        public void setsharedmailbox_folder_out(string sharedmailbox_folder_out)
        {
            this.sharedmailbox_folder_out = string.IsNullOrWhiteSpace(sharedmailbox_folder_out) ? "Archives" : sharedmailbox_folder_out;
        }

        public void setEmailInCaseOfTechnicalIssue(string email_in_case_of_technical_issue)
        {
            this.email_in_case_of_technical_issue = email_in_case_of_technical_issue ?? "";
        }


        public void setEmailDestinataire(string email_destinataire)
        {
            this.email_destinataire = email_destinataire ?? "";
        }

        public void setlogs_folder(string logs_folder)
        {
            this.logs_folder = logs_folder ?? "";
        }

        public void settemp_folder(string temp_folder)
        {
            this.temp_folder = temp_folder ?? "";
        }

        public void setSqlConnexion(string sql_connexion)
        {
            this.sql_connexion = sql_connexion ?? "";
        }


        public void setSqlConnexionParam(string sql_connexion_parameter_global)
        {
            this.sql_connexion_parameter_global = sql_connexion_parameter_global ?? "";
        }


        public void setEmailInCaseOfTechnicalIssueParam(string email_in_case_of_technical_issue_parameter_global)
        {
            this.email_in_case_of_technical_issue_parameter_global = email_in_case_of_technical_issue_parameter_global ?? "";
        }


        public void setSqlConnexionMacrosParam(string sql_macros_parameter_global)
        {
            this.sql_macros_parameter_global = sql_macros_parameter_global ?? "";
        }

        public void setSqlMacros(string sql_connexion_macros)
        {
            this.sql_connexion_macros = sql_connexion_macros ?? "";
        }

        public void setTimerAttenteImpulse(string timer_attente_impulse)
        {
            this.timer_attente_impulse = timer_attente_impulse ?? "";
        }


    }
}
