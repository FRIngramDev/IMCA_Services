using Microsoft.Graph;
using Microsoft.Graph.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using System.Net;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GESTION_REQUETE_FR
{
    /// <summary>
    /// Classe principale responsable de la lecture, du traitement et de l'envoi des emails
    /// pour la fonctionnalité GESTION_REQUETE_FR.
    /// Contient les paramètres de configuration, les helpers SQL / fichiers temporaires
    /// et le client Microsoft Graph utilisé pour accéder à la boîte partagée.
    /// </summary>
    public class GESTION_REQUETE_FR
    {
        /// <summary>Constructeur par défaut (ne fait rien pour l'instant).</summary>
        public GESTION_REQUETE_FR()
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

        // ----- Fichiers / logs / session -----
        private string logs_folder = "";                   // Répertoire des logs
        private string temp_folder = "";                   // Répertoire temporaire pour fichiers
        private string global_session_name = "";           // Nom de session pour logs et opérations

        // ----- Constantes / noms -----
        private string global_application_name = "GESTION_REQUETE_FR"; // Nom de l'application utilisé dans les logs

        // ----- Connexions SQL (chaînes / paramètres) -----
        private string sql_connexion = "";                         // Connexion principale (résolue dynamiquement)
        private string sql_connexion_parameter_global = "";        // Nom du paramètre contenant la connexion
        private string sql_connexion_annuaire = "";                // Connexion annuaire
        private string sql_annuaire_parameter_global = "";         // Nom du param annuaire
        private string sql_connexion_incentives = "";              // Connexion incentives
        private string sql_incentives_parameter_global = "";       // Nom du param incentives

        // ----- Client Microsoft Graph (utilisé pour lecture/envoi emails) -----
        private GraphServiceClient graphService = null;



        /// <summary>
        /// Représentation du fichier JSON de configuration attendu.
        /// Contient le dossier de logs global et la liste des pays configurés.
        /// </summary>
        public class JSON_file
        {
            public string logs_folder { get; set; } = "";                 // Répertoire des logs défini dans le JSON
            public List<Country> countries { get; set; } = new List<Country>(); // Liste des entrées pays
        }

        /// <summary>
        /// Représente la configuration d'un seul pays / instance telle que
        /// fournie dans le JSON de paramètres.
        /// Chaque propriété correspond à un paramètre utilisable par la classe.
        /// </summary>
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

            public string sql_annuaire_parameter_global { get; set; } = "";         // Param nom chaine SQL annuaire

            public string sql_incentives_parameter_global { get; set; } = "";       // Param nom chaine SQL incentives
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


                    setSqlConnexionParam(p.sql_connexion_parameter_global);
                    setSqlConnexionAnnuaireParam(p.sql_annuaire_parameter_global);
                    setSqlConnexionIncentivesParam(p.sql_incentives_parameter_global);

                    setEmailInCaseOfTechnicalIssueParam(p.email_in_case_of_technical_issue_parameter_global);

                    email_in_case_of_technical_issue = get_IMCA_paramters(sql_con, email_in_case_of_technical_issue_parameter_global);

                    sql_connexion = get_IMCA_paramters(sql_con, sql_connexion_parameter_global);
                    sql_connexion_annuaire = get_IMCA_paramters(sql_con, sql_annuaire_parameter_global);
                    sql_connexion_incentives = get_IMCA_paramters(sql_con, sql_incentives_parameter_global);



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
                                    ProcessGestionRequeteEmail(graphService, email);
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
                                            "Gestion Requêtes FR",
                                            "Gestion Requêtes FR - Une erreur (" + ex.Message + ") est survenue lors du traitement d'un fichier. Le mail a été déplacé dans le dossier Erreur.",
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
        /// Traite un message provenant de la boîte partagée pour la fonctionnalité GESTION_REQUETE_FR.
        /// 
        /// Etapes réalisées par cette méthode (sans modifier la logique existante) :
        /// 1) Récupère le message complet via Microsoft Graph avec les champs nécessaires (sujet, expéditeur, cc, corps, pièces jointes).
        /// 2) Normalise et extrait les informations utiles (sujet, expéditeur, destinataires en copie, corps, numéro de demande).
        /// 3) Si le message est une réponse à une "Demande d'informations", met à jour le statut via SetInformationsReceived.
        /// 4) Détecte des cas métiers spécifique Incentive et intègre la requête externe en appelant IntegrateExternalRequest.
        ///    - Si un identifiant externe est présent, exécute une mise à jour SQL pour lier l'entité externe à la requête interne.
        /// 5) Sauvegarde le message au format EML et téléverse le .eml ainsi que les pièces jointes en base via UploadFile.
        /// 
        /// Notes :
        /// - La méthode laisse inchangée la logique métier existante ; elle ajoute uniquement des commentaires explicatifs.
        /// - Les appels réseau/SQL conservent les mêmes comportements d'origine (exceptions remontées et gestion supérieure).
        /// </summary>
        /// <param name="service">Instance GraphServiceClient déjà authentifiée pour accéder à la boîte partagée.</param>
        /// <param name="summary">Objet Message minimal récupéré lors de la collecte (contient l'Id du message).</param>
        private void ProcessGestionRequeteEmail(GraphServiceClient service, Message summary)
        {
            // Récupération du message complet avec les champs nécessaires via Graph
            Message email = service.Users[sharedmailbox_name].Messages[summary.Id]
                .GetAsync(c => c.QueryParameters.Select = new string[]
                { "id", "subject", "from", "ccRecipients", "body", "hasAttachments" })
                .GetAwaiter().GetResult();

            // Normalisation des valeurs lues
            string subject = email.Subject ?? "";
            string sender = email.From?.EmailAddress?.Address ?? "";
            string cc = string.Join(";", (email.CcRecipients ?? new List<Recipient>())
                .Where(x => !string.IsNullOrWhiteSpace(x.EmailAddress?.Address))
                .Select(x => x.EmailAddress.Address));
            string body = HtmlToText(email.Body?.Content ?? "");

            // Extraction du numéro de demande (s'il est présent dans le sujet)
            int requestId = ExtractNumber(subject, @"Demande\s+de\s+Requête\s+N[°o]?\s*(\d+)");
            // Résolution de l'utilisateur (id interne) à partir de l'adresse expéditeur
            int userId = GetUserId(sender);

            // Si c'est une réponse ("RE") à une demande d'informations, on marque les informations comme reçues
            if (subject.IndexOf("RE", StringComparison.OrdinalIgnoreCase) >= 0 &&
                subject.IndexOf("Demande d'informations", StringComparison.OrdinalIgnoreCase) >= 0 && requestId > 0)
            {
                SetInformationsReceived(requestId, userId);
            }

            // Cas métier : Critères statistiques - Incentive
            if (subject.IndexOf("Critères statistiques - Incentive n°", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // Le demandeur peut être l'expéditeur si le sujet commence par "RE:", sinon on prend l'adresse en CC.
                string requester = subject.StartsWith("RE:", StringComparison.OrdinalIgnoreCase) ? sender : cc;
                userId = GetUserId(FirstEmail(requester));
                int externalId = ExtractNumber(subject, @"Incentive[^0-9]*(\d+)");

                // Intégration ou mise à jour de la requête externe et récupération de l'ID de requête interne
                requestId = IntegrateExternalRequest(subject, body, requester, userId, 1, "INCENTIVE", "Récurrente", externalId);

                // Si l'enregistrement externe existe, on lie l'incentive à la requête interne via une mise à jour SQL
                if (externalId > 0)
                {
                    ExecuteNonQuerySql(nameof(ProcessGestionRequeteEmail), sql_connexion_incentives,
                        "UPDATE T_incentives SET id_requete=@request WHERE id_incentive=@external",
                        P("@request", requestId), P("@external", externalId));
                }
            }

            // Sauvegarde du message en .eml sur disque temporaire puis upload en base
            string emlPath = SaveMessageAsEml(service, email.Id);
            UploadFile(emlPath, userId, requestId);

        }

        private string SaveMessageAsEml(GraphServiceClient service, string messageId)
        {
            string path = Path.Combine(temp_folder,
                global_application_name + "_EMAIL_" + DateTime.Now.ToString("yyyyMMdd_HHmmssfff") + ".eml");
            try
            {
                using (Stream mime = service.Users[sharedmailbox_name].Messages[messageId]
                    .Content.GetAsync().GetAwaiter().GetResult())
                using (FileStream output = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    mime.CopyTo(output);
                }
                return path;
            }
            catch (Exception ex)
            {
                WriteToFile("Unable to save EML for message " + messageId + " : " + ex.Message);
                throw;
            }
        }

        private int IntegrateExternalRequest(string subject, string body, string recipient, int userId,
            int typeId, string marker, string requestType, int externalId)
        {
            const string findSql = @"SELECT TOP 1 id_dde,id_gestionnaire FROM T_RequeteurGestion_Demande
                WHERE id_type_dde=@type AND nom_req LIKE @name ORDER BY id_dde DESC";
            int requestId = 0, managerId = 0;
            using (SqlConnection con = new SqlConnection(sql_connexion))
            {
                try
                {
                    con.Open();
                    using (SqlCommand cmd = new SqlCommand(findSql, con))
                    {
                        cmd.CommandTimeout = 300; cmd.Parameters.AddRange(new[] { P("@type", typeId), P("@name", "%" + marker + "%" + externalId + "%") });
                        using (SqlDataReader dr = cmd.ExecuteReader()) if (dr.Read()) { requestId = Convert.ToInt32(dr["id_dde"]); managerId = Convert.ToInt32(dr["id_gestionnaire"] == DBNull.Value ? 0 : dr["id_gestionnaire"]); }
                    }
                }
                catch (Exception ex) { SendSqlTechnicalIssueMail(nameof(IntegrateExternalRequest), findSql, ex); throw; }
            }
            string proc = requestId == 0 ? "dbo.USP_GESTIONREQUETE_ADD_DEMANDE" : "dbo.USP_GESTIONREQUETE_UPDATE_DEMANDE";
            var ps = new List<SqlParameter> { P("@id_user_ddeur", userId), P("@id_statut", requestId == 0 ? 10 : 20), P("@nom_req", subject.ToUpperInvariant()), P("@type_dde", requestType), P("@frequence", "Weekly"), P("@frequence_weekly", ""), P("@mois_im_calend", "Pas de Date"), P("@periode_im_calend", ExtractPeriod(body)), P("@branche_client", ""), P("@branche_produit", ""), P("@destinataire", "Interne"), P("@destinataire_email", recipient ?? ""), P("@is_usage_externe", false), P("@format", "Excel"), P("@commentaire", body ?? ""), P("@date_souhaitee", DateTime.Today.AddDays(8)) };
            if (requestId == 0)
            {
                ps.Add(P("@id_type_dde", typeId)); var output = new SqlParameter("@@id_dde", SqlDbType.Int) { Direction = ParameterDirection.Output }; ps.Add(output);
                ExecuteStoredProcedureSql(nameof(IntegrateExternalRequest), sql_connexion, proc, ps.ToArray()); return Convert.ToInt32(output.Value);
            }
            ps.Add(P("@id_dde", requestId)); ps.Add(P("@id_gestionnaire", managerId)); ps.Add(P("@is_reouverture", true));
            ExecuteStoredProcedureSql(nameof(IntegrateExternalRequest), sql_connexion, proc, ps.ToArray()); return requestId;
        }

        private void SetInformationsReceived(int requestId, int userId)
        {
            const string sql = @"UPDATE T_RequeteurGestion_Demande SET id_statut=35 WHERE id_dde=@id;
INSERT INTO T_RequeteurGestion_Histo(id_dde,histo_date,id_user,id_statut,comment,id_gestionnaire)
SELECT @id,GETDATE(),@user,35,@comment,ISNULL(id_gestionnaire,0) FROM T_RequeteurGestion_Demande WHERE id_dde=@id;";
            ExecuteNonQuerySql(nameof(SetInformationsReceived), sql_connexion, sql, P("@id", requestId), P("@user", userId), P("@comment", "L'utilisateur a répondu à la demande d'informations"));
        }

        private int GetUserId(string email)
        {
            if (string.IsNullOrWhiteSpace(email)) return 0;
            try { var a = new class_dev_tools.DroitAnnuaire(sql_connexion_annuaire); a.Get_User_Info_By_MailAddress(email); return a.Id_User; }
            catch (Exception ex) { SendSqlTechnicalIssueMail(nameof(GetUserId), "DroitAnnuaire.Get_User_Info_By_MailAddress", ex); throw; }
        }

        /// <summary>
        /// Télécharge les pièces jointes d'un message et les enregistre dans le dossier temporaire.
        /// Retourne la liste des chemins vers les fichiers téléchargés.
        /// </summary>
        private List<string> DownloadAttachments(GraphServiceClient service, string messageId)
        {
            var files = new List<string>(); var r = service.Users[sharedmailbox_name].Messages[messageId].Attachments.GetAsync().GetAwaiter().GetResult();
            foreach (Attachment a in r?.Value ?? new List<Attachment>()) if (a is FileAttachment f && f.ContentBytes != null && !string.IsNullOrWhiteSpace(f.Name)) { string path = Path.Combine(temp_folder, global_application_name + "_" + CleanFileName(f.Name)); File.WriteAllBytes(path, f.ContentBytes); files.Add(path); }
            return files;
        }

        private void UploadFile(string path, int userId, int requestId)
        {
            if (requestId <= 0 || !File.Exists(path)) return;
            ExecuteStoredProcedureSql(nameof(UploadFile), sql_connexion, "dbo.USP_GESTIONREQUETE_ADD_FIC", P("@id_dde", requestId), P("@nom_fic", Path.GetFileNameWithoutExtension(path)), new SqlParameter("@document", SqlDbType.VarBinary, -1) { Value = File.ReadAllBytes(path) }, P("@extension", Path.GetExtension(path).TrimStart('.').ToLowerInvariant()), P("@id_user", userId));
        }

        private int ExecuteNonQuerySql(string method, string connection, string sql, params SqlParameter[] ps)
        { 
            try 
            { 
                using (var con = new SqlConnection(connection)) 
                { con.Open(); 
                    using (var cmd = new SqlCommand(sql, con)) 
                    { cmd.CommandTimeout = 300; 
                        if (ps != null) cmd.Parameters.AddRange(ps); 
                        return cmd.ExecuteNonQuery(); 
                    } 
                } 
            } catch (Exception ex) 
            { 
                SendSqlTechnicalIssueMail(method, sql, ex); throw; 
            } 
        }

        private int ExecuteStoredProcedureSql(string method, string connection, string proc, params SqlParameter[] ps)
        { 
            try 
            { 
                using (var con = new SqlConnection(connection)) 
                { con.Open(); 
                    using (var cmd = new SqlCommand(proc, con)) 
                    { cmd.CommandType = CommandType.StoredProcedure; 
                        cmd.CommandTimeout = 300; if (ps != null) cmd.Parameters.AddRange(ps);
                        return cmd.ExecuteNonQuery(); 
                    } 
                } 
            } 
            catch (Exception ex) 
            { 
                SendSqlTechnicalIssueMail(method, proc, ex); throw; 
            } 
        }

        private static SqlParameter P(string name, object value) => new SqlParameter(name, value ?? DBNull.Value);
        private int ExtractNumber(string text, string pattern) 
        { 
            var m = Regex.Match(text ?? "", pattern, RegexOptions.IgnoreCase); 
            return m.Success && int.TryParse(m.Groups[1].Value, out int n) ? n : 0; 
        }
        private string ExtractPeriod(string body) 
        { 
            var d = Regex.Match(body ?? "", @"DATE DE DEBUT[^0-9]*(\d{2}/\d{2}/\d{4})", RegexOptions.IgnoreCase); 
            var f = Regex.Match(body ?? "", @"DATE DE FIN[^0-9]*(\d{2}/\d{2}/\d{4})", RegexOptions.IgnoreCase); 
            return (d.Success ? d.Groups[1].Value : "") + " au " + (f.Success ? f.Groups[1].Value : ""); 
        }

        private string HtmlToText(string html) => WebUtility.HtmlDecode(Regex.Replace(Regex.Replace(html ?? "", "<br\\s*/?>", "\n", RegexOptions.IgnoreCase), "<[^>]+>", " ")).Trim();

        private string FirstEmail(string emails) => (emails ?? "").Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault()?.Trim() ?? "";

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

        private void EnvoiEmail_with_Graph(GraphServiceClient graphService, string subject, string body, string recipient)
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

            if (string.IsNullOrWhiteSpace(sql_connexion_annuaire))
            {
                throw new Exception("sql_connexion_annuaire is empty. Parameter used : " + sql_annuaire_parameter_global);
            }

            if (string.IsNullOrWhiteSpace(sql_connexion_incentives))
            {
                throw new Exception("sql_connexion_incentives is empty. Parameter used : " + sql_incentives_parameter_global);
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


        public void setSqlConnexionAnnuaireParam(string sql_annuaire_parameter_global)
        {
            this.sql_annuaire_parameter_global = sql_annuaire_parameter_global ?? "";
        }

        public void setSqlAnnuaire(string sql_connexion_annuaire)
        {
            this.sql_connexion_annuaire = sql_connexion_annuaire ?? "";
        }

        public void setSqlConnexionIncentivesParam(string sql_incentives_parameter_global)
        {
            this.sql_incentives_parameter_global = sql_incentives_parameter_global ?? "";
        }

        public void setSqlIncentives(string sql_connexion_incentives)
        {
            this.sql_connexion_incentives = sql_connexion_incentives ?? "";
        }

    }
}
