using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
    }
}
