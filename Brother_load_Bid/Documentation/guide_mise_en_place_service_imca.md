# Guide de mise en place d'un nouveau service dans IMCA Services

## 1. Objectif

Ce document décrit les étapes à suivre pour ajouter, configurer et déployer une nouvelle tâche dans **IMCA Services**.

Il couvre notamment :

- la création de l'action IMCA ;
- la configuration des paramètres globaux ;
- la mise à jour des fichiers JSON ;
- l'association des fonctions à exécuter ;
- les points d'attention avant mise en production.

> [!WARNING]
> Ne pas stocker de mots de passe en clair dans les fichiers JSON applicatifs. Les chaînes de connexion et paramètres sensibles doivent être centralisés dans `PCM_TAB_IMCA_PARAMETER_GLOBAL` ou dans un stockage sécurisé adapté.

---

## 2. Points d'attention avant mise en production

### 2.1 Mise à jour du paramétrage IDEP

> [!IMPORTANT]
> Modifier le paramètre `IDEP_FR` dans `PCM_TAB_IMCA_PARAMETER_GLOBAL`.

Nouvelle nomenclature à appliquer :

- remplacer l'ancien paramètre JSON `sql_con` par `sql_con_parameter_global` ;
- remplacer l'ancien paramètre JSON `sql_annuaire` par `sql_annuaire_parameter_global` ;
- utiliser les paramètres globaux suivants :
  - `FR_SQLCON_INTRASTAT` pour la connexion SQL Intrastat ;
  - `FR_SQLCON_ANNUAIRE` pour la connexion SQL Annuaire.

Exemple attendu dans le JSON `IDEP_FR` :

```json
{
  "param": "sql_con_parameter_global",
  "valeur": "FR_SQLCON_INTRASTAT"
},
{
  "param": "sql_annuaire_parameter_global",
  "valeur": "FR_SQLCON_ANNUAIRE"
}
```

---

### 2.2 Mise à jour du paramétrage Fortinet

> [!IMPORTANT]
> Modifier le paramètre `FORTINET_BID_LOAD` dans `PCM_TAB_IMCA_PARAMETER_GLOBAL`.

La nouvelle nomenclature doit utiliser le même nom que la classe/service associé :

```text
FORTINET_LOAD_BID
```

À prévoir :

- mettre à jour le JSON existant ;
- vérifier les noms de fonctions associés dans `PCM_TAB_IMCA_ACTION_FUNCTION` ;
- vérifier que le nom applicatif utilisé dans le code correspond bien au paramètre global.

---

### 2.3 Mise à jour de `IMCA_services.json`

> [!IMPORTANT]
> Ajouter le paramètre `temp_folder` dans `IMCA_services.json`.

Exemple :

```json
"temp_folder": "temp"
```

Ce dossier est utilisé comme dossier temporaire de travail par certains services, notamment pour le téléchargement ou le traitement de fichiers intermédiaires.

---

## 3. Paramètres globaux à créer en production

Les paramètres suivants doivent être créés dans `PCM_TAB_IMCA_PARAMETER_GLOBAL` avec `SK_VALID = 0`.

> [!WARNING]
> Les valeurs ci-dessous doivent être renseignées avec les valeurs réelles en production, mais les mots de passe ne doivent pas être diffusés dans la documentation. Utiliser des placeholders ou un coffre de secrets lorsque possible.

| PARAMETER                                  | Description                                   | Exemple de valeur attendue                                                                                            |
| ------------------------------------------ | --------------------------------------------- | --------------------------------------------------------------------------------------------------------------------- |
| `FR_SQLCON_FRFRDSS_OPENROWSET`             | Connexion OpenRowset vers FRFR DSS            | `'serveur,port';'user';'***PASSWORD***'`                                                                              |
| `FR_SQLCON_MEETCOMPS`                      | Connexion SQL vers la base Meetcomps          | `Data Source=...;Initial Catalog=Meetcomps;User ID=...;Password=***PASSWORD***;Connection Timeout=0;language=french;` |
| `FR_EMAIL_IN_CASE_OF_TECHNICAL_ISSUE`      | Adresse technique pour les erreurs de service | `programmeurs@ingrammicro.com`                                                                                        |
| `FR_SQLCON_INTRASTAT`                      | Connexion SQL vers la base INTRASTAT          | `Data Source=...;Initial Catalog=INTRASTAT;User ID=...;Password=***PASSWORD***;Connection Timeout=0`                  |
| `FR_SQLCON_ANNUAIRE`                       | Connexion SQL vers la base Annuaire           | `Data Source=...;Initial Catalog=Annuaire;User ID=...;Password=***PASSWORD***;Connection Timeout=0`                   |
| `FORTINET_BID_LOAD` ou `FORTINET_LOAD_BID` | JSON de configuration du service Fortinet     | JSON applicatif à jour                                                                                                |
| `BROTHER_LOAD_BID`                         | JSON de configuration du service Brother      | JSON applicatif à jour                                                                                                |

---

## 4. Exemple de JSON Brother

Exemple de JSON pour le service Brother avec paramètres sensibles externalisés :

```json
{
  "countries": [
    {
      "country": "FR",
      "sk_valid": "9",
      "name": "FRANCE",
      "active": "TRUE",
      "debug": "TRUE",
      "number_of_mails": "10",
      "start_date_scan": "20260801",

      "sharedmailbox_name": "cotationsbrother_fr@ingrammicro.com",
      "sharedmailbox_folder_in": "INBOX",
      "sharedmailbox_folder_out": "Archives",

      "email_in_case_of_technical_issue_parameter_global": "FR_EMAIL_IN_CASE_OF_TECHNICAL_ISSUE",
      "email_to": "jeremy.merlier@ingrammicro.com",
      "email_cc": "",

      "sql_connexion_parameter_global": "FR_SQLCON_MEETCOMPS",
      "dss_con_openrowset_parameter_global": "FR_SQLCON_FRFRDSS_OPENROWSET"
    }
  ]
}
```

---

## 5. Récapitulatif du processus de création d'une nouvelle tâche IMCA

### 5.1 Créer l'action

Table : `PCM_TAB_IMCA_ACTION_ADMIN`

Créer la tâche à faire tourner dans IMCA Services.

À renseigner notamment :

- nom de l'action ;
- description ;
- statut actif/inactif ;
- paramètres associés si nécessaire.

---

### 5.2 Créer ou modifier les paramètres globaux

Table : `PCM_TAB_IMCA_PARAMETER_GLOBAL`

Cette table contient :

- le JSON de configuration de la tâche ;
- les paramètres communs ;
- les noms de paramètres globaux ;
- les chaînes de connexion SQL ;
- les paramètres techniques partagés.

Convention recommandée :

```text
FR_SQLCON_<BASE_OU_USAGE>
FR_SQLCON_<BASE_OU_USAGE>_OPENROWSET
FR_EMAIL_IN_CASE_OF_TECHNICAL_ISSUE
<NOM_SERVICE>_SQL_COMMAND_TIMEOUT
```

Exemples :

```text
FR_SQLCON_MEETCOMPS
FR_SQLCON_INTRASTAT
FR_SQLCON_ANNUAIRE
FR_SQLCON_FRFRDSS_OPENROWSET
FR_BROTHER_SQL_COMMAND_TIMEOUT
```

---

### 5.3 Définir les utilisateurs autorisés

Table : `PCM_TAB_IMCA_ACTION_USERVALIDATION`

Cette table permet de définir les utilisateurs habilités à lancer la tâche.

En environnement de test :

- utiliser un utilisateur non production ;
- modifier si nécessaire `IMCA_services.json` dans `bin\Debug` ;
- vérifier que le compte utilisé a bien les droits nécessaires.

---

### 5.4 Définir les fonctions à exécuter

Table : `PCM_TAB_IMCA_ACTION_FUNCTION`

Ajouter la ou les fonctions exécutées par la tâche.

Points à vérifier :

- nom exact de la classe ;
- nom exact de la méthode ;
- ordre d'exécution ;
- cohérence avec le nom du paramètre global JSON.

Exemple de principe :

```text
Action : SBO Import Brother FR
Fonction : Brother_load_Bid.Brother_load_Bid.Read_Email_with_Graph
Nomenclature : [namespace].[class].[function]
Ordre : 1
```

---

### 5.5 Associer les flags externes

Table : `PCM_TAB_IMCA_ACTION_FLAG`

Associer l'ID de `PCM_TAB_IMCA_ACTION_ADMIN` au flag du composant externe à contrôler avant le démarrage de la tâche.

Objectif :

- éviter les traitements concurrents ;
- vérifier qu'un composant externe est disponible ;
- empêcher le démarrage si un prérequis n'est pas validé.

---

### 5.6 Initialiser le lancement de la tâche

Table : `PCM_TAB_IMCA_ACTION`

Ajouter au moins un enregistrement pour initialiser le lancement de la tâche.

> [!IMPORTANT]
> La tâche doit contenir au moins un enregistrement initial dans `PCM_TAB_IMCA_ACTION` pour être prise en compte par le moteur IMCA Services.

---

## 7. Checklist de mise en production

Avant déploiement, vérifier les points suivants :

- [ ] L'action existe dans `PCM_TAB_IMCA_ACTION_ADMIN`.
- [ ] Le JSON applicatif existe dans `PCM_TAB_IMCA_PARAMETER_GLOBAL`.
- [ ] Les paramètres globaux SQL sont créés avec `SK_VALID = 0`.
- [ ] Les chaînes de connexion ne sont plus stockées directement dans les JSON applicatifs.
- [ ] Les utilisateurs autorisés sont présents dans `PCM_TAB_IMCA_ACTION_USERVALIDATION`.
- [ ] Les fonctions sont déclarées dans `PCM_TAB_IMCA_ACTION_FUNCTION`.
- [ ] Les flags nécessaires sont présents dans `PCM_TAB_IMCA_ACTION_FLAG`.
- [ ] Un enregistrement d'initialisation existe dans `PCM_TAB_IMCA_ACTION`.
- [ ] `IMCA_services.json` contient le paramètre `temp_folder`.
- [ ] Les dossiers temporaires existent ou peuvent être créés par le service.
- [ ] Les droits Graph ou Exchange sont validés pour les services qui lisent ou envoient des emails.
- [ ] Les emails techniques sont configurés via `FR_EMAIL_IN_CASE_OF_TECHNICAL_ISSUE`.
- [ ] Les logs sont générés dans le dossier attendu.

---

## 7. Bonnes pratiques

### 7.1 Nommage

Utiliser une nomenclature claire et stable :

```text
<NOM_SERVICE>
FR_SQLCON_<BASE>
FR_SQLCON_<BASE>_OPENROWSET
FR_EMAIL_IN_CASE_OF_TECHNICAL_ISSUE
```

### 7.2 Sécurité

- éviter les mots de passe en clair dans les JSON ;
- éviter de logger les chaînes de connexion ;
- masquer les valeurs sensibles dans les emails d'erreur ;
- limiter les destinataires des emails techniques ;
- vérifier les droits des comptes de service.

### 7.3 Exploitation

- activer `debug = TRUE` uniquement en test ou investigation ;
- désactiver le debug en production si les logs deviennent trop volumineux ;
- conserver les logs suffisamment longtemps pour diagnostiquer les erreurs ;
- prévoir un timeout SQL raisonnable, par exemple `300` secondes.

---

## 9. Notes spécifiques Brother

Pour le service Brother :

- le service lit les mails via Microsoft Graph ;
- les pièces jointes Excel sont téléchargées dans le dossier temporaire ;
- les fichiers sont importés en base Meetcomps ;
- les emails sont déplacés vers `Archives` en cas de succès ;
- les emails sont déplacés vers `Erreur` en cas d'échec ;
- les alertes techniques sont envoyées à l'adresse définie dans `FR_EMAIL_IN_CASE_OF_TECHNICAL_ISSUE` ;
- les messages fonctionnels sont envoyés à `email_to`.

---

## 10. Notes spécifiques IDEP

Pour le service IDEP :

- `sql_con` doit être remplacé par `sql_con_parameter_global` ;
- `sql_annuaire` doit être remplacé par `sql_annuaire_parameter_global` ;
- les valeurs réelles doivent être stockées dans `PCM_TAB_IMCA_PARAMETER_GLOBAL` ;
- le JSON ne doit contenir que les noms de paramètres globaux.

---

## 11. Exemple de requête de contrôle

```sql
SELECT SK_VALID,
       PARAMETER,
       VALUE
FROM PCM_TAB_IMCA_PARAMETER_GLOBAL
WHERE SK_VALID = 0
AND PARAMETER IN
(
    'FR_SQLCON_MEETCOMPS',
    'FR_SQLCON_INTRASTAT',
    'FR_SQLCON_ANNUAIRE',
    'FR_SQLCON_MONARCH',
    'FR_SQLCON_FRFRDSS_OPENROWSET',
    'FR_EMAIL_IN_CASE_OF_TECHNICAL_ISSUE',
    'BROTHER_LOAD_BID',
    'IDEP_FR',
    'FORTINET_LOAD_BID'
)
ORDER BY PARAMETER;
```

---

## 12. Historique des changements à appliquer

- Externalisation des chaînes SQL depuis les JSON vers `PCM_TAB_IMCA_PARAMETER_GLOBAL`.
- Ajout d'un paramètre `temp_folder` dans `IMCA_services.json`.
- Harmonisation des noms de paramètres globaux.
- Migration progressive des lectures mail vers Microsoft Graph.
- Harmonisation des emails techniques via `FR_EMAIL_IN_CASE_OF_TECHNICAL_ISSUE`.
