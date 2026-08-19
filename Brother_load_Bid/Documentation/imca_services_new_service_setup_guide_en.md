# Guide for Setting Up a New Service in IMCA Services

## 1. Purpose

This document describes the steps required to add, configure, and deploy a new task in **IMCA Services**.

It covers:

- creating the IMCA action;
- configuring global parameters;
- updating JSON configuration files;
- linking the functions to execute;
- key checks before production deployment.

> [!WARNING]
> Do not store plain-text passwords in application JSON files. Connection strings and sensitive parameters must be centralized in `PCM_TAB_IMCA_PARAMETER_GLOBAL` or in an appropriate secure storage solution.

---

## 2. Pre-production attention points

### 2.1 Update IDEP configuration

> [!IMPORTANT]
> Update the `IDEP_FR` parameter in `PCM_TAB_IMCA_PARAMETER_GLOBAL`.

New naming convention to apply:

- replace the old JSON parameter `sql_con` with `sql_con_parameter_global`;
- replace the old JSON parameter `sql_annuaire` with `sql_annuaire_parameter_global`;
- use the following global parameters:
  - `FR_SQLCON_INTRASTAT` for the Intrastat SQL connection;
  - `FR_SQLCON_ANNUAIRE` for the Annuaire SQL connection.

Expected example in the `IDEP_FR` JSON:

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

### 2.2 Update Fortinet configuration

> [!IMPORTANT]
> Update the `FORTINET_BID_LOAD` parameter in `PCM_TAB_IMCA_PARAMETER_GLOBAL`.

The new naming convention must use the same name as the associated class or service:

```text
FORTINET_LOAD_BID
```

To do:

- update the existing JSON;
- check the associated function names in `PCM_TAB_IMCA_ACTION_FUNCTION`;
- ensure the application name used in the code matches the global parameter name.

---

### 2.3 Update `IMCA_services.json`

> [!IMPORTANT]
> Add the `temp_folder` parameter to `IMCA_services.json`.

Example:

```json
"temp_folder": "temp"
```

This folder is used as a temporary working directory by some services, especially for downloading or processing intermediate files.

---

## 3. Global parameters to create in production

The following parameters must be created in `PCM_TAB_IMCA_PARAMETER_GLOBAL` with `SK_VALID = 0`.

> [!WARNING]
> The values below must be filled with actual production values, but passwords must not be distributed in documentation. Use placeholders or a secrets vault where possible.

| PARAMETER | Description | Expected value example |
|---|---|---|
| `FR_SQLCON_FRFRDSS_OPENROWSET` | OpenRowset connection to FRFR DSS | `'server,port';'user';'***PASSWORD***'` |
| `FR_SQLCON_MEETCOMPS` | SQL connection to the Meetcomps database | `Data Source=...;Initial Catalog=Meetcomps;User ID=...;Password=***PASSWORD***;Connection Timeout=0;language=french;` |
| `FR_EMAIL_IN_CASE_OF_TECHNICAL_ISSUE` | Technical email address for service errors | `programmeurs@ingrammicro.com` |
| `FR_SQLCON_INTRASTAT` | SQL connection to the INTRASTAT database | `Data Source=...;Initial Catalog=INTRASTAT;User ID=...;Password=***PASSWORD***;Connection Timeout=0` |
| `FR_SQLCON_ANNUAIRE` | SQL connection to the Annuaire database | `Data Source=...;Initial Catalog=Annuaire;User ID=...;Password=***PASSWORD***;Connection Timeout=0` |
| `FORTINET_BID_LOAD` or `FORTINET_LOAD_BID` | Fortinet service configuration JSON | Updated application JSON |
| `BROTHER_LOAD_BID` | Brother service configuration JSON | Updated application JSON |

---

## 4. Brother JSON example

Example JSON for the Brother service with sensitive parameters externalized:

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

## 5. Summary of the process for creating a new IMCA task

### 5.1 Create the action

Table: `PCM_TAB_IMCA_ACTION_ADMIN`

Create the task to be run in IMCA Services.

Typical fields to define:

- action name;
- description;
- active or inactive status;
- associated parameters if required.

---

### 5.2 Create or update global parameters

Table: `PCM_TAB_IMCA_PARAMETER_GLOBAL`

This table contains:

- the task configuration JSON;
- shared parameters;
- global parameter names;
- SQL connection strings;
- shared technical parameters.

Recommended naming convention:

```text
FR_SQLCON_<DATABASE_OR_USAGE>
FR_SQLCON_<DATABASE_OR_USAGE>_OPENROWSET
FR_EMAIL_IN_CASE_OF_TECHNICAL_ISSUE
<SERVICE_NAME>_SQL_COMMAND_TIMEOUT
```

Examples:

```text
FR_SQLCON_MEETCOMPS
FR_SQLCON_INTRASTAT
FR_SQLCON_ANNUAIRE
FR_SQLCON_FRFRDSS_OPENROWSET
FR_BROTHER_SQL_COMMAND_TIMEOUT
```

---

### 5.3 Define authorized users

Table: `PCM_TAB_IMCA_ACTION_USERVALIDATION`

This table defines the users authorized to launch the task.

In a test environment:

- use a non-production user;
- update `IMCA_services.json` in `bin\Debug` if needed;
- check that the account has the required permissions.

---

### 5.4 Define the functions to execute

Table: `PCM_TAB_IMCA_ACTION_FUNCTION`

Add the function or functions executed by the task.

Check the following:

- exact class name;
- exact method name;
- execution order;
- consistency with the global JSON parameter name.

Example:

```text
Action: SBO Import Brother FR
Function: Brother_load_Bid.Brother_load_Bid.Read_Email_with_Graph
Naming convention: [namespace].[class].[function]
Order: 1
```

---

### 5.5 Link external flags

Table: `PCM_TAB_IMCA_ACTION_FLAG`

Link the ID from `PCM_TAB_IMCA_ACTION_ADMIN` to the external component flag to check before starting the task.

Purpose:

- avoid concurrent processing;
- verify that an external component is available;
- prevent startup if a prerequisite is not validated.

---

### 5.6 Initialize task launch

Table: `PCM_TAB_IMCA_ACTION`

Add at least one record to initialize the task launch.

> [!IMPORTANT]
> The task must contain at least one initial record in `PCM_TAB_IMCA_ACTION` to be picked up by the IMCA Services engine.

---

## 6. Production deployment checklist

Before deployment, check the following:

- [ ] The action exists in `PCM_TAB_IMCA_ACTION_ADMIN`.
- [ ] The application JSON exists in `PCM_TAB_IMCA_PARAMETER_GLOBAL`.
- [ ] SQL global parameters are created with `SK_VALID = 0`.
- [ ] Connection strings are no longer stored directly in application JSON files.
- [ ] Authorized users are present in `PCM_TAB_IMCA_ACTION_USERVALIDATION`.
- [ ] Functions are declared in `PCM_TAB_IMCA_ACTION_FUNCTION`.
- [ ] Required flags are present in `PCM_TAB_IMCA_ACTION_FLAG`.
- [ ] An initialization record exists in `PCM_TAB_IMCA_ACTION`.
- [ ] `IMCA_services.json` contains the `temp_folder` parameter.
- [ ] Temporary folders exist or can be created by the service.
- [ ] Graph or Exchange permissions are validated for services that read or send emails.
- [ ] Technical emails are configured through `FR_EMAIL_IN_CASE_OF_TECHNICAL_ISSUE`.
- [ ] Logs are generated in the expected folder.

---

## 7. Best practices

### 7.1 Naming

Use a clear and stable naming convention:

```text
<SERVICE_NAME>
FR_SQLCON_<DATABASE>
FR_SQLCON_<DATABASE>_OPENROWSET
FR_EMAIL_IN_CASE_OF_TECHNICAL_ISSUE
```

### 7.2 Security

- avoid plain-text passwords in JSON files;
- avoid logging connection strings;
- mask sensitive values in error emails;
- limit recipients of technical emails;
- verify service account permissions.

### 7.3 Operations

- enable `debug = TRUE` only for testing or investigation;
- disable debug in production if logs become too large;
- keep logs long enough to diagnose errors;
- configure a reasonable SQL timeout, for example `300` seconds.

---

## 8. Brother-specific notes

For the Brother service:

- the service reads emails through Microsoft Graph;
- Excel attachments are downloaded into the temporary folder;
- files are imported into the Meetcomps database;
- emails are moved to `Archives` on success;
- emails are moved to `Erreur` on failure;
- technical alerts are sent to the address defined in `FR_EMAIL_IN_CASE_OF_TECHNICAL_ISSUE`;
- functional messages are sent to `email_to`.

---

## 9. IDEP-specific notes

For the IDEP service:

- `sql_con` must be replaced with `sql_con_parameter_global`;
- `sql_annuaire` must be replaced with `sql_annuaire_parameter_global`;
- actual values must be stored in `PCM_TAB_IMCA_PARAMETER_GLOBAL`;
- the JSON must contain only the names of global parameters.

---

## 10. Control query example

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

## 11. Change history to apply

- Externalize SQL connection strings from JSON files to `PCM_TAB_IMCA_PARAMETER_GLOBAL`.
- Add a `temp_folder` parameter to `IMCA_services.json`.
- Standardize global parameter names.
- Gradually migrate email reading to Microsoft Graph.
- Standardize technical email notifications through `FR_EMAIL_IN_CASE_OF_TECHNICAL_ISSUE`.
