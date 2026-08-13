/!\ ATTENTION MODIFIER PARAM IDEP DANS PARAMETER_GLOBAL-->Nouvelle nomenclature FR_SQLCON_INTRASTAT

!\ ATTENTION MODIFIER PARAM FORTINET_BID_LOAD DANS PARAMETER_GLOBAL-->Nouvelle nomenclature même nom que la classe FORTINET_LOAD_BID



PCM_TAB_IMCA_ACTION_ADMIN --> créer la tâche (Action) à faire tourner  + param associés

PCM_TAB_IMCA_PARAMETER_GLOBAL --> contient le fichier Json avec les paramètres de la tâche + param commun comme chaine de connexion SQL ( respecter nomenclature FR_CONSQL_TABLE_( openrowset)

PCM_TAB_IMCA_ACTION_USERVALIDATION -> définir le user habilité à lancer la tâche ( en test utiliser un user No prod et modifier IMCA_service.json dans bin\Debug)

PCM_TAB_IMCA_ACTION_FUNCTION --> ajouter la/les fonctions avec l'ordre d'exécution à respecter

PCM_TAB_IMCA_ACTION_FLAG --> associer l'ID de PCM_TAB_IMCA_ACTION_ADMIN au flag du composant externe à contrôler avant démarrage de la tâche 

PCM_TAB_IMCA_ACTION --> Ajouter un enregistrement pour initialiser le lancement de la tâche ( doit contenir au moins 1 enr) 











