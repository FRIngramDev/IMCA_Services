/!\\ ATTENTION MODIFIER PARAM IDEP DANS PARAMETER\_GLOBAL-->Nouvelle nomenclature FR\_SQLCON\_INTRASTAT pour sql\_con et sql\_annuaire FR\_SQLCON\_ANNUAIRE et mettre a jour le json de IDEP\_FR 



/!\\ ATTENTION MODIFIER PARAM FORTINET\_BID\_LOAD DANS PARAMETER\_GLOBAL-->Nouvelle nomenclature même nom que la classe FORTINET\_LOAD\_BID



/!\\ ATTENTION MODIFIER IMCA\_services.json car ajout de "temp\_folder":"temp",





liste des nouveaux parametre globaux a créer dans la prod 



FR\_SQLCON\_FRFRDSS\_OPENROWSET |	'defrwsql1025c04,1433';'usr\_read';'read'	

FR\_SQLCON\_MEETCOMPS   |          Data Source=DEFRWSQL1062D,1433;Initial Catalog=Meetcomps;User ID=usr\_meetcomp;Password=MIcOMp;Connection Timeout=0;language=french;

FR\_EMAIL\_IN\_CASE\_OF\_TECHNICAL\_ISSUE  |	programmeurs@ingrammicro.com

FR\_SQLCON\_INTRASTAT |	Data Source=DEFRWSQL1062D,1433;Initial Catalog=INTRASTAT;Trusted\_Connection=No;User ID=usr\_intrastat;Password=istik;Connection Timeout=0"

FR\_SQLCON\_ANNUAIRE  |	Data Source=DEFRWSQL1062D,1433;Initial Catalog=Annuaire;User ID=usr\_Annuaire;Password=ANnuR;Connection Timeout=0"











PCM\_TAB\_IMCA\_ACTION\_ADMIN --> créer la tâche (Action) à faire tourner  + param associés



PCM\_TAB\_IMCA\_PARAMETER\_GLOBAL --> contient le fichier Json avec les paramètres de la tâche + param commun comme chaine de connexion SQL ( respecter nomenclature FR\_CONSQL\_TABLE\_( openrowset)



PCM\_TAB\_IMCA\_ACTION\_USERVALIDATION -> définir le user habilité à lancer la tâche ( en test utiliser un user No prod et modifier IMCA\_service.json dans bin\\Debug)



PCM\_TAB\_IMCA\_ACTION\_FUNCTION --> ajouter la/les fonctions avec l'ordre d'exécution à respecter



PCM\_TAB\_IMCA\_ACTION\_FLAG --> associer l'ID de PCM\_TAB\_IMCA\_ACTION\_ADMIN au flag du composant externe à contrôler avant démarrage de la tâche 



PCM\_TAB\_IMCA\_ACTION --> Ajouter un enregistrement pour initialiser le lancement de la tâche ( doit contenir au moins 1 enr) 

























