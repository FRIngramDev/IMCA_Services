using ClosedXML.Excel;
using Microsoft.VisualBasic;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Windows.Forms;
using System.Xml;
using System.Data.Odbc;

namespace Idep_FR
{
    public class Idep_FR
    {
        public string strRep_Travail;
        public string Nom_modele_MDB;

        public string rep_modele_monarch;
        public string NomFicModele_XML;

        public SqlConnection ConBase;
        public string strCon = "";
        public string sql_con_parameter_global = "";

        public string str_annuaire_con = "";
        public string sql_annuaire_parameter_global = "";

        public string strMonarch;
        public string sql_monarch_parameter_global = "";

        public string date_rap;
        public string rep_fichier_traites;
        public DataSet DS;
        private string logs_folder = "";
        private string global_session_name = "";
        private readonly string global_application_name = "IDEP_FR";

        // ----- Technical issue email parameters -----
        private string email_in_case_of_technical_issue_parameter_global = "";
        private string email_in_case_of_technical_issue = "";
        private int technical_email_is_sending = 0;
        private int timer_attente_monarch_minutes = 10; // Alert interval while Monarch is still running

        // Synchronizes concurrent writes to the IDEP log file.
        private static readonly object logLock = new object();


        public Idep_FR()
        {

        }
        public class parameters
        {
            public string param { get; set; }
            public string valeur { get; set; }
        }
        public class IDEP_FRJSON_file
        {
            public List<parameters> list_param { get; set; }
        }
        private string get_IMCA_paramters(string sql_con, string param_name)
        {
            const string sql = @"SELECT ISNULL(VALUE, '') AS VALUE
FROM [PCM_TAB_IMCA_PARAMETER_GLOBAL]
WHERE SK_VALID = 0 AND PARAMETER = @PARAMETER";

            try
            {
                using (SqlConnection con = new SqlConnection(sql_con))
                using (SqlCommand cmd = new SqlCommand(sql, con))
                {
                    cmd.CommandTimeout = 300;
                    cmd.Parameters.Add("@PARAMETER", SqlDbType.VarChar, 255).Value = param_name ?? "";
                    con.Open();
                    object result = cmd.ExecuteScalar();
                    return result == null || result == DBNull.Value
                        ? ""
                        : Convert.ToString(result);
                }
            }
            catch (Exception ex)
            {
                LogTechnicalError(nameof(get_IMCA_paramters), ex, sql);
                throw;
            }
        }
        private bool isFileOpen(string filename)
        {
            System.IO.FileStream sf = null;

            try
            {
                sf = System.IO.File.Open(filename, System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite, System.IO.FileShare.None);

                return false;
            }
            catch (IOException)
            {
                return true;
            }

            finally
            {
                if (!Information.IsNothing(sf))
                {
                    sf.Close();

                    sf = null;
                }
            }
        }

        public void setlogs_folder(string logs_folder)
        {
            this.logs_folder = logs_folder;
        }



        public string InsertSQL_with_odbc(string strSql, string strSqlBDD, OdbcConnection Connexion, string MODELE)
        {
            string result = "";
            LogStep(nameof(InsertSQL_with_odbc), "START", "Model=" + MODELE);
            OdbcCommand ocmd = new OdbcCommand();
            OdbcDataAdapter oda;
            DataTable table_access;
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter da;
            DataSet DataSetAll = new DataSet();
            DataRow newLigne;

            cmd.CommandText = strSqlBDD;
            cmd.Connection = ConBase;
            da = new SqlDataAdapter(cmd);
            if (MODELE == "PUP7902")
            {
                da.Fill(DataSetAll, "PUP7902");
            }
            else
            {
                da.Fill(DataSetAll, "ARP0292");
            }

            ocmd.CommandText = strSql;
            ocmd.Connection = Connexion;
            oda = new OdbcDataAdapter(ocmd);
            oda.Fill(DataSetAll, "tmp");
            table_access = DataSetAll.Tables["tmp"];

            // On traite les lignes

            foreach (DataRow ligne in table_access.Rows)
            {
                if (MODELE == "PUP7902")
                    newLigne = DataSetAll.Tables["PUP7902"].NewRow();
                else
                    newLigne = DataSetAll.Tables["ARP0292"].NewRow();

                object[] the_ligne = ligne.ItemArray;
                // Modifie le format de la colonne pour l'inserer dans SQL/SERVEUR
                if (MODELE == "PUP7902")
                {
                    //the_ligne[7] = the_ligne[7].ToString().Replace(".", "");
                    the_ligne[7] = the_ligne[7].ToString().Replace(",", "");
                    //the_ligne[8] = the_ligne[8].ToString().Replace(".", "");
                    the_ligne[8] = the_ligne[8].ToString().Replace(",", "");
                    //the_ligne[9] = the_ligne[9].ToString().Replace(".", "");
                    the_ligne[9] = the_ligne[9].ToString().Replace(",", "");
                    //the_ligne[14] = the_ligne[14].ToString().Replace(".", "");
                    the_ligne[14] = the_ligne[14].ToString().Replace(",", "");
                    newLigne.ItemArray = (object[])the_ligne;
                    DataSetAll.Tables["PUP7902"].Rows.Add(newLigne);
                }
                else
                {
                    if (Information.IsDBNull(the_ligne[4]) == false)
                    {
                        //the_ligne[4] = the_ligne[4].ToString().Replace(".", "");
                        //the_ligne[5] = the_ligne[5].ToString().Replace(".", "");
                        //the_ligne[8] = the_ligne[8].ToString().Replace(".", "");

                        the_ligne[4] = the_ligne[4].ToString().Replace(",", "");
                        the_ligne[5] = the_ligne[5].ToString().Replace(",", "");
                        the_ligne[8] = the_ligne[8].ToString().Replace(",", "");

                        if (Strings.Right(the_ligne[5].ToString(), 1) == "-")
                            // C'est un nombre negatif
                            the_ligne[5] = "-" + the_ligne[5].ToString().Replace("-", "").Trim();

                        if (Strings.Right(the_ligne[8].ToString(), 1) == "-")
                            // C'est un nombre negatif
                            the_ligne[8] = "-" + the_ligne[8].ToString().Replace("-", "").Trim();
                    }

                    newLigne.ItemArray = (object[])the_ligne;
                    DataSetAll.Tables["ARP0292"].Rows.Add(newLigne);
                }
            }

            SqlCommandBuilder dtr = new SqlCommandBuilder(da);
            if (MODELE == "PUP7902")
                da.Update(DataSetAll, "PUP7902");
            else
                da.Update(DataSetAll, "ARP0292");
            ocmd.Dispose();
            oda.Dispose();

            cmd.Dispose();
            da.Dispose();

            int sourceRowCount = table_access.Rows.Count;

            table_access.Dispose();

            DataSetAll.Clear();
            DataSetAll.Dispose();

            LogStep(nameof(InsertSQL_with_odbc), "END",
                "Model=" + MODELE + ", source rows=" + sourceRowCount);
            return result;
        }

        public string InsertSQL(string strSql, string strSqlBDD, OleDbConnection Connexion, string MODELE)
        {
            string result = "";
            LogStep(nameof(InsertSQL), "START", "Model=" + MODELE);
            OleDbCommand ocmd = new OleDbCommand();
            OleDbDataAdapter oda;
            DataTable table_access;
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter da;
            DataSet DataSetAll = new DataSet();
            DataRow newLigne;

            cmd.CommandText = strSqlBDD;
            cmd.Connection = ConBase;
            da = new SqlDataAdapter(cmd);
            if (MODELE == "PUP7902")
            {
                da.Fill(DataSetAll, "PUP7902");
            }
            else
            {
                da.Fill(DataSetAll, "ARP0292");
            }

            ocmd.CommandText = strSql;
            ocmd.Connection = Connexion;
            oda = new OleDbDataAdapter(ocmd);
            oda.Fill(DataSetAll, "tmp");
            table_access = DataSetAll.Tables["tmp"];

            // On traite les lignes

            foreach (DataRow ligne in table_access.Rows)
            {
                if (MODELE == "PUP7902")
                    newLigne = DataSetAll.Tables["PUP7902"].NewRow();
                else
                    newLigne = DataSetAll.Tables["ARP0292"].NewRow();

                object[] the_ligne = ligne.ItemArray;
                // Modifie le format de la colonne pour l'inserer dans SQL/SERVEUR
                if (MODELE == "PUP7902")
                {
                    the_ligne[7] = the_ligne[7].ToString().Replace(".", "");
                    the_ligne[8] = the_ligne[8].ToString().Replace(".", "");
                    the_ligne[9] = the_ligne[9].ToString().Replace(".", "");
                    the_ligne[14] = the_ligne[14].ToString().Replace(".", "");
                    newLigne.ItemArray = (object[])the_ligne;
                    DataSetAll.Tables["PUP7902"].Rows.Add(newLigne);
                }
                else
                {
                    if (Information.IsDBNull(the_ligne[4]) == false)
                    {
                        the_ligne[4] = the_ligne[4].ToString().Replace(".", "");
                        the_ligne[5] = the_ligne[5].ToString().Replace(".", "");
                        the_ligne[8] = the_ligne[8].ToString().Replace(".", "");

                        if (Strings.Right(the_ligne[5].ToString(), 1) == "-")
                            // C'est un nombre negatif
                            the_ligne[5] = "-" + the_ligne[5].ToString().Replace("-", "").Trim();

                        if (Strings.Right(the_ligne[8].ToString(), 1) == "-")
                            // C'est un nombre negatif
                            the_ligne[8] = "-" + the_ligne[8].ToString().Replace("-", "").Trim();
                    }

                    newLigne.ItemArray = (object[])the_ligne;
                    DataSetAll.Tables["ARP0292"].Rows.Add(newLigne);
                }
            }

            SqlCommandBuilder dtr = new SqlCommandBuilder(da);
            if (MODELE == "PUP7902")
                da.Update(DataSetAll, "PUP7902");
            else
                da.Update(DataSetAll, "ARP0292");
            ocmd.Dispose();
            oda.Dispose();

            cmd.Dispose();
            da.Dispose();

            int sourceRowCount = table_access.Rows.Count;

            table_access.Dispose();

            DataSetAll.Clear();
            DataSetAll.Dispose();

            LogStep(nameof(InsertSQL), "END",
                "Model=" + MODELE + ", source rows=" + sourceRowCount);
            return result;
        }
        private void export_XML(string type_exportation, DataView dv0, string nom_fichier_dest, string nom_fic_ori)
        {
            LogStep(nameof(export_XML), "START", nom_fic_ori);

            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                XmlNode noeud_item, noeud, noeud_parent;
                string date_mois;
                int i;


                if (System.IO.File.Exists(NomFicModele_XML))
                {
                    // date_mois = InputBox("Periode de référence (ex: 2011-02)", "Merci de saisir la période de référence", "")

                    date_mois = Strings.Left(nom_fic_ori, 7);
                    if (date_mois != "")
                    {
                        try
                        {
                            xmlDoc.Load(NomFicModele_XML);
                        }
                        catch (Exception ex)
                        {
                            throw new InvalidOperationException(
                                "Le fichier modèle XML [" + NomFicModele_XML + "] n'est pas valide.", ex);
                        }

                        // On modifie la partie Entête

                        if (xmlDoc.DocumentElement.SelectSingleNode("//INSTAT/Envelope/DateTime/date") != null)
                            xmlDoc.DocumentElement.SelectSingleNode("//INSTAT/Envelope/DateTime/date").InnerText = DateTime.Now.Year + "-" + Strings.Right("00" + DateTime.Now.Month, 2) + "-" + Strings.Right("00" + DateTime.Now.Day, 2);


                        if (xmlDoc.DocumentElement.SelectSingleNode("//INSTAT/Envelope/DateTime/time") != null)
                            xmlDoc.DocumentElement.SelectSingleNode("//INSTAT/Envelope/DateTime/time").InnerText = Strings.Right("00" + DateTime.Now.Hour, 2) + ":" + Strings.Right("00" + DateTime.Now.Minute, 2) + ":" + Strings.Right("00" + DateTime.Now.Second, 2);


                        if (xmlDoc.DocumentElement.SelectSingleNode("//INSTAT/Envelope/Declaration/flowCode") != null)
                            xmlDoc.DocumentElement.SelectSingleNode("//INSTAT/Envelope/Declaration/flowCode").InnerText = type_exportation;


                        if (xmlDoc.DocumentElement.SelectSingleNode("//INSTAT/Envelope/Declaration/referencePeriod") != null)
                            xmlDoc.DocumentElement.SelectSingleNode("//INSTAT/Envelope/Declaration/referencePeriod").InnerText = date_mois;


                        noeud = xmlDoc.DocumentElement.SelectSingleNode("//INSTAT/Envelope/Declaration/Item");
                        noeud_item = xmlDoc.DocumentElement.SelectSingleNode("//INSTAT/Envelope/Declaration/Item").Clone();

                        // On supprime le noeud du document XML
                        xmlDoc.DocumentElement.SelectSingleNode("//INSTAT/Envelope/Declaration").RemoveChild(noeud);

                        noeud_parent = xmlDoc.DocumentElement.SelectSingleNode("//INSTAT/Envelope/Declaration");

                        // Partie Détail
                        i = 1;
                        foreach (DataRowView dr in dv0)
                        {
                            if (noeud_item.SelectSingleNode("itemNumber") != null)
                                noeud_item.SelectSingleNode("itemNumber").InnerText = i.ToString();

                            if (noeud_item.SelectSingleNode("CN8/CN8Code") != null)
                                noeud_item.SelectSingleNode("CN8/CN8Code").InnerText = dr["NOMENCLATURE"].ToString().Trim();

                            if (noeud_item.SelectSingleNode("MSConsDestCode") != null)
                                noeud_item.SelectSingleNode("MSConsDestCode").InnerText = dr["PROVENANCE"].ToString().Trim();

                            // Select Case type_exportation
                            // Case "A"
                            if (noeud_item.SelectSingleNode("countryOfOriginCode") != null)
                                noeud_item.SelectSingleNode("countryOfOriginCode").InnerText = dr["ORIGINE"].ToString().Trim();
                            // End Select

                            if (noeud_item.SelectSingleNode("netMass") != null)
                                noeud_item.SelectSingleNode("netMass").InnerText = dr["MASSE_NET"].ToString().Trim();

                            if (noeud_item.SelectSingleNode("quantityInSU") != null)
                                noeud_item.SelectSingleNode("quantityInSU").InnerText = dr["QTY"].ToString().Trim();

                            if (noeud_item.SelectSingleNode("invoicedAmount") != null)
                                noeud_item.SelectSingleNode("invoicedAmount").InnerText = dr["VALEUR_FISCAL"].ToString().Trim();

                            switch (type_exportation)
                            {
                                case "D":
                                    {
                                        if (noeud_item.SelectSingleNode("partnerId") != null)
                                        {
                                            if (dr["PROVENANCE"].ToString().Trim() + dr["VAT_NBR"].ToString().Trim() != "")
                                                noeud_item.SelectSingleNode("partnerId").InnerText = (dr["PROVENANCE"].ToString().Trim() + dr["VAT_NBR"].ToString().Trim()).Trim();
                                        }

                                        break;
                                    }

                                case "A":
                                    {
                                        if (noeud_item.SelectSingleNode("statisticalProcedureCode") != null)
                                            noeud_item.SelectSingleNode("statisticalProcedureCode").InnerText = "11";
                                        break;
                                    }
                            }

                            XmlNode noeud_item2;
                            noeud_item2 = xmlDoc.CreateElement("Item");
                            noeud_item2.InnerXml = noeud_item.InnerXml;

                            // On ajoute le noeud au document XML
                            noeud_parent.AppendChild(noeud_item2);

                            i = i + 1;
                        }

                        if (System.IO.File.Exists(rep_fichier_traites + nom_fichier_dest.ToString().ToLower().Replace(".xlsx", ".xml")))
                            System.IO.File.Delete(rep_fichier_traites + nom_fichier_dest.ToString().ToLower().Replace(".xlsx", ".xml"));

                        // On sauvegarde le XML
                        xmlDoc.Save(rep_fichier_traites + nom_fichier_dest.ToString().ToLower().Replace(".xlsx", ".xml"));
                    }
                    else
                    {
                        throw new InvalidOperationException(
                            "Impossible de déterminer la période de référence depuis le fichier : " +
                            nom_fic_ori);
                    }
                }
                else
                {
                    throw new FileNotFoundException(
                        "Le fichier modèle XML est introuvable.",
                        NomFicModele_XML);
                }

                LogStep(nameof(export_XML), "END", "Completed successfully");
            }
            catch (Exception ex)
            {
                LogTechnicalError(nameof(export_XML), ex, "", nom_fic_ori);
                LogStep(nameof(export_XML), "END", "Failed");
                throw;
            }
        }
        public void Export_Excel(string type_fichier, string nom_fic, string nom_fic_ori)
        {
            LogStep(nameof(Export_Excel), "START", nom_fic_ori);

            try
            {
                DataView dv0 = new DataView();
                int i;
                bool bol_ko;

                Application.DoEvents();


                // Génération du fichier 
                // D : Expedition
                // A : Introduction

                switch (type_fichier)
                {
                    case "A":
                        {
                            dv0 = DS.DefaultViewManager.CreateDataView(DS.Tables["IFA"]);
                            break;
                        }

                    case "D":
                        {
                            dv0 = DS.DefaultViewManager.CreateDataView(DS.Tables["IFL"]);
                            break;
                        }
                }

                XLWorkbook MyWorkBook = new XLWorkbook();
                IXLWorksheet MyWorkSheet = MyWorkBook.AddWorksheet("def");

                i = 1;
                bol_ko = false;
                double doubleNumber;

                foreach (DataRowView dr in dv0)
                {
                    bol_ko = false;

                    if (double.TryParse(dr["MASSE_NET"].ToString().Trim(), out doubleNumber))
                    {
                        if (System.Convert.ToDouble(dr["MASSE_NET"].ToString().Trim()) <= 0)
                            bol_ko = true;
                    }

                    if (bol_ko == false)
                    {
                        switch (type_fichier)
                        {
                            case "A":
                                {
                                    MyWorkSheet.Cell(i, 1).Value = i.ToString();
                                    MyWorkSheet.Cell(i, 2).Value = dr["NOMENCLATURE"];
                                    MyWorkSheet.Cell(i, 3).Value = dr["PROVENANCE"];
                                    MyWorkSheet.Cell(i, 4).Value = dr["VALEUR_FISCAL"];
                                    MyWorkSheet.Cell(i, 5).Value = "11";
                                    MyWorkSheet.Cell(i, 6).Value = dr["VALEUR_STAT"];
                                    MyWorkSheet.Cell(i, 7).Value = dr["MASSE_NET"];
                                    MyWorkSheet.Cell(i, 8).Value = dr["QTY"];
                                    MyWorkSheet.Cell(i, 9).Value = "1";
                                    MyWorkSheet.Cell(i, 10).Value = "1";
                                    MyWorkSheet.Cell(i, 11).Value = "EXW";
                                    MyWorkSheet.Cell(i, 12).Value = "2";
                                    MyWorkSheet.Cell(i, 13).Value = "3";
                                    MyWorkSheet.Cell(i, 14).Value = "59";
                                    MyWorkSheet.Cell(i, 15).Value = dr["ORIGINE"];
                                    break;
                                }

                            case "D":
                                {
                                    MyWorkSheet.Cell(i, 1).Value = i.ToString();
                                    MyWorkSheet.Cell(i, 2).Value = dr["NOMENCLATURE"];
                                    MyWorkSheet.Cell(i, 3).Value = dr["PROVENANCE"];
                                    MyWorkSheet.Cell(i, 4).Value = dr["VALEUR_FISCAL"];
                                    MyWorkSheet.Cell(i, 5).Value = "21";
                                    MyWorkSheet.Cell(i, 6).Value = dr["VALEUR_STAT"];
                                    MyWorkSheet.Cell(i, 7).Value = dr["MASSE_NET"];
                                    MyWorkSheet.Cell(i, 8).Value = dr["QTY"];
                                    MyWorkSheet.Cell(i, 9).Value = "1";
                                    MyWorkSheet.Cell(i, 10).Value = "1";
                                    MyWorkSheet.Cell(i, 11).Value = "EXW";
                                    MyWorkSheet.Cell(i, 12).Value = "1";
                                    MyWorkSheet.Cell(i, 13).Value = "3";
                                    MyWorkSheet.Cell(i, 14).Value = "59";
                                    MyWorkSheet.Cell(i, 15).Value = "";
                                    MyWorkSheet.Cell(i, 16).Value = dr["PROVENANCE"].ToString() + dr["VAT_NBR"].ToString();
                                    //MyWorkSheet.Cell(i, 16).Value = dr["VAT_NBR"];
                                    MyWorkSheet.Cell(i, 17).Value = dr["ORIGINE"];
                                    break;
                                }
                        }

                        i = i + 1;
                    }
                    else
                        // On exclut la ligne du Dataview
                        dv0.Delete(i - 1);
                }

                MyWorkSheet.Columns().AdjustToContents();

                // MyWorkBook.SaveAs(SaveFileDialog1.FileName)

                if (System.IO.File.Exists(rep_fichier_traites + nom_fic.ToString()))
                    System.IO.File.Delete(rep_fichier_traites + nom_fic.ToString());

                MyWorkBook.SaveAs(rep_fichier_traites + nom_fic.ToString());

                MyWorkBook.Dispose();

                // Génération du fichier XML : Prod Doaune
                // D : Expedition
                // A : Introduction

                export_XML(type_fichier, dv0, System.IO.Path.GetFileName(rep_fichier_traites + nom_fic.ToString()), System.IO.Path.GetFileName(nom_fic_ori));

                LogStep(nameof(Export_Excel), "END", "Completed successfully");
            }
            catch (Exception ex)
            {
                LogTechnicalError(nameof(Export_Excel), ex, "", nom_fic_ori);
                LogStep(nameof(Export_Excel), "END", "Failed");
                throw;
            }
        }
        public void importation_expedition(string FileName)
        {
            LogStep(nameof(importation_expedition), "START", FileName);

            try
            {
                int Num_Dde;
                SqlCommand CommandeSQL = new SqlCommand();
                OleDbCommand CommandeMDB = new OleDbCommand();
                SqlDataAdapter DA_SQL;
                string strSql;
                string strSql2;
                string msg;
                string date_rapSQL;

                do
                    System.Threading.Thread.Sleep(500);
                while (System.IO.File.Exists(FileName) && isFileOpen(FileName));


                DS.Clear();

                // C'est bien un fichier PUP7902.TXT qui a été choisi
                // On modifie le libellé pour indiquer l'etat d'avancement
                // Copie le fichier dans le repertoire 
                // IO.File.Copy(OpenFileDialog1.FileName, strRep_Travail & "arp0292.txt", True)

                LogStep(nameof(importation_expedition), "COPY SOURCE FILE", FileName);

                System.IO.File.Copy(FileName, strRep_Travail + "arp0292.txt", true);

                LogStep(nameof(importation_expedition), "DELETE OLD ACCESS FILE", Nom_modele_MDB + "arp0292.mdb");
                // Supprime la base MDB 
                try
                {
                    System.IO.File.Delete(Nom_modele_MDB + "arp0292.mdb");
                }
                catch (Exception ex)
                {
                    LogStep(nameof(importation_expedition), "DELETE OLD ACCESS FILE WARNING", ex.Message);
                }

                LogStep(nameof(importation_expedition), "MONARCH START");
                class_dev_tools.Dde_Monarch obj_monarch = new class_dev_tools.Dde_Monarch(strMonarch);
                obj_monarch.Delai_Attente_En_Minute = 2;

                // Execute Monarch sur le PC dédié
                Num_Dde = obj_monarch.Traitement_Monarch_sans_boucle(strRep_Travail + "arp0292.txt", strRep_Travail + "arp0292", rep_modele_monarch + "mod_arp0292.mod", "mdb", 1, DateTime.Now);

                LogStep(nameof(importation_expedition), "MONARCH REQUEST CREATED", "Request=" + Num_Dde);


                // Attente le retour de Monarch
                WaitForMonarchCompletion(
                    obj_monarch,
                    Num_Dde,
                    nameof(importation_expedition),
                    FileName);

                LogStep(nameof(importation_expedition), "MONARCH END", "Request=" + Num_Dde);

                try
                {
                    // Supprime le fichier du repertoire de travail
                    System.IO.File.Delete(strRep_Travail + "arp0292.txt");
                }
                catch (Exception ex)
                {
                    LogStep(nameof(importation_expedition), "DELETE TEMP FILE WARNING", ex.Message);
                }

                System.Threading.Thread.Sleep(2000);
                // Rend visible le Tab View


                LogStep(nameof(importation_expedition), "ACCESS START", Nom_modele_MDB + "arp0292.mdb");
                OleDbConnection ConBaseMDB = new OleDbConnection();
                ConBase = new SqlConnection();

                ConBaseMDB.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Nom_modele_MDB + "arp0292.mdb";

                try
                {
                    ConBaseMDB.Open();
                }
                catch (Exception ex)
                {
                    LogTechnicalError(
                        nameof(importation_expedition),
                        ex,
                        "Open Microsoft Access database",
                        Nom_modele_MDB + "arp0292.mdb");
                    throw;
                }


                CommandeMDB.CommandText = "select [DATE_RAP] from filtre_code_intra group by [DATE_RAP]";
                CommandeMDB.Connection = ConBaseMDB;

                date_rap = (string)CommandeMDB.ExecuteScalar();
                Application.DoEvents();

                // Ouverture de la base de Données Intrastat
                if (ConBase.State == ConnectionState.Closed)
                {
                    ConBase.ConnectionString = strCon;
                    ConBase.Open();
                }

                // Recherche si le rapport existe déjà
                CommandeSQL.CommandText = "SELECT DATE_RAP FROM ARP0292 WHERE DATE_RAP ='" + date_rap + "'";
                CommandeSQL.Connection = ConBase;
                date_rapSQL = (string)CommandeSQL.ExecuteScalar();

                if (date_rapSQL == date_rap)
                {
                    CommandeSQL.CommandText = "delete from ARP0292 WHERE DATE_RAP ='" + date_rap + "'";
                    CommandeSQL.Connection = ConBase;
                    CommandeSQL.ExecuteNonQuery();
                    Application.DoEvents();
                }

                // insertion dans sql
                strSql = "select *  from filtre_code_intra";
                strSql2 = "select *  from arp0292";
                msg = InsertSQL(strSql, strSql2, ConBaseMDB, "ARP0292");
                if (msg != "")
                {
                    return;
                }
                Application.DoEvents();

                // On ferme la connexion a la base MDB
                ConBaseMDB.Dispose();
                ConBaseMDB.Close();


                CommandeSQL.CommandText = "update ARP0292 set MF='QU' where rtrim(MF)='' and DATE_RAP='" + date_rap + "'";
                CommandeSQL.CommandTimeout = 300;
                CommandeSQL.ExecuteNonQuery();

                CommandeSQL.CommandText = "update ARP0292 set MF='QU' where isnumeric(MF)=1 and DATE_RAP='" + date_rap + "'";
                CommandeSQL.CommandTimeout = 300;
                CommandeSQL.ExecuteNonQuery();

                CommandeSQL.CommandText = "SELECT CODE_INTRA AS NOMENCLATURE, SF as PROVENANCE, round(SUM(LOCAL_AMT),0) AS VALEUR_FISCAL,round(SUM(LOCAL_AMT*1.001),0) AS VALEUR_STAT,CASE WHEN round(SUM(POIDS),0) = 0 AND round(SUM(QTY_RECVD),0) < 50 THEN 1 ELSE round(SUM(POIDS),0) END AS MASSE_NET, round(SUM(QTY_RECVD),0) AS QTY,VAT_NBR,left(LTRIM(RTRIM(MF)),3) as ORIGINE" + " FROM ARP0292  GROUP BY NUM_RAP, DATE_RAP, CODE_INTRA, SF,MF,VAT_NBR HAVING DATE_RAP ='" + date_rap + "' ORDER BY CODE_INTRA";
                // " FROM ARP0292 WHERE (POIDS >= 0) AND (LOCAL_AMT > 0) GROUP BY NUM_RAP, DATE_RAP, CODE_INTRA, SF,MF,VAT_NBR HAVING DATE_RAP ='" & date_rap & "' ORDER BY CODE_INTRA"

                CommandeSQL.Connection = ConBase;
                DA_SQL = new SqlDataAdapter(CommandeSQL);
                DA_SQL.Fill(DS, "IFL");

                Application.DoEvents();
                Application.DoEvents();

                ConBase.Close();

                // Exportation Excel
                Export_Excel("D", "idep_expedition_" + DateTime.Now.Year + "_" + Strings.Right("00" + DateTime.Now.Month, 2) + "_" + Strings.Right("00" + DateTime.Now.Day, 2) + ".xlsx", FileName);

                LogStep(nameof(importation_expedition), "END", "Completed successfully");
            }
            catch (Exception ex)
            {
                LogTechnicalError(nameof(importation_expedition), ex, "", FileName);
                LogStep(nameof(importation_expedition), "END", "Failed");
                throw;
            }
        }

        public void importation_expedition_with_odbc(string FileName)
        {
            LogStep(nameof(importation_expedition_with_odbc), "START", FileName);

            try
            {
                int Num_Dde;
                SqlCommand CommandeSQL = new SqlCommand();
                OdbcCommand CommandeMDB = new OdbcCommand();
                SqlDataAdapter DA_SQL;
                string strSql;
                string strSql2;
                string msg;
                string date_rapSQL;

                do
                    System.Threading.Thread.Sleep(500);
                while (System.IO.File.Exists(FileName) && isFileOpen(FileName));


                DS.Clear();

                // C'est bien un fichier PUP7902.TXT qui a été choisi
                // On modifie le libellé pour indiquer l'etat d'avancement
                // Copie le fichier dans le repertoire 
                // IO.File.Copy(OpenFileDialog1.FileName, strRep_Travail & "arp0292.txt", True)


                System.IO.File.Copy(FileName, strRep_Travail + "arp0292.txt", true);

                // Supprime la base MDB 
                try
                {
                    System.IO.File.Delete(Nom_modele_MDB + "arp0292.mdb");
                }
                catch (Exception ex)
                {
                    LogStep(
                        nameof(importation_expedition_with_odbc),
                        "DELETE OLD ACCESS FILE WARNING",
                        ex.Message);
                }

                class_dev_tools.Dde_Monarch obj_monarch = new class_dev_tools.Dde_Monarch(strMonarch);
                obj_monarch.Delai_Attente_En_Minute = 2;

                // Execute Monarch sur le PC dédié
                Num_Dde = obj_monarch.Traitement_Monarch_sans_boucle(strRep_Travail + "arp0292.txt", strRep_Travail + "arp0292", rep_modele_monarch + "mod_arp0292.mod", "mdb", 1, DateTime.Now);



                // Attente le retour de Monarch
                WaitForMonarchCompletion(
                    obj_monarch,
                    Num_Dde,
                    nameof(importation_expedition_with_odbc),
                    FileName);


                try
                {
                    // Supprime le fichier du repertoire de travail
                    System.IO.File.Delete(strRep_Travail + "arp0292.txt");
                }
                catch (Exception ex)
                {
                    LogStep(nameof(importation_expedition_with_odbc), "DELETE TEMP FILE WARNING", ex.Message);
                }

                System.Threading.Thread.Sleep(2000);
                // Rend visible le Tab View


                string connectionString = "Driver={Microsoft Access Driver (*.mdb, *.accdb)}; Dbq=" + Nom_modele_MDB + "arp0292.mdb";
                //string connectionString = "Driver={MS Access Driver (*.mdb, *.accdb)}; Dbq=" + Nom_modele_MDB + "arp0292.mdb";

                OdbcConnection ConBaseMDB = new OdbcConnection(connectionString);
                try
                {
                    ConBaseMDB.Open();
                }
                catch (Exception ex)
                {
                    LogTechnicalError(
                        nameof(importation_expedition_with_odbc),
                        ex,
                        "Open ODBC Microsoft Access database",
                        Nom_modele_MDB + "arp0292.mdb");
                    throw;
                }

                CommandeMDB.CommandText = "select [DATE_RAP] from filtre_code_intra group by [DATE_RAP]";
                CommandeMDB.Connection = ConBaseMDB;

                date_rap = (string)CommandeMDB.ExecuteScalar();
                Application.DoEvents();

                ConBase = new SqlConnection();

                // Ouverture de la base de Données Intrastat
                if (ConBase.State == ConnectionState.Closed)
                {
                    ConBase.ConnectionString = strCon;
                    ConBase.Open();
                }

                // Recherche si le rapport existe déjà
                CommandeSQL.CommandText = "SELECT DATE_RAP FROM ARP0292 WHERE DATE_RAP ='" + date_rap + "'";
                CommandeSQL.Connection = ConBase;
                date_rapSQL = (string)CommandeSQL.ExecuteScalar();

                if (date_rapSQL == date_rap)
                {
                    CommandeSQL.CommandText = "delete from ARP0292 WHERE DATE_RAP ='" + date_rap + "'";
                    CommandeSQL.Connection = ConBase;
                    CommandeSQL.ExecuteNonQuery();
                    Application.DoEvents();
                }

                // insertion dans sql
                strSql = "select *  from filtre_code_intra";
                strSql2 = "select *  from arp0292";
                msg = InsertSQL_with_odbc(strSql, strSql2, ConBaseMDB, "ARP0292");
                if (msg != "")
                {
                    return;
                }
                Application.DoEvents();

                // On ferme la connexion a la base MDB
                ConBaseMDB.Dispose();
                ConBaseMDB.Close();


                CommandeSQL.CommandText = "update ARP0292 set MF='QU' where rtrim(MF)='' and DATE_RAP='" + date_rap + "'";
                CommandeSQL.CommandTimeout = 300;
                CommandeSQL.ExecuteNonQuery();

                CommandeSQL.CommandText = "update ARP0292 set MF='QU' where isnumeric(MF)=1 and DATE_RAP='" + date_rap + "'";
                CommandeSQL.CommandTimeout = 300;
                CommandeSQL.ExecuteNonQuery();

                CommandeSQL.CommandText = "SELECT CODE_INTRA AS NOMENCLATURE, SF as PROVENANCE, round(SUM(LOCAL_AMT),0) AS VALEUR_FISCAL,round(SUM(LOCAL_AMT*1.001),0) AS VALEUR_STAT,CASE WHEN round(SUM(POIDS),0) = 0 AND round(SUM(QTY_RECVD),0) < 50 THEN 1 ELSE round(SUM(POIDS),0) END AS MASSE_NET, round(SUM(QTY_RECVD),0) AS QTY,VAT_NBR,left(LTRIM(RTRIM(MF)),3) as ORIGINE" + " FROM ARP0292  GROUP BY NUM_RAP, DATE_RAP, CODE_INTRA, SF,MF,VAT_NBR HAVING DATE_RAP ='" + date_rap + "' ORDER BY CODE_INTRA";
                // " FROM ARP0292 WHERE (POIDS >= 0) AND (LOCAL_AMT > 0) GROUP BY NUM_RAP, DATE_RAP, CODE_INTRA, SF,MF,VAT_NBR HAVING DATE_RAP ='" & date_rap & "' ORDER BY CODE_INTRA"

                CommandeSQL.Connection = ConBase;
                DA_SQL = new SqlDataAdapter(CommandeSQL);
                DA_SQL.Fill(DS, "IFL");

                Application.DoEvents();
                Application.DoEvents();

                ConBase.Close();

                // Exportation Excel
                Export_Excel("D", "idep_expedition_" + DateTime.Now.Year + "_" + Strings.Right("00" + DateTime.Now.Month, 2) + "_" + Strings.Right("00" + DateTime.Now.Day, 2) + ".xlsx", FileName);

                LogStep(nameof(importation_expedition_with_odbc), "END", "Completed successfully");
            }
            catch (Exception ex)
            {
                LogTechnicalError(nameof(importation_expedition_with_odbc), ex, "", FileName);
                LogStep(nameof(importation_expedition_with_odbc), "END", "Failed");
                throw;
            }
        }
        public void importation_acquisition(string FileName)
        {
            LogStep(nameof(importation_acquisition), "START", FileName);

            try
            {
                int Num_Dde;
                SqlCommand CommandeSQL = new SqlCommand();
                OleDbCommand CommandeMDB = new OleDbCommand();
                SqlDataAdapter DA_SQL;
                string strSql;
                string strSql2;
                string msg;
                string date_rapSQL;

                do
                    System.Threading.Thread.Sleep(500);
                while (System.IO.File.Exists(FileName) && isFileOpen(FileName));

                DS.Clear();


                // On modifie le libellé pour indiquer l'etat d'avancement
                // Copie le fichier dans le repertoire 
                System.IO.File.Copy(FileName, strRep_Travail + "pup7902.txt", true);

                // Supprime la base MDB 
                System.IO.File.Delete(Nom_modele_MDB + "pup7902.mdb");

                class_dev_tools.Dde_Monarch obj_monarch = new class_dev_tools.Dde_Monarch(strMonarch);
                obj_monarch.Delai_Attente_En_Minute = 2;

                // Execute Monarch sur le PC dédié
                Num_Dde = obj_monarch.Traitement_Monarch_sans_boucle(strRep_Travail + "pup7902.txt", strRep_Travail + "pup7902", rep_modele_monarch + "mod_pup7902.mod", "mdb", 1, DateTime.Now);



                // Attente le retour de Monarch
                WaitForMonarchCompletion(
                    obj_monarch,
                    Num_Dde,
                    nameof(importation_acquisition),
                    FileName);

                // Supprime le fichier du repertoire de travail
                System.IO.File.Delete(strRep_Travail + "pup7902.txt");
                System.Threading.Thread.Sleep(2000);

                OleDbConnection ConBaseMDB = new OleDbConnection();
                ConBaseMDB.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Nom_modele_MDB + "pup7902.mdb";
                ConBaseMDB.Open();
                CommandeMDB.CommandText = "select [DATE_RAP] from filtre_code_intra group by [DATE_RAP]";
                CommandeMDB.Connection = ConBaseMDB;

                date_rap = (string)CommandeMDB.ExecuteScalar();

                Application.DoEvents();

                // Ouverture de la base de Données Intrastat
                if (ConBase == null)
                {
                    ConBase = new SqlConnection();
                }

                if (ConBase.State == ConnectionState.Closed)
                {
                    ConBase.ConnectionString = strCon;
                    ConBase.Open();
                }

                // Recherche si le rapport existe déjà
                CommandeSQL.CommandText = "SELECT DATE_RAP FROM PUP7902 WHERE DATE_RAP ='" + date_rap + "'";
                CommandeSQL.Connection = ConBase;
                date_rapSQL = (string)CommandeSQL.ExecuteScalar();

                if (date_rapSQL == date_rap)
                {
                    CommandeSQL.CommandText = "delete from PUP7902 WHERE DATE_RAP ='" + date_rap + "'";
                    CommandeSQL.Connection = ConBase;
                    CommandeSQL.CommandTimeout = 300;
                    CommandeSQL.ExecuteNonQuery();
                    Application.DoEvents();
                }

                // insertion dans sql
                strSql = "select *  from filtre_code_intra";
                strSql2 = "select *  from pup7902";
                msg = InsertSQL(strSql, strSql2, ConBaseMDB, "PUP7902");
                if (msg != "")
                {
                    return;
                }
                Application.DoEvents();
                // On ferme la connexion a la base MDB
                ConBaseMDB.Dispose();
                ConBaseMDB.Close();

                CommandeSQL.Connection = ConBase;

                CommandeSQL.CommandText = "delete from PUP7902 where SF='FR' and DATE_RAP='" + date_rap + "'";
                CommandeSQL.CommandTimeout = 300;
                CommandeSQL.ExecuteNonQuery();

                CommandeSQL.CommandText = "update PUP7902 set MF='QU' where rtrim(MF)='' and DATE_RAP='" + date_rap + "'";
                CommandeSQL.CommandTimeout = 300;
                CommandeSQL.ExecuteNonQuery();

                CommandeSQL.CommandText = "update PUP7902 set MF='QU' where isnumeric(MF)=1 and DATE_RAP='" + date_rap + "'";
                CommandeSQL.CommandTimeout = 300;
                CommandeSQL.ExecuteNonQuery();

                CommandeSQL.CommandText = "SELECT CODE_INTRA AS NOMENCLATURE, SF as PROVENANCE, round(SUM(LOCAL_AMT),0) AS VALEUR_FISCAL,round(SUM(LOCAL_AMT*1.001),0) AS VALEUR_STAT,CASE WHEN round(SUM(POIDS),0) = 0 AND round(SUM(QTY_RCVD),0) < 50 THEN 1 ELSE round(SUM(POIDS),0) END AS MASSE_NET, round(SUM(QTY_RCVD),0) AS QTY,DEVISE,left(LTRIM(RTRIM(MF)),3) as ORIGINE" + " FROM PUP7902 GROUP BY NUM_RAP, DATE_RAP, CODE_INTRA, SF, DEVISE,MF,LEFT(LTRIM(RTRIM(MF)), 3) HAVING DATE_RAP ='" + date_rap + "' ORDER BY CODE_INTRA";
                DA_SQL = new SqlDataAdapter(CommandeSQL);
                DA_SQL.Fill(DS, "IFA");

                Application.DoEvents();
                ConBase.Close();

                // Exportation Excel
                Export_Excel("A", "idep_acquisition_" + DateTime.Now.Year + "_" + Strings.Right("00" + DateTime.Now.Month, 2) + "_" + Strings.Right("00" + DateTime.Now.Day, 2) + ".xlsx", FileName);

                LogStep(nameof(importation_acquisition), "END", "Completed successfully");
            }
            catch (Exception ex)
            {
                LogTechnicalError(nameof(importation_acquisition), ex, "", FileName);
                LogStep(nameof(importation_acquisition), "END", "Failed");
                throw;
            }
        }


        public void importation_acquisition_with_odbc(string FileName)
        {
            LogStep(nameof(importation_acquisition_with_odbc), "START", FileName);

            try
            {
                int Num_Dde;
                SqlCommand CommandeSQL = new SqlCommand();
                OdbcCommand CommandeMDB = new OdbcCommand();
                SqlDataAdapter DA_SQL;
                string strSql;
                string strSql2;
                string msg;
                string date_rapSQL;

                do
                    System.Threading.Thread.Sleep(500);
                while (System.IO.File.Exists(FileName) && isFileOpen(FileName));

                DS.Clear();


                // On modifie le libellé pour indiquer l'etat d'avancement
                // Copie le fichier dans le repertoire 
                System.IO.File.Copy(FileName, strRep_Travail + "pup7902.txt", true);

                // Supprime la base MDB 
                System.IO.File.Delete(Nom_modele_MDB + "pup7902.mdb");

                class_dev_tools.Dde_Monarch obj_monarch = new class_dev_tools.Dde_Monarch(strMonarch);
                obj_monarch.Delai_Attente_En_Minute = 2;

                // Execute Monarch sur le PC dédié
                Num_Dde = obj_monarch.Traitement_Monarch_sans_boucle(strRep_Travail + "pup7902.txt", strRep_Travail + "pup7902", rep_modele_monarch + "mod_pup7902.mod", "mdb", 1, DateTime.Now);



                // Attente le retour de Monarch
                WaitForMonarchCompletion(
                    obj_monarch,
                    Num_Dde,
                    nameof(importation_acquisition_with_odbc),
                    FileName);

                // Supprime le fichier du repertoire de travail
                System.IO.File.Delete(strRep_Travail + "pup7902.txt");
                System.Threading.Thread.Sleep(2000);

                //OdbcConnection ConBaseMDB = new OdbcConnection(connectionString);
                string connectionString = "Driver={Microsoft Access Driver (*.mdb, *.accdb)}; Dbq=" + Nom_modele_MDB + "pup7902.mdb";
                //string connectionString = "Driver={MS Access Driver (*.mdb, *.accdb)}; Dbq=" + Nom_modele_MDB + "arp0292.mdb";

                OdbcConnection ConBaseMDB = new OdbcConnection(connectionString);
                try
                {
                    ConBaseMDB.Open();
                }
                catch (Exception ex)
                {
                    LogTechnicalError(
                        nameof(importation_acquisition_with_odbc),
                        ex,
                        "Open ODBC Microsoft Access database",
                        Nom_modele_MDB + "pup7902.mdb");
                    throw;
                }
                CommandeMDB.CommandText = "select [DATE_RAP] from filtre_code_intra group by [DATE_RAP]";
                CommandeMDB.Connection = ConBaseMDB;

                date_rap = (string)CommandeMDB.ExecuteScalar();

                Application.DoEvents();

                ConBase = new SqlConnection();

                // Ouverture de la base de Données Intrastat
                if (ConBase.State == ConnectionState.Closed)
                {
                    ConBase.ConnectionString = strCon;
                    ConBase.Open();
                }

                // Recherche si le rapport existe déjà
                CommandeSQL.CommandText = "SELECT DATE_RAP FROM PUP7902 WHERE DATE_RAP ='" + date_rap + "'";
                CommandeSQL.Connection = ConBase;
                date_rapSQL = (string)CommandeSQL.ExecuteScalar();

                if (date_rapSQL == date_rap)
                {
                    CommandeSQL.CommandText = "delete from PUP7902 WHERE DATE_RAP ='" + date_rap + "'";
                    CommandeSQL.Connection = ConBase;
                    CommandeSQL.CommandTimeout = 300;
                    CommandeSQL.ExecuteNonQuery();
                    Application.DoEvents();
                }

                // insertion dans sql
                strSql = "select *  from filtre_code_intra";
                strSql2 = "select *  from pup7902";
                msg = InsertSQL_with_odbc(strSql, strSql2, ConBaseMDB, "PUP7902");
                if (msg != "")
                {
                    return;
                }
                Application.DoEvents();
                // On ferme la connexion a la base MDB
                ConBaseMDB.Dispose();
                ConBaseMDB.Close();

                CommandeSQL.Connection = ConBase;

                CommandeSQL.CommandText = "delete from PUP7902 where SF='FR' and DATE_RAP='" + date_rap + "'";
                CommandeSQL.CommandTimeout = 300;
                CommandeSQL.ExecuteNonQuery();

                CommandeSQL.CommandText = "update PUP7902 set MF='QU' where rtrim(MF)='' and DATE_RAP='" + date_rap + "'";
                CommandeSQL.CommandTimeout = 300;
                CommandeSQL.ExecuteNonQuery();

                CommandeSQL.CommandText = "update PUP7902 set MF='QU' where isnumeric(MF)=1 and DATE_RAP='" + date_rap + "'";
                CommandeSQL.CommandTimeout = 300;
                CommandeSQL.ExecuteNonQuery();

                CommandeSQL.CommandText = "SELECT CODE_INTRA AS NOMENCLATURE, SF as PROVENANCE, round(SUM(LOCAL_AMT),0) AS VALEUR_FISCAL,round(SUM(LOCAL_AMT*1.001),0) AS VALEUR_STAT,CASE WHEN round(SUM(POIDS),0) = 0 AND round(SUM(QTY_RCVD),0) < 50 THEN 1 ELSE round(SUM(POIDS),0) END AS MASSE_NET, round(SUM(QTY_RCVD),0) AS QTY,DEVISE,left(LTRIM(RTRIM(MF)),3) as ORIGINE" + " FROM PUP7902 GROUP BY NUM_RAP, DATE_RAP, CODE_INTRA, SF, DEVISE,MF,LEFT(LTRIM(RTRIM(MF)), 3) HAVING DATE_RAP ='" + date_rap + "' ORDER BY CODE_INTRA";
                DA_SQL = new SqlDataAdapter(CommandeSQL);
                DA_SQL.Fill(DS, "IFA");

                Application.DoEvents();
                ConBase.Close();

                // Exportation Excel
                Export_Excel("A", "idep_acquisition_" + DateTime.Now.Year + "_" + Strings.Right("00" + DateTime.Now.Month, 2) + "_" + Strings.Right("00" + DateTime.Now.Day, 2) + ".xlsx", FileName);

                LogStep(nameof(importation_acquisition_with_odbc), "END", "Completed successfully");
            }
            catch (Exception ex)
            {
                LogTechnicalError(nameof(importation_acquisition_with_odbc), ex, "", FileName);
                LogStep(nameof(importation_acquisition_with_odbc), "END", "Failed");
                throw;
            }
        }

        private void WaitForMonarchCompletion(
            class_dev_tools.Dde_Monarch monarch,
            int requestNumber,
            string callerMethod,
            string sourceFile)
        {
            DateTime startTime = DateTime.Now;
            int alertInterval = timer_attente_monarch_minutes < 1
                ? 10
                : timer_attente_monarch_minutes;
            DateTime nextAlertTime = startTime.AddMinutes(alertInterval);
            int status = 1;

            LogStep(
                callerMethod,
                "MONARCH WAIT START",
                "Request=" + requestNumber +
                ", alert interval=" + alertInterval + " minute(s)");

            while (status == 1)
            {
                string rawStatus;
                try
                {
                    rawStatus = monarch.Check_status_Dde_Monarch(requestNumber);
                }
                catch (Exception ex)
                {
                    LogTechnicalError(
                        callerMethod,
                        ex,
                        "Check_status_Dde_Monarch(" + requestNumber + ")",
                        sourceFile);
                    throw;
                }

                if (!int.TryParse(rawStatus, out status))
                {
                    throw new InvalidOperationException(
                        "Invalid Monarch status for request " + requestNumber +
                        " : " + rawStatus);
                }

                if (status != 1)
                {
                    break;
                }

                TimeSpan elapsed = DateTime.Now - startTime;

                if (DateTime.Now >= nextAlertTime)
                {
                    SendMonarchWaitingAlert(
                        callerMethod,
                        requestNumber,
                        sourceFile,
                        elapsed);

                    nextAlertTime = DateTime.Now.AddMinutes(alertInterval);
                }

                System.Threading.Thread.Sleep(30000);
                Application.DoEvents();
            }

            if (status != 2)
            {
                throw new InvalidOperationException(
                    "Monarch request " + requestNumber +
                    " ended with status " + status + ".");
            }

            LogStep(
                callerMethod,
                "MONARCH WAIT END",
                "Request=" + requestNumber +
                ", duration=" +
                (DateTime.Now - startTime).ToString(@"hh\:mm\:ss"));
        }

        private void SendMonarchWaitingAlert(
            string callerMethod,
            int requestNumber,
            string sourceFile,
            TimeSpan elapsed)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(email_in_case_of_technical_issue) ||
                    !email_in_case_of_technical_issue.Contains("@"))
                {
                    WriteToFile(
                        "       Unable to send Monarch waiting alert because recipient is empty.");
                    return;
                }

                class_dev_tools.envoi_mail alertEmail =
                    new class_dev_tools.envoi_mail();

                alertEmail.mail_to = email_in_case_of_technical_issue;
                alertEmail.mail_from = "Intrastat";
                alertEmail.nom_fichier_xml = "intrastat.xml";
                alertEmail.mail_subject =
                    global_application_name + " - Monarch request still running";
                alertEmail.priorite = "143";
                alertEmail.mail_body =
                    "The Monarch request is still running.<br/><br/>" +
                    "<b>Method:</b> " + HtmlEncode(callerMethod) + "<br/>" +
                    "<b>Request:</b> " + requestNumber + "<br/>" +
                    "<b>Elapsed time:</b> " + elapsed.ToString(@"hh\:mm\:ss") + "<br/>" +
                    "<b>Source file:</b> " + HtmlEncode(sourceFile) + "<br/><br/>" +
                    "The IDEP process will continue waiting for Monarch to finish.";

                alertEmail.envoi_mail();

                WriteToFile(
                    "       Monarch waiting alert sent. Request=" +
                    requestNumber + ", elapsed=" +
                    elapsed.ToString(@"hh\:mm\:ss"));
            }
            catch (Exception ex)
            {
                WriteToFile(
                    "       Unable to send Monarch waiting alert : " + ex);
            }
        }

        public void WriteToFile(string message)
        {
            try
            {
                string targetFolder = string.IsNullOrWhiteSpace(logs_folder)
                    ? AppDomain.CurrentDomain.BaseDirectory
                    : logs_folder;

                lock (logLock)
                {
                    Directory.CreateDirectory(targetFolder);

                    string filePath = Path.Combine(
                        targetFolder,
                        "IMCA_" + (global_session_name ?? "").ToUpperInvariant() + "_" +
                        DateTime.Now.ToString("dd_MM_yyyy") + "_FR_" +
                        global_application_name + ".txt");

                    File.AppendAllText(
                        filePath,
                        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") +
                        " - " + message + Environment.NewLine);
                }
            }
            catch
            {
                // Logging failures must never hide the original processing error.
            }
        }

        private void LogStep(string methodName, string step, string details = "")
        {
            WriteToFile(
                "   [" + methodName + "] " + step +
                (string.IsNullOrWhiteSpace(details) ? "" : " - " + details));
        }

        private void LogTechnicalError(
            string methodName,
            Exception exception,
            string commandText = "",
            string fileName = "")
        {
            if (exception == null)
            {
                return;
            }

            const string loggedKey = "IDEP_FR_TECHNICAL_ERROR_LOGGED";
            if (exception.Data.Contains(loggedKey))
            {
                return;
            }

            exception.Data[loggedKey] = true;
            const string indent = "       ";

            WriteToFile(
                indent + "TECHNICAL ERROR" + Environment.NewLine +
                indent + "   Method  : " + methodName + Environment.NewLine +
                indent + "   Type    : " + exception.GetType().FullName + Environment.NewLine +
                indent + "   Message : " + exception.Message + Environment.NewLine +
                indent + "   File    : " +
                    (string.IsNullOrWhiteSpace(fileName) ? "<not available>" : fileName) +
                    Environment.NewLine +
                indent + "   Command : " +
                    (string.IsNullOrWhiteSpace(commandText) ? "<not available>" : commandText) +
                    Environment.NewLine +
                indent + "   Details : " + exception);

            SendTechnicalIssueMail(
                methodName,
                exception,
                commandText,
                fileName);
        }

        private void SendTechnicalIssueMail(
            string methodName,
            Exception exception,
            string commandText = "",
            string fileName = "")
        {
            if (string.IsNullOrWhiteSpace(email_in_case_of_technical_issue) ||
                !email_in_case_of_technical_issue.Contains("@"))
            {
                WriteToFile(
                    "       Unable to send technical issue email because recipient is empty.");
                return;
            }

            // Prevent recursive alerts if the email component itself fails.
            if (System.Threading.Interlocked.CompareExchange(
                    ref technical_email_is_sending,
                    1,
                    0) != 0)
            {
                return;
            }

            try
            {
                class_dev_tools.envoi_mail technicalEmail =
                    new class_dev_tools.envoi_mail();

                technicalEmail.mail_to = email_in_case_of_technical_issue;
                technicalEmail.mail_from = "Intrastat";
                technicalEmail.nom_fichier_xml = "intrastat.xml";
                technicalEmail.mail_subject =
                    global_application_name + " - Technical error in " + methodName;
                technicalEmail.priorite = "143";
                technicalEmail.mail_body =
                    "A technical error occurred in " + global_application_name + ".<br/><br/>" +
                    "<b>Method:</b> " + HtmlEncode(methodName) + "<br/>" +
                    "<b>Type:</b> " + HtmlEncode(exception.GetType().FullName) + "<br/>" +
                    "<b>Message:</b> " + HtmlEncode(exception.Message) + "<br/>" +
                    "<b>File:</b> " + HtmlEncode(
                        string.IsNullOrWhiteSpace(fileName)
                            ? "<not available>"
                            : fileName) + "<br/>" +
                    "<b>Command:</b><br/><pre>" + HtmlEncode(
                        string.IsNullOrWhiteSpace(commandText)
                            ? "<not available>"
                            : commandText) + "</pre>";

                technicalEmail.envoi_mail();

                WriteToFile(
                    "       Technical issue email sent to " +
                    email_in_case_of_technical_issue + ".");
            }
            catch (Exception mailException)
            {
                // Do not call LogTechnicalError here because that would attempt
                // to send another email for the email failure.
                WriteToFile(
                    "       Unable to send technical issue email : " +
                    mailException);
            }
            finally
            {
                System.Threading.Interlocked.Exchange(
                    ref technical_email_is_sending,
                    0);
            }
        }

        private static string HtmlEncode(string value)
        {
            return System.Net.WebUtility.HtmlEncode(value ?? "");
        }

        public void Files_Management(string sql_con, string logs, string temp_folder, string session_name)
        {

            string[] tab_fich;
            System.IO.FileInfo fich;
            FileSecurity titi;
            string proprietaire;
            string global_parameters = "";
            string service_path = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);


            setlogs_folder(service_path + "\\" + logs);

            global_session_name = session_name;

            bool processingSucceeded = false;

            try
            {
                LogStep(
                    nameof(Files_Management),
                    "START",
                    "Session=" + global_session_name);

                // Get the global parameters of the ACTION
                global_parameters = get_IMCA_paramters(sql_con, "IDEP_FR");

                var idep_param = new IDEP_FRJSON_file();
                idep_param = JsonConvert.DeserializeObject<IDEP_FRJSON_file>(global_parameters);

                // str_annuaire_con = idep_param.list_param.FirstOrDefault(x => x.param == "sql_annuaire").valeur.ToString();
                //strCon = idep_param.list_param.FirstOrDefault(x => x.param == "sql_con").valeur.ToString();

                sql_con_parameter_global = idep_param.list_param.FirstOrDefault(x => x.param == "sql_con_parameter_global").valeur.ToString();

                sql_annuaire_parameter_global = idep_param.list_param.FirstOrDefault(x => x.param == "sql_annuaire_parameter_global").valeur.ToString();

                sql_monarch_parameter_global = idep_param.list_param.FirstOrDefault(x => x.param == "sql_monarch_parameter_global").valeur.ToString();

                parameters monarchTimerParameter = idep_param.list_param.FirstOrDefault(
                    x => string.Equals(
                        x.param,
                        "timer_attente_monarch",
                        StringComparison.OrdinalIgnoreCase));

                if (monarchTimerParameter != null &&
                    !int.TryParse(monarchTimerParameter.valeur, out timer_attente_monarch_minutes))
                {
                    timer_attente_monarch_minutes = 10;
                }

                if (timer_attente_monarch_minutes < 1)
                {
                    timer_attente_monarch_minutes = 10;
                }

                parameters technicalEmailParameter = idep_param.list_param.FirstOrDefault(
                    x => string.Equals(
                        x.param,
                        "email_in_case_of_technical_issue_parameter_global",
                        StringComparison.OrdinalIgnoreCase));

                email_in_case_of_technical_issue_parameter_global =
                    technicalEmailParameter == null
                        ? ""
                        : technicalEmailParameter.valeur ?? "";

                strCon = get_IMCA_paramters(sql_con, sql_con_parameter_global);

                str_annuaire_con = get_IMCA_paramters(sql_con, sql_annuaire_parameter_global);

                strMonarch = get_IMCA_paramters(sql_con, sql_monarch_parameter_global);

                if (!string.IsNullOrWhiteSpace(
                    email_in_case_of_technical_issue_parameter_global))
                {
                    email_in_case_of_technical_issue = get_IMCA_paramters(
                        sql_con,
                        email_in_case_of_technical_issue_parameter_global);
                }

                LogStep(
                    nameof(Files_Management),
                    "TECHNICAL EMAIL CONFIGURATION",
                    string.IsNullOrWhiteSpace(email_in_case_of_technical_issue)
                        ? "No technical recipient configured"
                        : "Technical recipient configured");

                LogStep(
                    nameof(Files_Management),
                    "MONARCH TIMER CONFIGURATION",
                    "Alert interval=" + timer_attente_monarch_minutes + " minute(s)");

                string rep_a_scanner = idep_param.list_param.FirstOrDefault(x => x.param == "rep_a_scanner").valeur.ToString();
                rep_fichier_traites = idep_param.list_param.FirstOrDefault(x => x.param == "rep_fichier_traites").valeur.ToString();
                strRep_Travail = idep_param.list_param.FirstOrDefault(x => x.param == "repertoire_temporaire").valeur.ToString();
                Nom_modele_MDB = strRep_Travail;
                rep_modele_monarch = idep_param.list_param.FirstOrDefault(x => x.param == "modeles_monarch").valeur.ToString();

                NomFicModele_XML = idep_param.list_param.FirstOrDefault(x => x.param == "modele_xml").valeur.ToString();
                class_dev_tools.envoi_mail obj_email = new class_dev_tools.envoi_mail();
                class_dev_tools.DroitAnnuaire user_info = new class_dev_tools.DroitAnnuaire(str_annuaire_con);

                DS = new DataSet();

                LogStep(nameof(Files_Management), "FILE SCAN START");

                tab_fich = System.IO.Directory.GetFiles(rep_a_scanner);

                LogStep(nameof(Files_Management), "FILES FOUND",
                    "Folder=" + rep_a_scanner + ", count=" + tab_fich.Length);

                for (int i = 0; i <= tab_fich.Length - 1; i++)
                {
                    fich = new System.IO.FileInfo(tab_fich[i]);
                    if ((fich.Name.ToUpper().Contains("PUP7902.TXT") || fich.Name.ToUpper().Contains("ARP0292.TXT")) && fich.Extension.ToLower() == ".txt")
                    {

                        LogStep(nameof(Files_Management), "FILE FOUND", fich.Name.ToUpperInvariant());

                        // Le fichier ne comporte pas la bonne extension.
                        // On envoie un email au proprietaire 

                        titi = fich.GetAccessControl();
                        proprietaire = titi.GetOwner(typeof(NTAccount)).Value.ToString();
                        proprietaire = proprietaire.ToLower().Replace(@"corporate\", "");

                        // On envoi un mail au proprietaire 
                        user_info.Get_User_Info(proprietaire);


                        if (fich.Name.ToUpper().Contains("PUP7902.TXT"))
                        {

                            LogStep(nameof(Files_Management), "ACQUISITION START", fich.FullName);
                            importation_acquisition_with_odbc(fich.FullName);
                            LogStep(nameof(Files_Management), "ACQUISITION END", fich.FullName);

                            obj_email.mail_to = user_info.Adresse_Email;
                            obj_email.mail_from = "Intrastat";
                            obj_email.nom_fichier_xml = "intrastat.xml";
                            obj_email.mail_subject = "Fichier Acquisition Intrastat";
                            obj_email.priorite = "143";
                            obj_email.mail_body = "Le fichier " + fich.Name + "  a été traité. <br/><br/> Les fichiers résultats sont présents dans le repertoire " + rep_fichier_traites;
                            obj_email.envoi_mail();
                        }
                        else
                        {
                            LogStep(nameof(Files_Management), "EXPEDITION START", fich.FullName);
                            importation_expedition_with_odbc(fich.FullName);
                            LogStep(nameof(Files_Management), "EXPEDITION END", fich.FullName);

                            obj_email.mail_to = user_info.Adresse_Email;
                            obj_email.mail_from = "Intrastat";
                            obj_email.nom_fichier_xml = "intrastat.xml";
                            obj_email.mail_subject = "Fichier Expedition Intrastat";
                            obj_email.priorite = "143";
                            obj_email.mail_body = "Le fichier " + fich.Name + "  a été traité. <br/><br/> Les fichiers résultats sont présents dans le repertoire " + rep_fichier_traites;
                            obj_email.envoi_mail();
                        }


                        // On supprime le fichier
                        System.IO.File.Delete(fich.FullName);
                    }
                    else
                    {
                        // Le fichier ne comporte pas la bonne extension.
                        // On envoie un email au proprietaire 

                        titi = fich.GetAccessControl();
                        proprietaire = titi.GetOwner(typeof(NTAccount)).Value.ToString();
                        proprietaire = proprietaire.ToLower().Replace(@"corporate\", "");

                        // On envoi un mail au proprietaire 
                        user_info.Get_User_Info(proprietaire);

                        obj_email.mail_to = user_info.Adresse_Email;

                        obj_email.mail_from = "Intrastat";
                        obj_email.nom_fichier_xml = "intrastat.xml";
                        obj_email.mail_subject = "Fichier Intrastat";
                        obj_email.priorite = "143";
                        obj_email.mail_body = "Le fichier " + fich.Name + " n'est pas valide pour ce traitement. (Seul les fichier PUP7902.txt et ARP0292.txt sont autorisés). Merci de le supprimer";
                        obj_email.envoi_mail();

                        LogStep(nameof(Files_Management), "FILE REJECTED", fich.Name.ToUpperInvariant());

                    }
                }

                LogStep(nameof(Files_Management), "SCAN COMPLETE",
                    "Files found=" + tab_fich.Length);

                processingSucceeded = true;
            }
            catch (Exception ex)
            {
                LogTechnicalError(
                    nameof(Files_Management),
                    ex);
                throw;
            }
            finally
            {
                LogStep(
                    nameof(Files_Management),
                    "END",
                    processingSucceeded
                        ? "Completed successfully"
                        : "Failed");
            }
        }
    }
}
