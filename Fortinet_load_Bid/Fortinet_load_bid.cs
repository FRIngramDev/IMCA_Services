using ClosedXML.Excel;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.VisualBasic;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;

namespace IMCA_Services
{
    public class Fortinet_load_bid
    {
        private string pays = "";
        private string sk_valid = "";
        private string name = "";
        private string active = "";
        private string debug = "";
        private string start_date_scan = "";
        private string number_of_mails = "10";
        private string sharedmailbox_name = "";
        private string sharedmailbox_account = "";
        private string sharedmailbox_password = "";
        private string sharedmailbox_server = "";
        private string logs_folder = "";
        private string api_url = "";
        private string sender_email_adress = "";
        private string email_sharing_folder = "";
        private string sharedmailbox_folder_in = "";
        private string sharedmailbox_folder_out = "";
        private string email_sku_management = "";
        private string buyer_email_to = "";
        private string buyer_email_cc = "";
        private string sku_creation_template = "";
        private string email_in_case_of_technical_issue = "";
        private string Append_FC_Quote_ID_To_Bid_Number = "";

        private string global_session_name = "";

        public Fortinet_load_bid()
        {

        }
        public class BidBranchGroup
        {
            public string IM_BRANCH { get; set; } = "";
            public string IM_COMPANY { get; set; } = "";
            public string IM_GROUP { get; set; } = "";
            public string BRANCH_INCLUSION_EXCLUSION { get; set; } = "";
        }
        public class BidDocument
        {
            public string FILE_NAME { get; set; } = "";
            public string LINK { get; set; } = "";
        }
        public class BidEnduser
        {
            public string IM_END_USER_NUMBER { get; set; } = "";
            public string IM_END_USER_NAME { get; set; } = "";
            public string BID_END_USER { get; set; } = "";
        }
        public class BidHeader
        {
            public string USER_ID { get; set; } = "";
            public string BID_NBR { get; set; } = "";
            public string BID_VERSION { get; set; } = "";
            public string BID_TYPE { get; set; } = "";
            public string BID_CURRENCY { get; set; } = "";
            public string MASTER_VENDOR_NUMBER { get; set; } = "";
            public string ACCOUNT_NUMBER { get; set; } = "";
            public string OWNER { get; set; } = "";
            public string BID_COUNTRY { get; set; } = "";
            public string BID_COMBINABILITY { get; set; } = "";
            public string BID_AUTHORIZATION { get; set; } = "";
            public string BID_START_DATE { get; set; } = "";
            public string BID_END_DATE { get; set; } = "";
            public string CLAIM_START_DATE { get; set; } = "";
            public string CLAIM_END_DATE { get; set; } = "";
            public string BID_COMMENT { get; set; } = "";
            public string BID_NBR_ERP { get; set; } = "";
            public string IMPORT_REFERENCE_ID { get; set; } = "";
            public string CONTROL_NBR { get; set; } = "";
            public string VENDOR_BID_CONTACT { get; set; } = "";
            public string VENDOR_PROGRAM_NAME { get; set; } = "";
        }
        public class BidLine
        {
            public string BID_SKU { get; set; } = "";
            public string BID_VPN { get; set; } = "";
            public string BID_CALC_VALUE { get; set; } = "";
            public string BID_CALC_TYPE { get; set; } = "";
            public string BID_CALC_BASIS { get; set; } = "";
            public string BID_CALC_BASIS_VALUE { get; set; } = "";
            public string BID_MAX_QTY { get; set; } = "";
            public string BID_COMBINABILITY { get; set; } = "";
            public string BID_AUTHORIZATION { get; set; } = "";
            public string PRODUCTSET_TYPE { get; set; } = "";
            public string BID_KIT_SKU { get; set; } = "";
            public string BID_BUNDLE_ID { get; set; } = "";
            public string BID_BUNDLE_QTY { get; set; } = "";
            public string CURRENCY { get; set; } = "";
            public string MAX_QTY_PER_RESELLER { get; set; } = "";
            public string MIN_QTY { get; set; } = "";
            public string BID_STARTDATE { get; set; } = "";
            public string BID_ENDDATE { get; set; } = "";
            public string IM_COMMENT { get; set; } = "";
            public string BID_VPN_DESCR { get; set; } = "";
        }
        public class BidReseller
        {
            public string IM_RESELLER_NUMBER { get; set; } = "";
            public string IM_RESELLER_NAME { get; set; } = "";
            public string BID_RESELLER { get; set; } = "";
            public string VAT_CODE { get; set; } = "";
            public string RESELLER_ADDRESS1 { get; set; } = "";
            public string RESELLER_ADDRESS2 { get; set; } = "";
            public string RESELLER_ADDRESS3 { get; set; } = "";
            public string RESELLER_CITY { get; set; } = "";
            public string RESELLER_POST_CODE { get; set; } = "";
            public string RESELLER_COUNTRY { get; set; } = "";
            public string RESELLER_DESCRIPTION { get; set; } = "";
            public string RESELLER_INCLUSION_EXCLUSION { get; set; } = "";
            public string RESELLER_EMAIL { get; set; } = "";

            public string VENDOR_RESELLER_ID { get; set; } = "";
        }
        private static void CopyClass<T>(T copyFrom, T copyTo, bool copyChildren)
        {
            if (copyFrom == null || copyTo == null)
                throw new Exception("Must not specify null parameters");

            var properties = copyFrom.GetType().GetProperties();

            foreach (var p in properties.Where(prop => prop.CanRead && prop.CanWrite))
            {
                if (p.PropertyType.IsClass && p.PropertyType != typeof(string))
                {
                    if (!copyChildren) continue;

                    var destinationClass = Activator.CreateInstance(p.PropertyType);
                    object copyValue = p.GetValue(copyFrom);

                    CopyClass(copyValue, destinationClass, copyChildren);

                    p.SetValue(copyTo, destinationClass);
                }
                else
                {
                    object copyValue = p.GetValue(copyFrom);
                    p.SetValue(copyTo, copyValue);
                }
            }
        }
        public COP_BID DeepCopy<COP_BID>(COP_BID objectToCopy)
        {
            var objectSerialized = JsonConvert.SerializeObject(objectToCopy);
            return JsonConvert.DeserializeObject<COP_BID>(objectSerialized);
        }
        public class COP_BID
        {

            public BidHeader bid_header { get; set; }
            public List<BidLine> bid_lines { get; set; }
            public List<BidEnduser> bid_enduser { get; set; }
            public List<BidReseller> bid_reseller { get; set; }
            public List<BidBranchGroup> bid_branch_group { get; set; }
            public List<BidDocument> bid_documents { get; set; }
        }
        public class Country
        {
            public string country { get; set; }
            public string sk_valid { get; set; }
            public string name { get; set; }
            public string active { get; set; }
            public string debug { get; set; }
            public string number_of_mails { get; set; }
            public string start_date_scan { get; set; }
            public string timer_interval { get; set; }
            public string sharedmailbox_name { get; set; }
            public string sharedmailbox_account { get; set; }
            public string sharedmailbox_password { get; set; }
            public string sharedmailbox_server { get; set; }
            public string sharedmailbox_folder_in { get; set; }
            public string sharedmailbox_folder_out { get; set; }
            public string sender_email_adress { get; set; }
            public string email_sharing_folder { get; set; }
            public string api_url { get; set; }
            public string email_sku_management { get; set; }
            public string email_in_case_of_technical_issue { get; set; }

            public string buyer_email_to { get; set; }
            public string buyer_email_cc { get; set; }

            public string sku_creation_template { get; set; }
            public string Append_FC_Quote_ID_To_Bid_Number { get; set; }
        }
        public class JSON_file
        {
            public string logs_folder { get; set; }
            public List<Country> countries { get; set; }
        }
        public void setParamCountry(string country)
        {
            this.pays = country;
        }
        public void setParamSK_Valid(string sk_valid)
        {
            this.sk_valid = sk_valid;
        }
        public void setParamName(string name)
        {
            this.name = name;
        }
        public void setParamActive(string active)
        {
            this.active = active;
        }
        public void setParamDebug(string debug)
        {
            this.debug = debug;
        }
        public void setEmail_Sku_Management(string email_sku_management)
        {
            this.email_sku_management = email_sku_management;
        }
        public void setEmail_in_case_of_technical_issue(string email_in_case_of_technical_issue)
        {
            this.email_in_case_of_technical_issue = email_in_case_of_technical_issue;
        }
        public void setbuyer_email_to(string buyer_email_to)
        {
            this.buyer_email_to = buyer_email_to;
        }
        public void setbuyer_email_cc(string buyer_email_cc)
        {
            this.buyer_email_cc = buyer_email_cc;
        }
        public void setsku_creation_template(string sku_creation_template)
        {
            this.sku_creation_template = sku_creation_template;
        }
        public void setParamAppend_FC_Quote_ID_To_Bid_Number(string Append_FC_Quote_ID_To_Bid_Number)
        {
            this.Append_FC_Quote_ID_To_Bid_Number = Append_FC_Quote_ID_To_Bid_Number;
        }
        public void setStartDateScan(string start_date_scan)
        {
            this.start_date_scan = start_date_scan;
        }
        public void setSharedMailboxName(string sharedmailbox_name)
        {
            this.sharedmailbox_name = sharedmailbox_name;
        }
        public void setSharedMailboxAccount(string sharedmailbox_account)
        {
            this.sharedmailbox_account = sharedmailbox_account;
        }
        public void setSharedMailboxPassword(string sharedmailbox_password)
        {
            this.sharedmailbox_password = sharedmailbox_password;
        }
        public void setSharedMailboxServer(string sharedmailbox__server)
        {
            this.sharedmailbox_server = sharedmailbox__server;
        }
        public void setlogs_folder(string logs_folder)
        {
            this.logs_folder = logs_folder;
        }
        public void setapi_url(string api_url)
        {
            this.api_url = api_url;
        }
        public void setsender_email_adress(string sender_email_adress)
        {
            this.sender_email_adress = sender_email_adress;
        }
        public void setemail_sharing_folder(string email_sharing_folder)
        {
            this.email_sharing_folder = email_sharing_folder;
        }
        public void setsharedmailbox_folder_in(string sharedmailbox_folder_in)
        {
            this.sharedmailbox_folder_in = sharedmailbox_folder_in;
        }
        public void setsharedmailbox_folder_out(string sharedmailbox_folder_out)
        {
            this.sharedmailbox_folder_out = sharedmailbox_folder_out;
        }
        public void setNumber_of_mails(string number_of_mails)
        {
            this.number_of_mails = number_of_mails;
        }
        public string DecodeFrom64(string encodedData)
        {
            System.Text.UTF8Encoding encoder = new System.Text.UTF8Encoding();
            System.Text.Decoder utf8Decode = encoder.GetDecoder();
            byte[] todecode_byte = Convert.FromBase64String(encodedData);
            int charCount = utf8Decode.GetCharCount(todecode_byte, 0, todecode_byte.Length);
            char[] decoded_char = new char[charCount];
            utf8Decode.GetChars(todecode_byte, 0, todecode_byte.Length, decoded_char, 0);
            string result = new String(decoded_char);
            return result;
        }
        private string Map_VPN_With_SKU_New(string sql_con, string sk_valid, string vpn)
        {
            string sku = "";

            try
            {
                using (SqlConnection con = new SqlConnection(sql_con))
                {
                    con.Open();

                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.Connection = con;
                        cmd.CommandTimeout = 0;

                        cmd.CommandText = "select TOP 1 DSSP_SKU as sku,DSSP_MFR_PART_NBR as mfr_part_nbr FROM openquery(DWHP_IMT,' " +
                                        " select PCM.TAB_PRODUCT.* from PCM.TAB_PRODUCT " +
                                        " INNER JOIN(" +
                                        " SELECT DSSP_MFR_PART_NBR FROM PCM.TAB_PRODUCT " +
                                        " WHERE sk_valid=" + sk_valid +
                                        " AND DSSV_VENDORNBRMASTERNBR=''N670''" +
                                        " AND DSSP_CUSTOMER_CODE like ''FORTINET CUST%'' " +
                                        " AND DSSP_PRODUCT_CLASS<>''X'' AND DSSP_PRODUCT_CLASS <> ''D'' " +
                                        " AND DSSP_CRC_CODE=''PRC'' and DWH_MODCD<>''D''" +
                                        " GROUP BY TAB_PRODUCT.DSSP_MFR_PART_NBR  HAVING COUNT(*) = 1" +
                                        ") " +
                                        " TAB_PRODUCT_BIS on PCM.TAB_PRODUCT.DSSP_MFR_PART_NBR = TAB_PRODUCT_BIS.DSSP_MFR_PART_NBR " +
                                        " where PCM.TAB_PRODUCT.SK_VALID=" + sk_valid + " and PCM.TAB_PRODUCT.DSSV_VENDORNBRMASTERNBR=''N670'' " +
                                        " AND PCM.TAB_PRODUCT.DSSP_CUSTOMER_CODE like ''FORTINET CUST%'' " +
                                        " AND PCM.TAB_PRODUCT.DSSP_PRODUCT_CLASS<>''X'' " +
                                        " AND PCM.TAB_PRODUCT.DSSP_CRC_CODE=''PRC'' AND DSSP_PRODUCT_CLASS <> ''D'' " +
                                        " AND PCM.TAB_PRODUCT.DWH_MODCD<>''D''') DSS " +
                                        " WHERE upper(DSS.DSSP_MFR_PART_NBR)='" + vpn.ToUpper().Trim().Replace("+", "") + "' OR upper(DSS.DSSP_MFR_PART_NBR)='" + vpn.ToUpper().Replace("-", "").Replace("+", "") + "'";

                        SqlDataReader dr;
                        dr = cmd.ExecuteReader();
                        if (dr.HasRows)
                        {
                            while (dr.Read())
                            {
                                sku = dr["sku"].ToString().Trim();
                            }
                        }
                        else
                        {
                            dr.Close();
                            cmd.CommandText = "SELECT DSSP_SKU as sku FROM openquery(DWHP_IMT,'" +
                                                " SELECT DSSP_SKU,DSSP_MFR_PART_NBR FROM PCM.TAB_PRODUCT " +
                                                " WHERE sk_valid=" + sk_valid +
                                                " AND DSSV_VENDORNBRMASTERNBR=''N670''" +
                                                " AND DSSP_CUSTOMER_CODE like ''FORTINET CUST%'' " +
                                                " AND DSSP_PRODUCT_CLASS<>''X'' AND DSSP_PRODUCT_CLASS <> ''D'' " +
                                                " AND DSSP_CRC_CODE=''PRC'' and DWH_MODCD<>''D''" +
                                                " AND (upper(DSSP_MFR_PART_NBR) = ''" + vpn.ToUpper().Trim().Replace("+", "") + "'' OR upper(DSSP_MFR_PART_NBR)= ''" + vpn.ToUpper().Replace("-", "").Replace("+", "") + "'')')";
                            dr = cmd.ExecuteReader();
                            sku = "";
                            if (dr.HasRows)
                            {
                                sku = "MORE THAN ONE SKU FOR THE VPN <b>" + vpn.ToUpper().Trim() + "</b> - SKUS : ";

                                while (dr.Read())
                                {
                                    sku = sku + "<b>" + dr["sku"].ToString().Trim() + "</b>" + " , ";
                                }
                                sku = Strings.Left(sku.Trim(), sku.Trim().Length - 1);

                            }

                        }

                        dr.Close();
                    }

                    con.Close();
                }

            }
            catch (Exception ex)
            {
                WriteToFile("Error Map_VPN_With_SKU_NEW : " + ex.Message);
            }

            return sku;
        }
        public COP_BID traitement_mail_fortinet(string body, string nom_fic_eml, string country_code, string sql_con, string sk_valid)
        {
            string creation_date;
            string nom_contact;
            string nom_eu;
            string date_fin_validite;
            string cust_name;
            string quote_name;
            string contact_email;
            string num_cotavendor;
            string reseller_pos_ID;

            num_cotavendor = "";

            if (body.Contains("Hello EMEA - INGRAM MICRO SAS") == false)
            {
                return null;
            }

            int place_element;
            int place_element_fin_table;
            int place_element_suivant;

            // 
            // Numéro de la cotation
            // 
            num_cotavendor = "";
            place_element = body.IndexOf("has been approved.");

            if (place_element > 0)
            {
                place_element_suivant = body.Substring(0, place_element).LastIndexOf(">");

                if (place_element_suivant > 0)
                    num_cotavendor = body.Substring(place_element_suivant + 1, place_element - place_element_suivant - 1).Trim();
            }

            // 
            // Client Final
            // 
            nom_eu = "";
            place_element = body.IndexOf("End User: ");

            if (place_element > 0)
            {
                place_element_suivant = body.IndexOf("<", place_element);
                nom_eu = body.Substring(place_element, place_element_suivant - place_element).Replace("End User: ", "").Trim();
            }


            // 
            // Quote Name
            // 
            quote_name = "";
            place_element = body.IndexOf("Quote Name: ");

            if (place_element > 0)
            {
                place_element_suivant = body.IndexOf("<", place_element);
                quote_name = body.Substring(place_element, place_element_suivant - place_element).Replace("Quote Name: ", "").Trim();
            }

            // 
            // Nom du revendeur
            // 
            cust_name = "";
            place_element = body.IndexOf("Reseller: ");

            if (place_element > 0)
            {
                place_element_suivant = body.IndexOf("<", place_element);
                cust_name = body.Substring(place_element, place_element_suivant - place_element).Replace("Reseller: ", "").Replace("&amp;", "&").Trim();
            }

            // 
            // Reseller POS ID
            // 
            reseller_pos_ID = "";
            place_element = body.IndexOf("Reseller POS ID: ");

            if (place_element > 0)
            {
                place_element_suivant = body.IndexOf("<", place_element);
                reseller_pos_ID = body.Substring(place_element, place_element_suivant - place_element).Replace("Reseller POS ID: ", "").Trim();
            }


            // 
            // Nom Contact
            // 
            nom_contact = "";
            place_element = body.IndexOf("Partner Contact: ");

            if (place_element > 0)
            {
                place_element_suivant = body.IndexOf("<", place_element);
                nom_contact = body.Substring(place_element, place_element_suivant - place_element).Replace("Partner Contact: ", "").Trim();
            }


            // 
            // Contact Email
            // 
            contact_email = "";
            place_element = body.IndexOf("Partner Contact Email: ");

            if (place_element > 0)
            {
                place_element_suivant = body.IndexOf("<", place_element);
                contact_email = body.Substring(place_element, place_element_suivant - place_element).Replace("Partner Contact Email: ", "").Trim();
            }

            // 
            // Date de Création
            // 
            creation_date = "";
            place_element = body.IndexOf("Approved:");

            if (place_element > 0)
            {
                place_element_suivant = body.IndexOf("<", place_element);
                creation_date = body.Substring(place_element, place_element_suivant - place_element).Replace("Approved:", "").Trim();

                creation_date = creation_date.Substring(6, 4) + "-" + creation_date.Substring(0, 2) + "-" + creation_date.Substring(3, 2);
            }




            // 
            // Date de fin
            // 
            date_fin_validite = "";
            place_element = body.IndexOf("Expires:");

            if (place_element > 0)
            {
                place_element_suivant = body.IndexOf("<", place_element);
                date_fin_validite = body.Substring(place_element, place_element_suivant - place_element).Replace("Expires:", "").Trim();

                date_fin_validite = date_fin_validite.Substring(6, 4) + "-" + date_fin_validite.Substring(0, 2) + "-" + date_fin_validite.Substring(3, 2);
            }

            COP_BID Fortinet_BID = new COP_BID();

            BidHeader fortinet_bid_header = new BidHeader();
            Fortinet_BID.bid_lines = new List<BidLine>();

            fortinet_bid_header.BID_NBR = num_cotavendor;
            fortinet_bid_header.BID_NBR_ERP = num_cotavendor;
            fortinet_bid_header.USER_ID = "bot_" + country_code;
            fortinet_bid_header.BID_COUNTRY = country_code;
            fortinet_bid_header.BID_START_DATE = creation_date;
            fortinet_bid_header.BID_END_DATE = date_fin_validite;
            fortinet_bid_header.BID_VERSION = "1";
            fortinet_bid_header.BID_TYPE = "Vendor Claimed Bid";
            fortinet_bid_header.MASTER_VENDOR_NUMBER = "N670";
            fortinet_bid_header.VENDOR_BID_CONTACT = "";
            fortinet_bid_header.VENDOR_PROGRAM_NAME = "";


            if (quote_name.ToUpper().Contains("RENEW"))
            {
                fortinet_bid_header.BID_COMMENT = "RENEW";
            }
            else
            {
                if (quote_name.ToUpper().Contains("TRADE"))
                {
                    fortinet_bid_header.BID_COMMENT = "TRADE UP";
                }
                else
                {
                    if (quote_name.ToUpper().Contains("COTERM"))
                    {
                        fortinet_bid_header.BID_COMMENT = "COTERM";
                    }
                    else
                    {
                        fortinet_bid_header.BID_COMMENT = "";
                    }
                }
            }

            BidReseller fortinet_bid_reseller = new BidReseller();
            fortinet_bid_reseller.BID_RESELLER = cust_name.Replace("&", " ");
            fortinet_bid_reseller.VENDOR_RESELLER_ID = reseller_pos_ID;
            fortinet_bid_reseller.RESELLER_EMAIL = contact_email;
            fortinet_bid_reseller.RESELLER_DESCRIPTION = nom_contact;
            //fortinet_bid_reseller.IM_RESELLER_NUMBER = map_customer_name_with_customer_number(cust_name);

            // 
            // On traite les lignes de Produits présents dans le mail
            // 

            place_element = body.IndexOf("<table", place_element);
            place_element_fin_table = body.IndexOf("</table>", place_element);

            if (place_element > 0)
            {
                DataTable dt = ConvertHTMLTablesToDataSet(body.Substring(place_element, place_element_fin_table - place_element + 8));

                if (dt != null)
                {
                    foreach (DataRow k in dt.Rows)
                    {

                        switch (k["Disti $ Per Unit"].ToString().Contains("$"))
                        {
                            case true:
                                {
                                    fortinet_bid_header.BID_CURRENCY = "USD";
                                    break;
                                }

                            default:
                                {
                                    fortinet_bid_header.BID_CURRENCY = "EUR";
                                    break;
                                }
                        }

                        if (fortinet_bid_header.BID_COMMENT == "RENEW" && int.Parse(k["QTY"].ToString()) > 1)
                        {
                            int nb_qte = int.Parse(k["QTY"].ToString());

                            for (int indice_qty = 1; indice_qty <= nb_qte; indice_qty++)
                            {
                                BidLine fortinet_bid_line = new BidLine();
                                //fortinet_bid_line.BID_SKU = Map_VPN_With_SKU_New(sql_con,sk_valid,k["SKU"].ToString().Replace("'", "''").Replace("&#43;", "+"));
                                fortinet_bid_line.BID_VPN = k["SKU"].ToString().Replace("'", "''").Replace("&#43;", "+");
                                fortinet_bid_line.MIN_QTY = "1";
                                fortinet_bid_line.BID_MAX_QTY = "1";
                                fortinet_bid_line.BID_CALC_TYPE = "FIXED";
                                fortinet_bid_line.BID_STARTDATE = creation_date;
                                fortinet_bid_line.BID_ENDDATE = date_fin_validite;
                                fortinet_bid_line.BID_CALC_VALUE = k["Disti $ Per Unit"].ToString().Replace(",", ".").Replace("$", "");
                                fortinet_bid_line.IM_COMMENT = k["FC Quote ID"].ToString();
                                fortinet_bid_line.CURRENCY = fortinet_bid_header.BID_CURRENCY;
                                fortinet_bid_line.BID_VPN_DESCR = k["Description"].ToString().Trim();
                                Fortinet_BID.bid_lines.Add(fortinet_bid_line);
                            }
                        }
                        else
                        {
                            BidLine fortinet_bid_line = new BidLine();
                            //fortinet_bid_line.BID_SKU = Map_VPN_With_SKU_New(sql_con, sk_valid,k["SKU"].ToString().Replace("'", "''").Replace("&#43;", "+"));
                            fortinet_bid_line.BID_VPN = k["SKU"].ToString().Replace("'", "''").Replace("&#43;", "+");
                            fortinet_bid_line.MIN_QTY = "1";
                            fortinet_bid_line.BID_MAX_QTY = k["QTY"].ToString();
                            fortinet_bid_line.BID_CALC_TYPE = "FIXED";
                            fortinet_bid_line.BID_STARTDATE = creation_date;
                            fortinet_bid_line.BID_ENDDATE = date_fin_validite;
                            fortinet_bid_line.BID_CALC_VALUE = k["Disti $ Per Unit"].ToString().Replace(",", ".").Replace("$", "");
                            fortinet_bid_line.IM_COMMENT = k["FC Quote ID"].ToString();
                            fortinet_bid_line.CURRENCY = fortinet_bid_header.BID_CURRENCY;
                            fortinet_bid_line.BID_VPN_DESCR = k["Description"].ToString().Trim();
                            Fortinet_BID.bid_lines.Add(fortinet_bid_line);

                        }

                    }

                }
            }


            // Fill the class

            Fortinet_BID.bid_enduser = new List<BidEnduser>();
            BidEnduser fortinet_bid_enduser = new BidEnduser();
            fortinet_bid_enduser.BID_END_USER = nom_eu;
            Fortinet_BID.bid_enduser.Add(fortinet_bid_enduser);


            Fortinet_BID.bid_documents = new List<BidDocument>();
            BidDocument fortinet_bid_documents = new BidDocument();
            fortinet_bid_documents.FILE_NAME = Path.GetFileName(nom_fic_eml);
            fortinet_bid_documents.LINK = nom_fic_eml;
            Fortinet_BID.bid_documents.Add(fortinet_bid_documents);

            Fortinet_BID.bid_branch_group = new List<BidBranchGroup>();
            BidBranchGroup fortinet_bid_branch_group = new BidBranchGroup();
            Fortinet_BID.bid_branch_group.Add(fortinet_bid_branch_group);


            Fortinet_BID.bid_header = new BidHeader();
            Fortinet_BID.bid_header = fortinet_bid_header;

            Fortinet_BID.bid_reseller = new List<BidReseller>();
            Fortinet_BID.bid_reseller.Add(fortinet_bid_reseller);

            return Fortinet_BID;

        }
        private string create_fortinet_BID_in_COP(string postData, ExchangeService service, string BID_NBR_ERP = "", int nb_trials = 5)
        {
            int trials = 1;
            string result = "";
            ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;

            System.Net.ServicePointManager.ServerCertificateValidationCallback += delegate (object sender, System.Security.Cryptography.X509Certificates.X509Certificate certificate,
                        System.Security.Cryptography.X509Certificates.X509Chain chain,
                        System.Net.Security.SslPolicyErrors sslPolicyErrors)
            {
                return true;
            };

            while (trials <= nb_trials)
            {

                if (debug.ToUpper() == "TRUE")
                {
                    WriteToFile("       api_url : " + api_url);
                }
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(api_url);
                request.ContentType = "application/x-www-form-urlencoded";
                request.Method = "POST";
                request.ContentType = "application/json";

                byte[] bytes = Encoding.UTF8.GetBytes(postData);
                request.ContentLength = bytes.Length;

                try
                {
                    Stream requestStream = request.GetRequestStream();
                    requestStream.Write(bytes, 0, bytes.Length);
                    HttpWebResponse response = (HttpWebResponse)request.GetResponse();

                    using (StreamReader rdr = new StreamReader(response.GetResponseStream()))
                    {
                        result = rdr.ReadToEnd();
                    }
                    requestStream.Close();

                    if (debug.ToUpper() == "TRUE")
                    {
                        WriteToFile("       BID COP created");
                        WriteToFile("       JSON : " + postData);
                    }

                    break;

                }
                catch (Exception ex)
                {
                    result = "ERROR";
                    WriteToFile("  POST error :" + ex.Message);
                    WriteToFile("       JSON : " + postData);

                    if (email_in_case_of_technical_issue.Trim() != "" && email_in_case_of_technical_issue.Contains("@"))
                    {
                        EmailMessage new_email = new EmailMessage(service);
                        new_email.Subject = "FORTINET - " + BID_NBR_ERP + " - Unable to create the BID (BidIT) - " + trials.ToString() + " in " + nb_trials.ToString() + " tries";
                        new_email.Body = "Hi <br/><br/>We encountered an error (" + ex.Message + ") while creating the BID in BidIT<br/><br/>Please find below the posted JSON<br/><br/>" + postData + "<br/><br/>Regards";
                        new_email.ToRecipients.Add(email_in_case_of_technical_issue);
                        new_email.SendAndSaveCopy(WellKnownFolderName.SentItems);
                    }

                    // We wait 30 secs
                    System.Threading.Thread.Sleep(30000);


                    //if (!ex.Message.ToLower().Contains("timed out"))
                    //    {
                    //    break;
                    //    }

                }

                trials++;
            }

            return result;
        }
        private string create_fortinet_BID_in_COP_with_Graph(string postData, GraphServiceClient service, string BID_NBR_ERP = "", int nb_trials = 5)
        {
            int trials = 1;
            string result = "";
            ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;

            System.Net.ServicePointManager.ServerCertificateValidationCallback += delegate (object sender, System.Security.Cryptography.X509Certificates.X509Certificate certificate,
                        System.Security.Cryptography.X509Certificates.X509Chain chain,
                        System.Net.Security.SslPolicyErrors sslPolicyErrors)
            {
                return true;
            };

            while (trials <= nb_trials)
            {

                if (debug.ToUpper() == "TRUE")
                {
                    WriteToFile("       api_url : " + api_url);
                }
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(api_url);
                request.ContentType = "application/x-www-form-urlencoded";
                request.Method = "POST";
                request.ContentType = "application/json";

                byte[] bytes = Encoding.UTF8.GetBytes(postData);
                request.ContentLength = bytes.Length;

                try
                {
                    Stream requestStream = request.GetRequestStream();
                    requestStream.Write(bytes, 0, bytes.Length);
                    HttpWebResponse response = (HttpWebResponse)request.GetResponse();

                    using (StreamReader rdr = new StreamReader(response.GetResponseStream()))
                    {
                        result = rdr.ReadToEnd();
                    }
                    requestStream.Close();

                    if (debug.ToUpper() == "TRUE")
                    {
                        WriteToFile("       BID COP created");
                        WriteToFile("       JSON : " + postData);
                    }

                    break;

                }
                catch (Exception ex)
                {
                    result = "ERROR";
                    WriteToFile("  POST error :" + ex.Message);
                    WriteToFile("       JSON : " + postData);

                    if (email_in_case_of_technical_issue.Trim() != "" && email_in_case_of_technical_issue.Contains("@"))
                    {
                        var requestBody = new Microsoft.Graph.Me.SendMail.SendMailPostRequestBody
                        {
                            Message = new Message
                            {
                                Subject = "FORTINET - " + BID_NBR_ERP + " - Unable to create the BID (BidIT)",
                                Body = new ItemBody
                                {
                                    ContentType = Microsoft.Graph.Models.BodyType.Html,
                                    Content = "Hi <br/><br/>We encountered an error (" + ex.Message + ") while creating the BID in BidIT<br/><br/>Please find below the posted JSON<br/><br/>" + postData + "<br/><br/>Regards",
                                },
                                ToRecipients = new List<Recipient>
                                {
                                    new Recipient
                                    {
                                        EmailAddress = new Microsoft.Graph.Models.EmailAddress
                                        {
                                            Address =email_in_case_of_technical_issue,
                                        },
                                    },
                                }
                            },
                            SaveToSentItems = true,
                        };

                        service.Me.SendMail.PostAsync(requestBody);

                    }

                    // We wait 30 secs
                    System.Threading.Thread.Sleep(30000);

                }

                trials++;
            }

            return result;
        }
        private string create_JSON(COP_BID COP)
        {

            string json = JsonConvert.SerializeObject(COP, Newtonsoft.Json.Formatting.Indented);

            return json;

        }
        public string CleanFileName(string fileName)
        {
            return string.Join("_", fileName.Split(Path.GetInvalidFileNameChars())).Trim();
        }
        public string CleanPathName(string fileName)
        {
            return string.Join("_", fileName.Split(Path.GetInvalidPathChars())).Trim();
        }
        private FolderId Find_Folder(ExchangeService server, string name_folder_to_find)
        {
            Microsoft.Exchange.WebServices.Data.FolderView view = new Microsoft.Exchange.WebServices.Data.FolderView(10);
            // Dim filter = New SearchFilter.IsEqualTo(FolderSchema.DisplayName, name_folder_to_find)

            view.PropertySet = new PropertySet(BasePropertySet.IdOnly, FolderSchema.DisplayName);
            view.Traversal = FolderTraversal.Shallow;

            FindFoldersResults findFolderResults = server.FindFolders(WellKnownFolderName.Inbox, view);
            // find specific folder
            foreach (Microsoft.Exchange.WebServices.Data.Folder f in findFolderResults)
            {
                if (f.DisplayName.ToLower() == name_folder_to_find.ToLower())
                {
                    Console.WriteLine(f.Id);
                    return f.Id;
                }
            }

            return null;
        }
        private ExchangeService Connexion_Au_Service_Exchange_O365(string la_boite_mail)
        {
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;

            // On utilise EWS API MANAGED REFERENCE 2.1
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            service.Timeout = 120000;
            //service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, la_boite_mail);

            class_dev_tools.Ews_Modern_Auth Ews_Modern_Auth = new class_dev_tools.Ews_Modern_Auth();
            Ews_Modern_Auth.boite_mail = la_boite_mail;
            Ews_Modern_Auth.email = sharedmailbox_account;
            Ews_Modern_Auth.password = sharedmailbox_password;
            Ews_Modern_Auth.serveur_mail = sharedmailbox_server;
            service = Ews_Modern_Auth.Get_EWS_Service();

            return service;
        }
        private GraphServiceClient Connexion_Microsoft_Graph()
        {
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;

            // On utilise EWS API MANAGED REFERENCE 2.1
            GraphServiceClient GraphService;

            class_dev_tools.Ews_Modern_Auth Ews_Modern_Auth = new class_dev_tools.Ews_Modern_Auth();

            GraphService = Ews_Modern_Auth.Get_Graph_Service();

            return GraphService;
        }
        public void WriteToFile(string message)
        {
            if (!Directory.Exists(logs_folder))
            {
                Directory.CreateDirectory(logs_folder);
            }
            string filepath = logs_folder + "\\IMCA_" + global_session_name + "_" + DateTime.Now.Date.ToString("dd_MM_yyyy") + "_" + pays + "_FORTINET_BID_TO_COP.txt";

            if (!File.Exists(filepath))
            {
                using (StreamWriter fic = File.CreateText(filepath))
                {
                    fic.WriteLine(message);
                }
            }
            else
            {
                using (StreamWriter fic = File.AppendText(filepath))
                {
                    fic.WriteLine(message);
                }
            }

        }
        private DataTable ConvertHTMLTablesToDataSet(string HTML)
        {
            // Declarations   
            DataTable dt = null;
            DataRow dr = null;
            string TableExpression = "<table[^>]*>(.*?)</table>";
            // Dim HeaderExpression As String = "<th[^>]*>(.*?)</th>"
            string HeaderExpression = "<th>(.*?)</th>";
            string RowExpression = "<tr[^>]*>(.*?)</tr>";
            // Dim ColumnExpression As String = "<td[^>]*>(.*?)</td>"
            string ColumnExpression = "<td>(.*?)</td>";
            string ColumnExpression_th = "<th>(.*?)</th>";
            int iCurrentColumn = 0;
            int iCurrentRow = 0;
            string expr1_a_supprimer = "<p class=" + Convert.ToChar(34) + "MsoNormal" + Convert.ToChar(34) + " align=" + Convert.ToChar(34) + "center" + Convert.ToChar(34) + " style=" + Convert.ToChar(34) + "text-align:center" + Convert.ToChar(34) + "><b><span style=" + Convert.ToChar(34) + "font-size:8.5pt" + Convert.ToChar(34) + ">";
            string expr2_a_supprimer = "<o:p></o:p></span></b></p>";
            string expr3_a_supprimer = "<o:p></o:p></span></p>";
            string expr4_a_supprimer = "<p><span style=" + Convert.ToChar(34) + "font-size:8.5pt" + Convert.ToChar(34) + ">";
            string expr5_a_supprimer = "</span></b><b><span style=" + Convert.ToChar(34) + "font-size:8.5pt" + Convert.ToChar(34) + ">";
            string expr6_a_supprimer = "<p class=" + Convert.ToChar(34) + "MsoNormal" + Convert.ToChar(34) + " align=" + Convert.ToChar(34) + "center" + Convert.ToChar(34) + " style=" + Convert.ToChar(34) + "text-align:center" + Convert.ToChar(34) + "><span style=" + Convert.ToChar(34) + "font-size:8.5pt" + Convert.ToChar(34) + ">";
            string expr7_a_supprimer = " style=" + Convert.ToChar(34) + "border:none;background:" + Convert.ToChar(35) + "CCCCCC;padding:.75pt .75pt .75pt .75pt" + Convert.ToChar(34);
            string expr8_a_supprimer = " style=" + Convert.ToChar(34) + "border:none;padding:.75pt .75pt .75pt .75pt" + Convert.ToChar(34);
            string expr9_a_supprimer = "</span><span style=" + Convert.ToChar(34) + "font-size:8.5pt" + Convert.ToChar(34) + ">";
            string expr10_a_supprimer = "<p class=" + Convert.ToChar(34) + "MsoNormal" + Convert.ToChar(34) + " align=" + Convert.ToChar(34) + "center" + Convert.ToChar(34) + " style=" + Convert.ToChar(34) + "text-align:center" + Convert.ToChar(34) + "><b><span style=" + Convert.ToChar(34) + "font-size:8.5pt;color:black" + Convert.ToChar(34) + ">";
            string expr11_a_supprimer = "<p align=" + Convert.ToChar(34) + "center" + Convert.ToChar(34) + " style=" + Convert.ToChar(34) + "text-align:center;background:#FF7068" + Convert.ToChar(34) + "><b><span style=" + Convert.ToChar(34) + "font-size:8.5pt;color:black" + Convert.ToChar(34) + ">";
            string expr12_a_supprimer = "<p style=" + Convert.ToChar(34) + "background-color:#FF7068" + Convert.ToChar(34) + ">";
            string expr13_a_supprimer = "<p style=" + Convert.ToChar(34) + "text-align: left" + Convert.ToChar(34) + ">";
            string expr14_a_supprimer = "<p align=" + Convert.ToChar(34) + "center" + Convert.ToChar(34) + " style=" + Convert.ToChar(34) + "text-align:center;background:#FF7068" + Convert.ToChar(34) + "><span style=" + Convert.ToChar(34) + "font-size:8.5pt;color:black" + Convert.ToChar(34) + ">";
            string expr15_a_remplacer = "Disti Ext $" + Convert.ToChar(13) + Convert.ToChar(10) + "</td>" + Convert.ToChar(13) + Convert.ToChar(10) + "</tr>";
            string expr20_a_remplacer = "Disti Ext $</span></b><o:p></o:p></p>" + Convert.ToChar(13) + Convert.ToChar(10) + "</td>" + Convert.ToChar(13) + Convert.ToChar(10) + "</tr>";
            string expr16_a_remplacer = "<tbody>" + Convert.ToChar(13) + Convert.ToChar(10) + "<tr>";
            string expr17_a_remplacer = "</tbody>";
            string expr18_a_remplacer = "</span></b><o:p></o:p></p>";
            string expr19_a_remplacer = "</span><o:p></o:p></p>";
            string expr21_a_remplacer = "</p>";
            string expr22_a_remplacer = " style=" + Convert.ToChar(34) + Convert.ToChar(34);
            string expr23_a_remplacer = "<th><th>";
            string expr24_a_remplacer = "<tr><th>";
            string expr25_a_remplacer = "</th></tr>";
            string expr26_a_remplacer = "</th><th>";
            string expr27_a_supprimer = "class=" + Convert.ToChar(34) + "MsoNormalTable" + Convert.ToChar(34) + " border=" + Convert.ToChar(34) + "1" + Convert.ToChar(34) + " cellspacing=" + Convert.ToChar(34) + "3" + Convert.ToChar(34) + " cellpadding=" + Convert.ToChar(34) + "0" + Convert.ToChar(34) + " style=" + Convert.ToChar(34) + "border:solid black 1.0pt" + Convert.ToChar(34);
            string expr28_a_supprimer = "<p style=" + Convert.ToChar(34) + "text-align:left" + Convert.ToChar(34) + ">";

            HTML = HTML.Replace(expr1_a_supprimer, "");
            HTML = HTML.Replace(expr2_a_supprimer, "");
            HTML = HTML.Replace(expr3_a_supprimer, "");
            HTML = HTML.Replace(expr4_a_supprimer, "");
            HTML = HTML.Replace(expr5_a_supprimer, "");
            HTML = HTML.Replace(expr6_a_supprimer, "");
            HTML = HTML.Replace(expr7_a_supprimer, "");
            HTML = HTML.Replace(expr8_a_supprimer, "");
            HTML = HTML.Replace(expr9_a_supprimer, "");
            HTML = HTML.Replace(expr10_a_supprimer, "");
            HTML = HTML.Replace(expr11_a_supprimer, "");
            HTML = HTML.Replace(expr12_a_supprimer, "");
            HTML = HTML.Replace(expr13_a_supprimer, "");
            HTML = HTML.Replace(expr14_a_supprimer, "");
            HTML = HTML.Replace(expr15_a_remplacer, "Disti Ext $</td></th>");
            HTML = HTML.Replace(expr20_a_remplacer, "Disti Ext $</td></th>");
            HTML = HTML.Replace(expr16_a_remplacer, "<th>");
            HTML = HTML.Replace("\r\n", "");
            HTML = HTML.Replace(expr17_a_remplacer, "");
            HTML = HTML.Replace(expr18_a_remplacer, "");
            HTML = HTML.Replace(expr19_a_remplacer, "");
            HTML = HTML.Replace(expr21_a_remplacer, "");
            HTML = HTML.Replace(expr22_a_remplacer, "");
            HTML = HTML.Replace(expr23_a_remplacer, "<tr><th>");
            HTML = HTML.Replace(expr24_a_remplacer, "<th><td>");
            HTML = HTML.Replace(expr25_a_remplacer, "</td></th>");
            HTML = HTML.Replace(expr26_a_remplacer, "</td><td>");
            HTML = HTML.Replace(expr27_a_supprimer, "");
            HTML = HTML.Replace(expr28_a_supprimer, "");
            // Get a match for all the tables in the HTML   
            MatchCollection Tables = Regex.Matches(HTML, TableExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase);


            // Loop through each table element   
            foreach (Match Table in Tables)
            {

                // Reset the current row counter and the header flag   
                iCurrentRow = 0;

                // Add a new table to the DataSet   
                dt = new DataTable();


                // Create the relevant amount of columns for this table (use the headers if they exist, otherwise use default names)   
                if (Table.Value.Contains("<th") | Table.Value.Contains("tbody"))
                {
                    // Set the HeadersExist flag   

                    // Get a match for all the rows in the table   
                    MatchCollection Headers = Regex.Matches(Table.Value, HeaderExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase);

                    if (Headers.Count == 0)
                        Headers = Regex.Matches(Table.Value, "<tr>(.*?)</tr>", RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase);

                    // Loop through each header element   
                    foreach (Match Header in Headers)
                    {

                        // Get a match for all the columns in the row   
                        MatchCollection Columns = Regex.Matches(Header.Groups[1].ToString(), ColumnExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase);

                        if (Columns.Count > 0)
                        {
                            foreach (Match Column in Columns)
                            {
                                if (!dt.Columns.Contains(Column.Groups[1].ToString()))
                                    dt.Columns.Add(Column.Groups[1].ToString().Replace("&nbsp;", ""));
                            }
                        }
                        else
                        {
                            Columns = Regex.Matches(Header.Groups[1].ToString(), ColumnExpression_th, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase);
                            foreach (Match Column in Columns)
                            {
                                if (!dt.Columns.Contains(Column.Groups[1].ToString()))
                                    dt.Columns.Add(Column.Groups[1].ToString().Replace("&nbsp;", ""));
                            }
                        }

                        if (Headers.Count > 1)
                            break;
                    }
                }
                else
                    for (int iColumns = 1; iColumns <= Regex.Matches(Regex.Matches(Regex.Matches(Table.Value, TableExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase).ToString(), RowExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase).ToString(), ColumnExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase).Count; iColumns++)
                        dt.Columns.Add("Column " + iColumns);


                // Get a match for all the rows in the table   
                MatchCollection Rows = Regex.Matches(Table.Value, RowExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase);


                // Loop through each row element   
                foreach (Match Row in Rows)
                {


                    // Only loop through the row if it isn't a header row   
                    // If Not (iCurrentRow = 0 And HeadersExist = True) Then

                    // Create a new row and reset the current column counter   
                    dr = dt.NewRow();
                    iCurrentColumn = 0;

                    // Get a match for all the columns in the row   
                    MatchCollection Columns = Regex.Matches(Row.Value, ColumnExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase);


                    // Loop through each column element   
                    foreach (Match Column in Columns)
                    {
                        // Add the value to the DataRow   
                        dr[iCurrentColumn] = Column.Groups[1].ToString().Replace("&nbsp;", "");

                        // Increase the current column    
                        iCurrentColumn += 1;
                    }

                    if (dr[0].ToString().ToUpper() != "FC QUOTE ID")
                        // Add the DataRow to the DataTable   
                        dt.Rows.Add(dr);

                    // End If


                    // Increase the current row counter   
                    iCurrentRow += 1;
                }
            }

            return dt;
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
                    cmd.CommandTimeout = 0;

                    cmd.CommandText = "SELECT isnull(VALUE,'') as VALUE from [PCM_TAB_IMCA_PARAMETER_GLOBAL] where SK_VALID=0 AND PARAMETER='" + param_name + "'";
                    ret = cmd.ExecuteScalar().ToString();
                }

                con.Close();
            }

            return ret;

        }
        public void send_missing_or_duplicates_skus_to_sku_management_or_buyer_with_Graph(GraphServiceClient service, string sql_con, COP_BID Fortinet_BID)
        {

            string vpn_list = "";
            string sku = "";
            int sku_line = 2;
            XLWorkbook wb = null;
            IXLWorksheet ws = null;
            Boolean top_header = false;
            Boolean top_clean_up_the_product_list = false;
            string sku_to_clean = "";

            string Sku_Creation_FileName = email_sharing_folder + "\\" + CleanFileName(Fortinet_BID.bid_header.BID_NBR_ERP) + ".xlsx";

            Sku_Creation_FileName = CleanPathName(Sku_Creation_FileName);
            Sku_Creation_FileName = Sku_Creation_FileName.Replace(";", "");

            foreach (BidLine lignes in Fortinet_BID.bid_lines)
            {
                if (lignes.BID_VPN.ToString().Trim() != "" && lignes.BID_SKU.ToString().Trim() == "")
                {
                    if (vpn_list.Contains(lignes.BID_VPN.ToString().Trim()) == false)
                    {
                        sku = Map_VPN_With_SKU_New(sql_con, sk_valid, lignes.BID_VPN.ToString().Trim());

                        if (sku.ToString().Trim().Contains("MORE THAN ONE SKU FOR THE VPN") == true)
                        {
                            top_clean_up_the_product_list = true;
                            sku_to_clean = sku_to_clean + sku + ((char)13);
                        }


                        if (sku.ToString().Trim() == "" && sku.ToString().Trim().Contains("MORE THAN ONE SKU FOR THE VPN") == false)
                        {
                            //
                            // Missing SKU
                            //

                            if (top_header == false)
                            {
                                wb = new XLWorkbook(CleanPathName(sku_creation_template));
                                ws = wb.Worksheet("Feuil1");
                                top_header = true;
                            }

                            if (lignes.BID_VPN.ToString().Trim().Length > 20)
                            {
                                ws.Cell(sku_line, 3).Value = lignes.BID_VPN.ToString().Trim().Replace("-", "");
                            }
                            else
                            {
                                ws.Cell(sku_line, 3).Value = lignes.BID_VPN.ToString().Trim();
                            }
                            ws.Cell(sku_line, 17).Value = "PRC";
                            ws.Cell(sku_line, 18).Value = "FORTINET CUST.";
                            ws.Cell(sku_line, 21).Value = lignes.BID_VPN_DESCR.ToString().Trim();

                            sku_line += 1;

                        }

                        vpn_list += lignes.BID_VPN.ToString().Trim() + "!";

                    }
                }

            }

            if (top_header == true)
            {

                if (debug.ToUpper() == "TRUE")
                {
                    WriteToFile("       Send Missing skus to : " + email_sku_management.Trim() + " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                }

                if (System.IO.File.Exists(Sku_Creation_FileName))
                {
                    try
                    {
                        System.IO.File.Delete(Sku_Creation_FileName);
                    }
                    catch (Exception)
                    {

                        throw;
                    }

                }

                wb.SaveAs(Sku_Creation_FileName);
                ws.Dispose();
                wb.Dispose();

                //
                // We send an email to sku management
                //
                byte[] contentBytes = System.IO.File.ReadAllBytes(Sku_Creation_FileName);
                string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                var requestBody = new Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody
                {
                    Message = new Message
                    {
                        Subject = "FORTINET - " + Fortinet_BID.bid_header.BID_NBR_ERP + " - Skus to create",
                        Body = new ItemBody
                        {
                            ContentType = Microsoft.Graph.Models.BodyType.Html,
                            Content = "Hi <br/><br/>Please create the following skus <br/><br/>Regards",
                        },
                        ToRecipients = new List<Recipient>
                                {
                                    new Recipient
                                    {
                                        EmailAddress = new Microsoft.Graph.Models.EmailAddress
                                        {
                                            Address =email_sku_management,
                                        },
                                    },
                                }
                        ,
                        Attachments = new List<Microsoft.Graph.Models.Attachment>
                            {
                                new Microsoft.Graph.Models.FileAttachment
                                {
                                    Name = System.IO.Path.GetFileName(Sku_Creation_FileName),
                                    ContentType = contentType,
                                    ContentBytes = contentBytes,
                                },
                            },
                    },
                    SaveToSentItems = true,
                };

                service.Users[sharedmailbox_name].SendMail.PostAsync(requestBody).GetAwaiter().GetResult();

            }

            //if (System.IO.File.Exists(Sku_Creation_FileName))
            //{
            //    try
            //    {
            //        System.IO.File.Delete(Sku_Creation_FileName);
            //    }
            //    catch (Exception)
            //    {

            //        throw;
            //    }

            //}

            if (top_clean_up_the_product_list == true && buyer_email_to.Trim() != "" && buyer_email_to.Contains("@"))
            {
                if (debug.ToUpper() == "TRUE")
                {
                    WriteToFile("       Send Duplicates skus to : " + buyer_email_to.Trim() + " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                }

                if (!buyer_email_cc.Contains("@"))
                {
                    buyer_email_cc = "";
                }
                else
                {
                    WriteToFile("       Send Duplicates skus CC to : " + buyer_email_cc.Trim() + " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                }

                //
                // We send an email to sku management
                //
                var requestBody = new Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody
                {
                    Message = new Message
                    {
                        Subject = "FORTINET - " + Fortinet_BID.bid_header.BID_NBR_ERP + " - Duplicates skus for the same VPN. Please Check and Clean Up",
                        Body = new ItemBody
                        {
                            ContentType = Microsoft.Graph.Models.BodyType.Html,
                            Content = "Hi <br/><br/>Please check (and clean up) the following skus <br/><br/>" + sku_to_clean + "<br/><br/>Regards",
                        },
                        ToRecipients = new List<Recipient>
                                {
                                    new Recipient
                                    {
                                        EmailAddress = new Microsoft.Graph.Models.EmailAddress
                                        {
                                            Address =buyer_email_to,
                                        },
                                    },
                                },
                        CcRecipients = new List<Recipient>
                            {
                                new Recipient
                                {
                                    EmailAddress = new Microsoft.Graph.Models.EmailAddress
                                        {
                                            Address = buyer_email_cc,
                                        },
                                }
                            },
                    },
                    SaveToSentItems = true
                };

                service.Users[sharedmailbox_name].SendMail.PostAsync(requestBody).GetAwaiter().GetResult();
            }

        }
        public void send_missing_or_duplicates_skus_to_sku_management_or_buyer(ExchangeService service, string sql_con, COP_BID Fortinet_BID)
        {

            string vpn_list = "";
            string sku = "";
            int sku_line = 2;
            XLWorkbook wb = null;
            IXLWorksheet ws = null;
            Boolean top_header = false;
            Boolean top_clean_up_the_product_list = false;
            string sku_to_clean = "";

            string Sku_Creation_FileName = email_sharing_folder + "\\" + CleanFileName(Fortinet_BID.bid_header.BID_NBR_ERP) + ".xlsx";

            Sku_Creation_FileName = CleanPathName(Sku_Creation_FileName);
            Sku_Creation_FileName = Sku_Creation_FileName.Replace(";", "");

            foreach (BidLine lignes in Fortinet_BID.bid_lines)
            {
                if (lignes.BID_VPN.ToString().Trim() != "" && lignes.BID_SKU.ToString().Trim() == "")
                {
                    if (vpn_list.Contains(lignes.BID_VPN.ToString().Trim()) == false)
                    {
                        sku = Map_VPN_With_SKU_New(sql_con, sk_valid, lignes.BID_VPN.ToString().Trim());

                        if (sku.ToString().Trim().Contains("MORE THAN ONE SKU FOR THE VPN") == true)
                        {
                            top_clean_up_the_product_list = true;
                            sku_to_clean = sku_to_clean + sku + ((char)13);
                        }


                        if (sku.ToString().Trim() == "" && sku.ToString().Trim().Contains("MORE THAN ONE SKU FOR THE VPN") == false)
                        {
                            //
                            // Missing SKU
                            //

                            if (top_header == false)
                            {
                                wb = new XLWorkbook(CleanPathName(sku_creation_template));
                                ws = wb.Worksheet("Feuil1");
                                top_header = true;
                            }

                            if (lignes.BID_VPN.ToString().Trim().Length > 20)
                            {
                                ws.Cell(sku_line, 3).Value = lignes.BID_VPN.ToString().Trim().Replace("-", "");
                            }
                            else
                            {
                                ws.Cell(sku_line, 3).Value = lignes.BID_VPN.ToString().Trim();
                            }
                            ws.Cell(sku_line, 17).Value = "PRC";
                            ws.Cell(sku_line, 18).Value = "FORTINET CUST.";
                            ws.Cell(sku_line, 21).Value = lignes.BID_VPN_DESCR.ToString().Trim();

                            sku_line += 1;

                        }

                        vpn_list += lignes.BID_VPN.ToString().Trim() + "!";

                    }
                }

            }

            if (top_header == true)
            {

                if (debug.ToUpper() == "TRUE")
                {
                    WriteToFile("       Send Missing skus to : " + email_sku_management.Trim() + " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                }

                if (System.IO.File.Exists(Sku_Creation_FileName))
                {
                    try
                    {
                        System.IO.File.Delete(Sku_Creation_FileName);
                    }
                    catch (Exception)
                    {

                        throw;
                    }

                }

                wb.SaveAs(Sku_Creation_FileName);
                wb.Dispose();

                //
                // We send an email to sku management
                //
                EmailMessage new_email = new EmailMessage(service);
                new_email.Subject = "FORTINET - " + Fortinet_BID.bid_header.BID_NBR_ERP + " - Skus to create";
                new_email.Body = "Hi <br/><br/>Please create the following skus <br/><br/>Regards";
                new_email.ToRecipients.Add(email_sku_management);
                new_email.Attachments.AddFileAttachment(Sku_Creation_FileName);
                new_email.SendAndSaveCopy(WellKnownFolderName.SentItems);

            }

            if (top_clean_up_the_product_list == true && buyer_email_to.Trim() != "" && buyer_email_to.Contains("@"))
            {
                if (debug.ToUpper() == "TRUE")
                {
                    WriteToFile("       Send Duplicates skus to : " + buyer_email_to.Trim() + " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                }

                //
                // We send an email to sku management
                //
                EmailMessage new_email = new EmailMessage(service);
                new_email.Subject = "FORTINET - " + Fortinet_BID.bid_header.BID_NBR_ERP + " - Duplicates skus for the same VPN. Please Check and Clean Up";
                new_email.Body = "Hi <br/><br/>Please check (and clean up) the following skus <br/><br/>" + sku_to_clean + "<br/><br/>Regards";
                new_email.ToRecipients.Add(buyer_email_to);

                if (buyer_email_cc.Trim() != "" && buyer_email_cc.Contains("@"))
                {
                    new_email.CcRecipients.Add(buyer_email_cc);
                }

                new_email.SendAndSaveCopy(WellKnownFolderName.SentItems);

            }


        }
        public void Read_Email(string sql_con, string logs, string session_name)
        {
            string global_parameters = "";
            string service_path = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
            int nb_mail = 0;
            Item Item_move;
            setlogs_folder(service_path + "\\" + logs);

            global_session_name = session_name;

            // Get the global parameters of the ACTION
            global_parameters = get_IMCA_paramters(sql_con, "FORTINET_BID_LOAD");

            var param = new JSON_file();
            param = JsonConvert.DeserializeObject<JSON_file>(global_parameters);

            foreach (var p in param.countries) // For each country
            {

                setParamCountry(p.country);
                setParamSK_Valid(p.sk_valid);
                setParamName(p.name);
                setParamActive(p.active);
                setParamDebug(p.debug);
                setStartDateScan(p.start_date_scan);
                setSharedMailboxName(p.sharedmailbox_name);
                setSharedMailboxAccount(p.sharedmailbox_account);
                setSharedMailboxPassword(DecodeFrom64(p.sharedmailbox_password));
                setSharedMailboxServer(p.sharedmailbox_server);
                setsharedmailbox_folder_in(p.sharedmailbox_folder_in);
                setNumber_of_mails(p.number_of_mails);
                setsharedmailbox_folder_out(p.sharedmailbox_folder_out);
                setsender_email_adress(p.sender_email_adress);
                setemail_sharing_folder(p.email_sharing_folder);
                setapi_url(p.api_url);
                setEmail_Sku_Management(p.email_sku_management);
                setEmail_in_case_of_technical_issue(p.email_in_case_of_technical_issue);
                setbuyer_email_to(p.buyer_email_to);
                setbuyer_email_cc(p.buyer_email_cc);
                setsku_creation_template(p.sku_creation_template);
                setParamAppend_FC_Quote_ID_To_Bid_Number(p.Append_FC_Quote_ID_To_Bid_Number);

                if (active.ToUpper() == "TRUE")
                {

                    try
                    {

                        // On utilise EWS API MANAGED REFERENCE 2.2
                        ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);

                        ServicePointManager.ServerCertificateValidationCallback = (Object obj, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors) =>
                        {
                            return true;
                        };

                        WriteToFile(name.ToUpper() + "(" + pays.ToUpper() + ")" + " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                        WriteToFile("   Debug Parameter is set to " + debug.ToUpper() + " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));

                        if (debug.ToUpper() == "TRUE")
                        {
                            WriteToFile("   Connexion to " + sharedmailbox_name + " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                        }
                        service = Connexion_Au_Service_Exchange_O365(sharedmailbox_name);


                        if (debug.ToUpper() == "TRUE")
                        {
                            WriteToFile("   Sender_email_adress      : " + sender_email_adress);
                            WriteToFile("   Sharedmailbox_folder_in  : " + sharedmailbox_folder_in);
                            WriteToFile("   Sharedmailbox_folder_out : " + sharedmailbox_folder_out);
                            WriteToFile("   StartDateScan            : " + start_date_scan);
                            WriteToFile("   Extracting the " + number_of_mails + " most recent messages");
                        }

                        Microsoft.Exchange.WebServices.Data.Folder inbox = Microsoft.Exchange.WebServices.Data.Folder.Bind(service, Find_Folder(service, sharedmailbox_folder_in));
                        ItemView view = new ItemView(int.Parse(number_of_mails));
                        view.PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.DateTimeReceived);
                        view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending);

                        SearchFilter.ContainsSubstring subjectFilter = new SearchFilter.ContainsSubstring(ItemSchema.Subject, "FTQ-", ContainmentMode.Substring, ComparisonMode.IgnoreCase);
                        SearchFilter senderFilter = new SearchFilter.IsEqualTo(EmailMessageSchema.From, sender_email_adress);

                        DateTime searchdate = new DateTime(int.Parse(Strings.Left(start_date_scan, 4)), int.Parse(Strings.Mid(start_date_scan, 5, 2)), int.Parse(Strings.Right(start_date_scan, 2))); //Year, month, day
                        SearchFilter greaterthanfilter = new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeReceived, searchdate);

                        SearchFilter allFilters = new SearchFilter.SearchFilterCollection(LogicalOperator.And, subjectFilter, senderFilter, greaterthanfilter);

                        // This results in a FindItem call to EWS.
                        FindItemsResults<Item> results = inbox.FindItems(allFilters, view);
                        nb_mail = 0;
                        if (results.TotalCount > 0)
                        {
                            if (debug.ToUpper() == "TRUE")
                            {
                                WriteToFile("   " + results.TotalCount + " Email(s) found at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                            }

                            foreach (Item item in results)
                            {
                                try
                                {
                                    PropertySet props = new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.MimeContent);

                                    EmailMessage email = EmailMessage.Bind(service, item.Id, props);

                                    string emlFileName = email_sharing_folder + "\\" + CleanFileName(email.Subject) + ".eml";
                                    emlFileName = CleanPathName(emlFileName);
                                    emlFileName = emlFileName.Replace(";", "");

                                    // Save the email as a EML file
                                    using (var fs = new FileStream(emlFileName, FileMode.Create, FileAccess.Write))
                                    {
                                        fs.Write(email.MimeContent.Content, 0, email.MimeContent.Content.Length);
                                    }

                                    string user_email = email.From.Address;
                                    string body = email.Body.Text;

                                    if (debug.ToUpper() == "TRUE")
                                    {
                                        WriteToFile("       Subject : " + email.Subject + " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                                    }

                                    COP_BID Fortinet_BID = new COP_BID();
                                    Fortinet_BID = traitement_mail_fortinet(email.Body, emlFileName, p.country, sql_con, p.sk_valid);

                                    // 
                                    ///  Check missing skus and send a mail to sku management team
                                    //
                                    if (email_sku_management.Trim() != "" && email_sku_management.Contains("@"))
                                    {
                                        send_missing_or_duplicates_skus_to_sku_management_or_buyer(service, sql_con, Fortinet_BID);
                                    }

                                    string result = "";
                                    if (Append_FC_Quote_ID_To_Bid_Number == "TRUE")
                                    {
                                        // 
                                        ///  Check if multiple FC-QUOTE ID
                                        //

                                        string list_fc_quote_id = "";

                                        foreach (BidLine lignes in Fortinet_BID.bid_lines)
                                        {
                                            if (lignes.IM_COMMENT.ToString().Trim() != "")
                                            {
                                                if (list_fc_quote_id.Contains(lignes.IM_COMMENT.ToString().Trim()) == false)
                                                {
                                                    list_fc_quote_id += lignes.IM_COMMENT.ToString().Trim() + "#";
                                                }
                                            }
                                        }

                                        if (list_fc_quote_id == "")
                                        {
                                            // We create the FORTINET BID
                                            result = "";
                                            result = create_fortinet_BID_in_COP(create_JSON(Fortinet_BID), service, Fortinet_BID.bid_header.BID_NBR_ERP);
                                        }
                                        else
                                        {
                                            string[] fc_quote_id = list_fc_quote_id.Split('#');

                                            foreach (var quote_id in fc_quote_id)
                                            {
                                                if (quote_id != "")
                                                {
                                                    var Fortinet_BID_quote_ID = new COP_BID();
                                                    Fortinet_BID_quote_ID = DeepCopy(Fortinet_BID);

                                                    Fortinet_BID_quote_ID.bid_header.BID_NBR = Fortinet_BID_quote_ID.bid_header.BID_NBR + "/" + quote_id;
                                                    Fortinet_BID_quote_ID.bid_header.BID_NBR_ERP = "";

                                                    foreach (BidLine lignes in Fortinet_BID_quote_ID.bid_lines.ToList())
                                                    {
                                                        if (lignes.IM_COMMENT.ToString().Trim() != quote_id)
                                                        {
                                                            // We remove the lignes
                                                            Fortinet_BID_quote_ID.bid_lines.Remove(lignes);
                                                        }
                                                    }

                                                    // We create the FORTINET BID
                                                    result = "";
                                                    result = create_fortinet_BID_in_COP(create_JSON(Fortinet_BID_quote_ID), service, Fortinet_BID.bid_header.BID_NBR_ERP);
                                                    if (result == "ERROR")
                                                    {
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        // We create the FORTINET BID
                                        result = "";
                                        result = create_fortinet_BID_in_COP(create_JSON(Fortinet_BID), service, Fortinet_BID.bid_header.BID_NBR_ERP);

                                    }

                                    if (result == "ERROR")
                                    {
                                        // Once processed we move the email to Archive folder
                                        Item_move = email.Move(Find_Folder(service, "Erreur"));
                                    }
                                    else
                                    {
                                        // Once processed we move the email to Archive folder
                                        Item_move = email.Move(Find_Folder(service, sharedmailbox_folder_out));
                                    }

                                    Item_move = null;

                                    nb_mail = nb_mail + 1;
                                    System.Threading.Thread.Sleep(500);
                                }
                                catch (Exception)
                                {

                                    // On archive le mail dans le dossier ERREUR
                                    Item_move = item.Move(Find_Folder(service, "Erreur"));
                                    Item_move = null;

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

                        results = null;
                    }
                    catch (Exception e)
                    {
                        if (debug.ToUpper() == "TRUE")
                        {
                            WriteToFile("   Error get emails         : " + e.Message);
                        }
                    }
                }
            }

            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();
        }
        public void Read_Email_with_Graph(string sql_con, string logs, string session_name)
        {
            string global_parameters = "";
            string service_path = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
            int nb_mail = 0;
            setlogs_folder(service_path + "\\" + logs);
            global_session_name = session_name;

            // Get the global parameters of the ACTION
            global_parameters = get_IMCA_paramters(sql_con, "FORTINET_BID_LOAD");

            var param = new JSON_file();
            param = JsonConvert.DeserializeObject<JSON_file>(global_parameters);

            MailFolderCollectionResponse Erreur_folder_id = null;
            MailFolderCollectionResponse sharedmailbox_folder_out_id = null;

            foreach (var p in param.countries) // For each country
            {
                setParamCountry(p.country);
                setParamSK_Valid(p.sk_valid);
                setParamName(p.name);
                setParamActive(p.active);
                setParamDebug(p.debug);
                setStartDateScan(p.start_date_scan);
                setSharedMailboxName(p.sharedmailbox_name);
                setSharedMailboxAccount(p.sharedmailbox_account);
                setSharedMailboxPassword(DecodeFrom64(p.sharedmailbox_password));
                setSharedMailboxServer(p.sharedmailbox_server);
                setsharedmailbox_folder_in(p.sharedmailbox_folder_in);
                setNumber_of_mails(p.number_of_mails);
                setsharedmailbox_folder_out(p.sharedmailbox_folder_out);
                setsender_email_adress(p.sender_email_adress);
                setemail_sharing_folder(p.email_sharing_folder);
                setapi_url(p.api_url);
                setEmail_Sku_Management(p.email_sku_management);
                setEmail_in_case_of_technical_issue(p.email_in_case_of_technical_issue);
                setbuyer_email_to(p.buyer_email_to);
                setbuyer_email_cc(p.buyer_email_cc);
                setsku_creation_template(p.sku_creation_template);
                setParamAppend_FC_Quote_ID_To_Bid_Number(p.Append_FC_Quote_ID_To_Bid_Number);

                if (active.ToUpper() == "TRUE")
                {
                    try
                    {
                        // On utilise EWS API MANAGED REFERENCE 2.2
                        GraphServiceClient GraphService;

                        ServicePointManager.ServerCertificateValidationCallback = (Object obj, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors) =>
                        {
                            return true;
                        };

                        WriteToFile(name.ToUpper() + "(" + pays.ToUpper() + ")" + " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                        WriteToFile("   Debug Parameter is set to " + debug.ToUpper() + " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));

                        if (debug.ToUpper() == "TRUE")
                        {
                            WriteToFile("   Connexion to " + sharedmailbox_name + " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                        }
                        GraphService = Connexion_Microsoft_Graph();

                        if (debug.ToUpper() == "TRUE")
                        {
                            WriteToFile("   Sender_email_adress      : " + sender_email_adress);
                            WriteToFile("   Sharedmailbox_folder_in  : " + sharedmailbox_folder_in);
                            WriteToFile("   Sharedmailbox_folder_out : " + sharedmailbox_folder_out);
                            WriteToFile("   StartDateScan            : " + start_date_scan);
                            WriteToFile("   Extracting the " + number_of_mails + " most recent messages");
                        }

                        DateTime searchdate = new DateTime(int.Parse(Strings.Left(start_date_scan, 4)), int.Parse(Strings.Mid(start_date_scan, 5, 2)), int.Parse(Strings.Right(start_date_scan, 2))); //Year, month, day
                        MessageCollectionResponse messages = null;

                        // Get the Folder Id of sharedmailbox_folder_in
                        Microsoft.Graph.Models.MailFolderCollectionResponse sharedmailbox_folder_id = new MailFolderCollectionResponse();

                        if (sharedmailbox_folder_in.ToUpper().Trim() != "INBOX")
                        {
                            sharedmailbox_folder_id = GraphService.Users[sharedmailbox_name].MailFolders["inbox"].ChildFolders.GetAsync(x =>
                            {
                                x.QueryParameters.Filter = $"displayName eq '{sharedmailbox_folder_in}'";
                            }).GetAwaiter().GetResult();
                        }

                        // Get the Folder Id of sharedmailbox_folder_out

                        sharedmailbox_folder_out_id = GraphService.Users[sharedmailbox_name].MailFolders["inbox"].ChildFolders.GetAsync(x =>
                        {
                            x.QueryParameters.Filter = $"displayName eq '{sharedmailbox_folder_out}'";
                        }).GetAwaiter().GetResult();

                        // Get the Folder Id of sharedmailbox_folder_in

                        Erreur_folder_id = GraphService.Users[sharedmailbox_name].MailFolders["inbox"].ChildFolders.GetAsync(x =>
                        {
                            x.QueryParameters.Filter = $"displayName eq 'Erreur'";
                        }).GetAwaiter().GetResult();

                        // Get the emails from Folder Id
                        if (sharedmailbox_folder_in.ToUpper().Trim() != "INBOX")
                        {
                            messages = GraphService.Users[sharedmailbox_name].MailFolders[sharedmailbox_folder_id.Value.FirstOrDefault().Id].Messages.GetAsync((config) =>
                            {
                                config.QueryParameters.Top = int.Parse(number_of_mails);
                                config.QueryParameters.Filter = $"receivedDateTime gt {searchdate.ToString("yyyy-MM-dd")} and (from/emailAddress/address) eq '" + sender_email_adress + "' and contains(subject,'FTQ-')";
                                config.QueryParameters.Orderby = new string[] { "receivedDateTime desc" };
                            }).GetAwaiter().GetResult();
                        }
                        else
                        {
                            messages = GraphService.Users[sharedmailbox_name].MailFolders["inbox"].Messages.GetAsync((config) =>
                                {
                                    config.QueryParameters.Top = int.Parse(number_of_mails);
                                    config.QueryParameters.Filter = $"receivedDateTime gt {searchdate.ToString("yyyy-MM-dd")} and (from/emailAddress/address) eq '" + sender_email_adress + "' and contains(subject,'FTQ-')";
                                    config.QueryParameters.Orderby = new string[] { "receivedDateTime desc" };
                                }).GetAwaiter().GetResult();
                        }

                        nb_mail = 0;

                        if (messages.Value.Count > 0)
                        {
                            if (debug.ToUpper() == "TRUE")
                            {
                                WriteToFile("   " + messages.Value.Count.ToString() + " Email(s) found at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                            }

                            foreach (var email in messages.Value)
                            {
                                try
                                {
                                    var mimeContentStream = GraphService.Users[sharedmailbox_name].Messages[email.Id].Content.GetAsync().GetAwaiter().GetResult();


                                    string emlFileName = email_sharing_folder + "\\" + CleanFileName(email.Subject) + ".eml";
                                    emlFileName = CleanPathName(emlFileName);
                                    emlFileName = emlFileName.Replace(";", "");

                                    // Save the email as a EML file
                                    using (var fs = new FileStream(emlFileName, FileMode.Create, FileAccess.Write))
                                    {
                                        mimeContentStream.CopyTo(fs);
                                    }

                                    string user_email = email.From.EmailAddress.ToString();
                                    string body = email.Body.ToString();

                                    if (debug.ToUpper() == "TRUE")
                                    {
                                        WriteToFile("       Subject : " + email.Subject + " at " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                                    }

                                    COP_BID Fortinet_BID = new COP_BID();
                                    Fortinet_BID = traitement_mail_fortinet(email.Body.Content.ToString(), emlFileName, p.country, sql_con, p.sk_valid);

                                    // 
                                    ///  Check missing skus and send a mail to sku management team
                                    //
                                    if (email_sku_management.Trim() != "" && email_sku_management.Contains("@"))
                                    {
                                        send_missing_or_duplicates_skus_to_sku_management_or_buyer_with_Graph(GraphService, sql_con, Fortinet_BID);
                                    }

                                    string result = "";
                                    if (Append_FC_Quote_ID_To_Bid_Number == "TRUE")
                                    {
                                        // 
                                        ///  Check if multiple FC-QUOTE ID
                                        //

                                        string list_fc_quote_id = "";

                                        foreach (BidLine lignes in Fortinet_BID.bid_lines)
                                        {
                                            if (lignes.IM_COMMENT.ToString().Trim() != "")
                                            {
                                                if (list_fc_quote_id.Contains(lignes.IM_COMMENT.ToString().Trim()) == false)
                                                {
                                                    list_fc_quote_id += lignes.IM_COMMENT.ToString().Trim() + "#";
                                                }
                                            }
                                        }

                                        if (list_fc_quote_id == "")
                                        {
                                            // We create the FORTINET BID
                                            result = "";
                                            result = create_fortinet_BID_in_COP_with_Graph(create_JSON(Fortinet_BID), GraphService, Fortinet_BID.bid_header.BID_NBR_ERP);
                                        }
                                        else
                                        {
                                            string[] fc_quote_id = list_fc_quote_id.Split('#');

                                            foreach (var quote_id in fc_quote_id)
                                            {
                                                if (quote_id != "")
                                                {
                                                    var Fortinet_BID_quote_ID = new COP_BID();
                                                    Fortinet_BID_quote_ID = DeepCopy(Fortinet_BID);

                                                    Fortinet_BID_quote_ID.bid_header.BID_NBR = Fortinet_BID_quote_ID.bid_header.BID_NBR + "/" + quote_id;
                                                    Fortinet_BID_quote_ID.bid_header.BID_NBR_ERP = "";

                                                    foreach (BidLine lignes in Fortinet_BID_quote_ID.bid_lines.ToList())
                                                    {
                                                        if (lignes.IM_COMMENT.ToString().Trim() != quote_id)
                                                        {
                                                            // We remove the lignes
                                                            Fortinet_BID_quote_ID.bid_lines.Remove(lignes);
                                                        }
                                                    }

                                                    // We create the FORTINET BID
                                                    result = "";
                                                    result = create_fortinet_BID_in_COP_with_Graph(create_JSON(Fortinet_BID_quote_ID), GraphService, Fortinet_BID.bid_header.BID_NBR_ERP);
                                                    if (result == "ERROR")
                                                    {
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        // We create the FORTINET BID
                                        result = "";
                                        result = create_fortinet_BID_in_COP_with_Graph(create_JSON(Fortinet_BID), GraphService, Fortinet_BID.bid_header.BID_NBR_ERP);

                                    }

                                    if (result == "ERROR")
                                    {
                                        // Once processed we move the email to Error folder
                                        var requestBody = new Microsoft.Graph.Users.Item.Messages.Item.Move.MovePostRequestBody
                                        {
                                            DestinationId = Erreur_folder_id.Value[0].Id
                                        };
                                        GraphService.Users[sharedmailbox_name].Messages[email.Id].Move.PostAsync(requestBody).GetAwaiter().GetResult(); ;
                                    }
                                    else
                                    {
                                        // Once processed we move the email to Archive folder
                                        var requestBody = new Microsoft.Graph.Users.Item.Messages.Item.Move.MovePostRequestBody
                                        {
                                            DestinationId = sharedmailbox_folder_out_id.Value[0].Id
                                        };
                                        GraphService.Users[sharedmailbox_name].Messages[email.Id].Move.PostAsync(requestBody).GetAwaiter().GetResult();
                                    }

                                    nb_mail = nb_mail + 1;
                                    System.Threading.Thread.Sleep(500);
                                }
                                catch (Exception ex)
                                {
                                    // Once processed we move the email to Error folder
                                    var requestBody = new Microsoft.Graph.Users.Item.Messages.Item.Move.MovePostRequestBody
                                    {
                                        DestinationId = Erreur_folder_id.Value[0].Id
                                    };
                                    GraphService.Users[sharedmailbox_name].Messages[email.Id].Move.PostAsync(requestBody).GetAwaiter().GetResult();

                                    WriteToFile("   Error : " + ex.Message + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
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
                        if (debug.ToUpper() == "TRUE")
                        {
                            WriteToFile("   Error get emails         : " + e.Message);
                        }
                    }
                }
            }

            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();
        }
    }
}
