using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Runtime.InteropServices;
using System.Text;

namespace dailyreportauto
{
    class Program
    {
        static void Main(string[] args)
        {

            function fn = new function();
        //variable
        string settlement_date = "";
            string status = "";
            string BOI = "";
            string CITRIX = "";
            string AUTOREPORT = "";
            string exc_itg = "";
            string exc_mac = "";
            string exc_ddt = "";
            string exc_exp = "";
            string exc_ind = "";
            string exc_ivd = "";
            string exc_unc = "";
            string exc_irv = "";
            string exc_xtx = "";
            string exc_cmr = "";
            string exc_pmr = "";
            string exc_psn = "";
            string bts_num_received_file = "";
            string bts_num_failed_file = "";
            string brt_num_received_file = "";
            string brt_num_failed_file = "";
            string bbl_num_received_file = "";
            string bbl_num_failed_file = "";
            string crr_num_received_file = "";
            string crr_num_failed_file = "";
            string maq_num_received_file = "";
            string maq_num_failed_file = "";
            string mpay_num_received_file = "";
            string mpay_num_failed_file = "";
            string bigc_num_received_file = "";
            string bigc_num_failed_file = "";
            string prg_num_received_file = "";
            string prg_num_failed_file = "";
            string eqf_num_received_file = "";
            string eqf_num_failed_file = "";
            string epf_num_received_file = "";
            string epf_num_failed_file = "";
            string mbf_num_received_file = "";
            string mbf_num_failed_file = "";
            string FSF_NUM_RECEIVED_FILE = "";
            string FSF_NUM_FAILED_FILE = "";
            string mtf_num_received_file = "";
            string mtf_num_failed_file = "";
            string bcc_num_received_file = "";
            string bcc_num_failed_file = "";
            string bkf_num_received_file = "";
            string bkf_num_failed_file = "";
            string fef_num_received_file = "";
            string fef_num_failed_file = "";
            string sef_num_received_file = "";
            string sef_num_failed_file = "";
            string ips_num_received_file = "";
            string ips_num_failed_file = "";
            string mwf_num_failed_file = "";
            string mwf_num_received_file = "";
            string mwf_num_junk_file = "";
            string cwf_num_failed_file = "";
            string cwf_num_received_file = "";
            string cwf_num_junk_file = "";

            ///select from file
            DataTable ds = new DataTable();

            ds = fn.connect_file();
            foreach (DataRow drCurrent in ds.Rows)
            {
                settlement_date = Convert.ToDateTime(drCurrent["settlement_date"].ToString()).ToString(CultureInfo.CreateSpecificCulture("en-US"));
                if (drCurrent["status"].ToString() == "0")
                {
                    status = "OK";
                }
                else
                {
                    status = "NotOK";
                }
                if (drCurrent["BOI"].ToString() == "0")
                {
                    BOI = "OK";
                }
                else
                {
                    BOI = "NotOK";
                }

                if (drCurrent["CITRIX"].ToString() == "0")
                {
                    CITRIX = "OK";
                }
                else
                {
                    CITRIX = "NotOK";
                }

                if (drCurrent["AUTOREPORT"].ToString() == "0")
                {
                    AUTOREPORT = "OK";
                }
                else
                {
                    AUTOREPORT = "NotOK";
                }
                if (drCurrent["exc_itg"].ToString().Length > 3) { exc_itg = Convert.ToDecimal(drCurrent["exc_itg"].ToString()).ToString("0,000"); } else { exc_itg = drCurrent["exc_itg"].ToString(); }
                if (drCurrent["exc_mac"].ToString().Length > 3) { exc_mac = Convert.ToDecimal(drCurrent["exc_mac"].ToString()).ToString("0,000"); } else { exc_mac = drCurrent["exc_mac"].ToString(); }
                if (drCurrent["exc_ddt"].ToString().Length > 3) { exc_ddt = Convert.ToDecimal(drCurrent["exc_ddt"].ToString()).ToString("0,000"); } else { exc_ddt = drCurrent["exc_ddt"].ToString(); }
                if (drCurrent["exc_exp"].ToString().Length > 3) { exc_exp = Convert.ToDecimal(drCurrent["exc_exp"].ToString()).ToString("0,000"); } else { exc_exp = drCurrent["exc_exp"].ToString(); }
                if (drCurrent["exc_ind"].ToString().Length > 3) { exc_ind = Convert.ToDecimal(drCurrent["exc_ind"].ToString()).ToString("0,000"); } else { exc_ind = drCurrent["exc_ind"].ToString(); }
                if (drCurrent["exc_ivd"].ToString().Length > 3) { exc_ivd = Convert.ToDecimal(drCurrent["exc_ivd"].ToString()).ToString("0,000"); } else { exc_ivd = drCurrent["exc_ivd"].ToString(); }
                if (drCurrent["exc_unc"].ToString().Length > 3) { exc_unc = Convert.ToDecimal(drCurrent["exc_unc"].ToString()).ToString("0,000"); } else { exc_unc = drCurrent["exc_unc"].ToString(); }
                if (drCurrent["exc_irv"].ToString().Length > 3) { exc_irv = Convert.ToDecimal(drCurrent["exc_irv"].ToString()).ToString("0,000"); } else { exc_irv = drCurrent["exc_irv"].ToString(); }
                if (drCurrent["exc_xtx"].ToString().Length > 3) { exc_xtx = Convert.ToDecimal(drCurrent["exc_xtx"].ToString()).ToString("0,000"); } else { exc_xtx = drCurrent["exc_xtx"].ToString(); }
                if (drCurrent["exc_cmr"].ToString().Length > 3) { exc_cmr = Convert.ToDecimal(drCurrent["exc_cmr"].ToString()).ToString("0,000"); } else { exc_cmr = drCurrent["exc_cmr"].ToString(); }
                if (drCurrent["exc_pmr"].ToString().Length > 3) { exc_pmr = Convert.ToDecimal(drCurrent["exc_pmr"].ToString()).ToString("0,000"); } else { exc_pmr = drCurrent["exc_pmr"].ToString(); }
                if (drCurrent["exc_psn"].ToString().Length > 3) { exc_psn = Convert.ToDecimal(drCurrent["exc_psn"].ToString()).ToString("0,000"); } else { exc_psn = drCurrent["exc_psn"].ToString(); }
                if (drCurrent["bts_num_received_file"].ToString().Length > 3) { bts_num_received_file = Convert.ToDecimal(drCurrent["bts_num_received_file"].ToString()).ToString("0,000"); } else { bts_num_received_file = drCurrent["bts_num_received_file"].ToString(); }
                if (drCurrent["bts_num_failed_file"].ToString().Length > 3) { bts_num_failed_file = Convert.ToDecimal(drCurrent["bts_num_failed_file"].ToString()).ToString("0,000"); } else { bts_num_failed_file = drCurrent["bts_num_failed_file"].ToString(); }
                if (drCurrent["brt_num_received_file"].ToString().Length > 3) { brt_num_received_file = Convert.ToDecimal(drCurrent["brt_num_received_file"].ToString()).ToString("0,000"); } else { brt_num_received_file = drCurrent["brt_num_received_file"].ToString(); }
                if (drCurrent["brt_num_failed_file"].ToString().Length > 3) { brt_num_failed_file = Convert.ToDecimal(drCurrent["brt_num_failed_file"].ToString()).ToString("0,000"); } else { brt_num_failed_file = drCurrent["brt_num_failed_file"].ToString(); }
                if (drCurrent["bbl_num_received_file"].ToString().Length > 3) { bbl_num_received_file = Convert.ToDecimal(drCurrent["bbl_num_received_file"].ToString()).ToString("0,000"); } else { bbl_num_received_file = drCurrent["bbl_num_received_file"].ToString(); }
                if (drCurrent["bbl_num_failed_file"].ToString().Length > 3) { bbl_num_failed_file = Convert.ToDecimal(drCurrent["bbl_num_failed_file"].ToString()).ToString("0,000"); } else { bbl_num_failed_file = drCurrent["bbl_num_failed_file"].ToString(); }
                if (drCurrent["crr_num_received_file"].ToString().Length > 3) { crr_num_received_file = Convert.ToDecimal(drCurrent["crr_num_received_file"].ToString()).ToString("0,000"); } else { crr_num_received_file = drCurrent["crr_num_received_file"].ToString(); }
                if (drCurrent["crr_num_failed_file"].ToString().Length > 3) { crr_num_failed_file = Convert.ToDecimal(drCurrent["crr_num_failed_file"].ToString()).ToString("0,000"); } else { crr_num_failed_file = drCurrent["crr_num_failed_file"].ToString(); }
                if (drCurrent["maq_num_received_file"].ToString().Length > 3) { maq_num_received_file = Convert.ToDecimal(drCurrent["maq_num_received_file"].ToString()).ToString("0,000"); } else { maq_num_received_file = drCurrent["maq_num_received_file"].ToString(); }
                if (drCurrent["maq_num_failed_file"].ToString().Length > 3) { maq_num_failed_file = Convert.ToDecimal(drCurrent["maq_num_failed_file"].ToString()).ToString("0,000"); } else { maq_num_failed_file = drCurrent["maq_num_failed_file"].ToString(); }
                if (drCurrent["mpay_num_received_file"].ToString().Length > 3) { mpay_num_received_file = Convert.ToDecimal(drCurrent["mpay_num_received_file"].ToString()).ToString("0,000"); } else { mpay_num_received_file = drCurrent["mpay_num_received_file"].ToString(); }
                if (drCurrent["mpay_num_failed_file"].ToString().Length > 3) { mpay_num_failed_file = Convert.ToDecimal(drCurrent["mpay_num_failed_file"].ToString()).ToString("0,000"); } else { mpay_num_failed_file = drCurrent["mpay_num_failed_file"].ToString(); }
                if (drCurrent["bigc_num_received_file"].ToString().Length > 3) { bigc_num_received_file = Convert.ToDecimal(drCurrent["bigc_num_received_file"].ToString()).ToString("0,000"); } else { bigc_num_received_file = drCurrent["bigc_num_received_file"].ToString(); }
                if (drCurrent["bigc_num_failed_file"].ToString().Length > 3) { bigc_num_failed_file = Convert.ToDecimal(drCurrent["bigc_num_failed_file"].ToString()).ToString("0,000"); } else { bigc_num_failed_file = drCurrent["bigc_num_failed_file"].ToString(); }
                if (drCurrent["prg_num_received_file"].ToString().Length > 3) { prg_num_received_file = Convert.ToDecimal(drCurrent["prg_num_received_file"].ToString()).ToString("0,000"); } else { prg_num_received_file = drCurrent["prg_num_received_file"].ToString(); }
                if (drCurrent["prg_num_failed_file"].ToString().Length > 3) { prg_num_failed_file = Convert.ToDecimal(drCurrent["prg_num_failed_file"].ToString()).ToString("0,000"); } else { prg_num_failed_file = drCurrent["prg_num_failed_file"].ToString(); }
                if (drCurrent["eqf_num_received_file"].ToString().Length > 3) { eqf_num_received_file = Convert.ToDecimal(drCurrent["eqf_num_received_file"].ToString()).ToString("0,000"); } else { eqf_num_received_file = drCurrent["eqf_num_received_file"].ToString(); }
                if (drCurrent["eqf_num_failed_file"].ToString().Length > 3) { eqf_num_failed_file = Convert.ToDecimal(drCurrent["eqf_num_failed_file"].ToString()).ToString("0,000"); } else { eqf_num_failed_file = drCurrent["eqf_num_failed_file"].ToString(); }
                if (drCurrent["epf_num_received_file"].ToString().Length > 3) { epf_num_received_file = Convert.ToDecimal(drCurrent["epf_num_received_file"].ToString()).ToString("0,000"); } else { epf_num_received_file = drCurrent["epf_num_received_file"].ToString(); }
                if (drCurrent["epf_num_failed_file"].ToString().Length > 3) { epf_num_failed_file = Convert.ToDecimal(drCurrent["epf_num_failed_file"].ToString()).ToString("0,000"); } else { epf_num_failed_file = drCurrent["epf_num_failed_file"].ToString(); }
                if (drCurrent["mbf_num_received_file"].ToString().Length > 3) { mbf_num_received_file = Convert.ToDecimal(drCurrent["mbf_num_received_file"].ToString()).ToString("0,000"); } else { mbf_num_received_file = drCurrent["mbf_num_received_file"].ToString(); }
                if (drCurrent["mbf_num_failed_file"].ToString().Length > 3) { mbf_num_failed_file = Convert.ToDecimal(drCurrent["mbf_num_failed_file"].ToString()).ToString("0,000"); } else { mbf_num_failed_file = drCurrent["mbf_num_failed_file"].ToString(); }
                if (drCurrent["FSF_NUM_RECEIVED_FILE"].ToString().Length > 3) { FSF_NUM_RECEIVED_FILE = Convert.ToDecimal(drCurrent["FSF_NUM_RECEIVED_FILE"].ToString()).ToString("0,000"); } else { FSF_NUM_RECEIVED_FILE = drCurrent["FSF_NUM_RECEIVED_FILE"].ToString(); }
                if (drCurrent["FSF_NUM_FAILED_FILE"].ToString().Length > 3) { FSF_NUM_FAILED_FILE = Convert.ToDecimal(drCurrent["FSF_NUM_FAILED_FILE"].ToString()).ToString("0,000"); } else { FSF_NUM_FAILED_FILE = drCurrent["FSF_NUM_FAILED_FILE"].ToString(); }
                if (drCurrent["mtf_num_received_file"].ToString().Length > 3) { mtf_num_received_file = Convert.ToDecimal(drCurrent["mtf_num_received_file"].ToString()).ToString("0,000"); } else { mtf_num_received_file = drCurrent["mtf_num_received_file"].ToString(); }
                if (drCurrent["mtf_num_failed_file"].ToString().Length > 3) { mtf_num_failed_file = Convert.ToDecimal(drCurrent["mtf_num_failed_file"].ToString()).ToString("0,000"); } else { mtf_num_failed_file = drCurrent["mtf_num_failed_file"].ToString(); }
                if (drCurrent["cirf_num_received_file"].ToString().Length > 3) { bcc_num_received_file = Convert.ToDecimal(drCurrent["cirf_num_received_file"].ToString()).ToString("0,000"); } else { bcc_num_received_file = drCurrent["cirf_num_received_file"].ToString(); }
                if (drCurrent["cirf_num_failed_file"].ToString().Length > 3) { bcc_num_failed_file = Convert.ToDecimal(drCurrent["cirf_num_failed_file"].ToString()).ToString("0,000"); } else { bcc_num_failed_file = drCurrent["cirf_num_failed_file"].ToString(); }
                if (drCurrent["bkf_num_received_file"].ToString().Length > 3) { bkf_num_received_file = Convert.ToDecimal(drCurrent["bkf_num_received_file"].ToString()).ToString("0,000"); } else { bkf_num_received_file = drCurrent["bkf_num_received_file"].ToString(); }
                if (drCurrent["bkf_num_failed_file"].ToString().Length > 3) { bkf_num_failed_file = Convert.ToDecimal(drCurrent["bkf_num_failed_file"].ToString()).ToString("0,000"); } else { bkf_num_failed_file = drCurrent["bkf_num_failed_file"].ToString(); }
                if (drCurrent["fef_num_received_file"].ToString().Length > 3) { fef_num_received_file = Convert.ToDecimal(drCurrent["fef_num_received_file"].ToString()).ToString("0,000"); } else { fef_num_received_file = drCurrent["fef_num_received_file"].ToString(); }
                if (drCurrent["fef_num_failed_file"].ToString().Length > 3) { fef_num_failed_file = Convert.ToDecimal(drCurrent["fef_num_failed_file"].ToString()).ToString("0,000"); } else { fef_num_failed_file = drCurrent["fef_num_failed_file"].ToString(); }
                if (drCurrent["sef_num_received_file"].ToString().Length > 3) { sef_num_received_file = Convert.ToDecimal(drCurrent["sef_num_received_file"].ToString()).ToString("0,000"); } else { sef_num_received_file = drCurrent["sef_num_received_file"].ToString(); }
                if (drCurrent["sef_num_failed_file"].ToString().Length > 3) { sef_num_failed_file = Convert.ToDecimal(drCurrent["sef_num_failed_file"].ToString()).ToString("0,000"); } else { sef_num_failed_file = drCurrent["sef_num_failed_file"].ToString(); }
                if (drCurrent["ips_num_received_file"].ToString().Length > 3) { ips_num_received_file = Convert.ToDecimal(drCurrent["ips_num_received_file"].ToString()).ToString("0,000"); } else { ips_num_received_file = drCurrent["ips_num_received_file"].ToString(); }
                if (drCurrent["ips_num_failed_file"].ToString().Length > 3) { ips_num_failed_file = Convert.ToDecimal(drCurrent["ips_num_failed_file"].ToString()).ToString("0,000"); } else { ips_num_failed_file = drCurrent["ips_num_failed_file"].ToString(); }
                if (drCurrent["mwf_num_received_file"].ToString().Length > 3) { mwf_num_received_file = Convert.ToDecimal(drCurrent["mwf_num_received_file"].ToString()).ToString("0,000"); } else { mwf_num_received_file = drCurrent["mwf_num_received_file"].ToString(); }
                if (drCurrent["mwf_num_failed_file"].ToString().Length > 3) { mwf_num_failed_file = Convert.ToDecimal(drCurrent["mwf_num_failed_file"].ToString()).ToString("0,000"); } else { mwf_num_failed_file = drCurrent["mwf_num_failed_file"].ToString(); }
                if (drCurrent["mwf_num_junk_file"].ToString().Length > 3) { mwf_num_junk_file = Convert.ToDecimal(drCurrent["mwf_num_junk_file"].ToString()).ToString("0,000"); } else { mwf_num_junk_file = drCurrent["mwf_num_junk_file"].ToString(); }
                if (drCurrent["cwf_num_received_file"].ToString().Length > 3) { cwf_num_received_file = Convert.ToDecimal(drCurrent["cwf_num_received_file"].ToString()).ToString("0,000"); } else { cwf_num_received_file = drCurrent["cwf_num_received_file"].ToString(); }
                if (drCurrent["cwf_num_failed_file"].ToString().Length > 3) { cwf_num_failed_file = Convert.ToDecimal(drCurrent["cwf_num_failed_file"].ToString()).ToString("0,000"); } else { cwf_num_failed_file = drCurrent["cwf_num_failed_file"].ToString(); }
                if (drCurrent["cwf_num_junk_file"].ToString().Length > 3) { cwf_num_junk_file = Convert.ToDecimal(drCurrent["cwf_num_junk_file"].ToString()).ToString("0,000"); } else { cwf_num_junk_file = drCurrent["cwf_num_junk_file"].ToString(); }

            }

            //select from purse
            string[] RPT_GRP = new string[100];
            string[] TXN_TYPE = new string[100];
            string[] TXN_VOL = new string[100];
            string[] TXN_VALUE = new string[100];
            string BBL_Purse_Add_Val = "0";
            string BBL_Purse_Add_Vol = "0";
            string BBL_Purse_Use_Val = "0";
            string BBL_Purse_Use_Vol = "0";
            string BBL_Purse_Use_Reverse = "0";
            string BCC_Card_Issue = "0";
            string BCC_Purse_Add_Vol = "0";
            string BCC_Purse_Add_Val = "0";
            string BCC_Card_Refund = "0";
            string BCC_Card_Refund_Deferred = "0";
            string BCC_Card_Surrender = "0";
            string BCC_Purse_Issue_Vol = "0";
            string BCC_Purse_Issue_Val = "0";
            string BCC_Purse_Refund_Val = "0";
            string BCC_Purse_Refund_Vol = "0";
            string BRT_Card_Issue = "0";
            string BRT_Card_Use_Failed = "0";
            string BRT_Purse_Add_Val = "0";
            string BRT_Purse_Add_Vol = "0";
            string BRT_Purse_Compensation_Fare_Val = "0";
            string BRT_Purse_Compensation_Fare_Vol = "0";
            string BRT_Purse_Issue_Vol = "0";
            string BRT_Purse_Issue_Val = "0";
            string BRT_Purse_Use_On_Entry_Vol = "0";
            string BRT_Purse_Use_On_Entry_Val = "0";
            string BRT_Purse_Use_On_Exit_Vol = "0";
            string BRT_Purse_Use_On_Exit_Val = "0";
            string BRT_Card_Block = "0";
            string BRT_Card_Refund = "0";
            string BTS_Card_Issue = "0";
            string BTS_Card_Refund = "0";
            string BTS_Card_Use_Failed = "0";
            string BTS_Multiride_Issue_Val = "0";
            string BTS_Multiride_Issue_Vol = "0";
            string BTS_Multiride_Issue_Reverse = "0";
            string BTS_Multiride_Use_On_Entry_Val = "0";
            string BTS_Multiride_Use_On_Entry_Vol = "0";
            string BTS_Multiride_Use_On_Exit_Val = "0";
            string BTS_Multiride_Use_On_Exit_Vol = "0";
            string BTS_Purse_Add_Vol = "0";
            string BTS_Purse_Add_Val = "0";
            string BTS_Purse_Add_Reverse = "0";
            string BTS_Purse_Compensation_Fare_Val = "0";
            string BTS_Purse_Compensation_Fare_Vol = "0";
            string BTS_Purse_Issue_Val = "0";
            string BTS_Purse_Issue_Vol = "0";
            string BTS_Purse_Refund = "0";
            string BTS_Purse_Use = "0";
            string BTS_Purse_Use_On_Entry_Val = "0";
            string BTS_Purse_Use_On_Entry_Vol = "0";
            string BTS_Purse_Use_On_Exit_Val = "0";
            string BTS_Purse_Use_On_Exit_Vol = "0";
            string BTS_Purse_Use_Reverse = "0";
            string BTS_Card_Block = "0";
            string CARR_Purse_Add_Vol = "0";
            string CARR_Purse_Add_Val = "0";
            string CARR_Card_Use_Failed = "0";
            string CARR_Card_Block = "0";
            string MCDIRECT_Card_Block = "0";
            string MCDIRECT_Purse_Add_Vol = "0";
            string MCDIRECT_Purse_Add_Val = "0";
            string MCDIRECT_Purse_Use_Vol = "0";
            string MCDIRECT_Purse_Use_Val = "0";
            string MCDIRECT_Purse_Use_Reverse = "0";
            string MPAY_Card_Personalise = "0";
            string MCDIRECT_Purse_Add_Reverse = "0";
            string MPAY_Purse_Add_Reverse = "0";
            string MPAY_Purse_Add_Vol = "0";
            string MPAY_Purse_Add_Val = "0";
            string BIGC_Purse_Add_Reverse = "0";
            string BIGC_Purse_Use_Reverse = "0";
            string BIGC_Purse_Add_Vol = "0";
            string BIGC_Purse_Add_Val = "0";
            string BIGC_Purse_Use_Vol = "0";
            string BIGC_Purse_Use_Val = "0";
            string FOOD_Purse_Add_Reverse_Vol = "0";
            string FOOD_Purse_Add_Reverse_Val = "0";
            string FOOD_Purse_Use_Reverse_Vol = "0";
            string FOOD_Purse_Use_Reverse_Val = "0";
            string FOOD_Purse_Add_Vol = "0";
            string FOOD_Purse_Add_Val = "0";
            string FOOD_Purse_Use_Vol = "0";
            string FOOD_Purse_Use_Val = "0";
            string ATU_Product_Autoload_Enable = "0";
            string ATU_Product_Autoload_Disable = "0";
            string ATU_Change_of_Service = "0";
            string TESCO_Purse_Use_Reverse_Vol = "0";
            string TESCO_Purse_Use_Reverse_Val = "0";
            string TESCO_Purse_Add_Vol = "0";
            string TESCO_Purse_Add_Val = "0";
            string TESCO_Purse_Use_Vol = "0";
            string TESCO_Purse_Use_Val = "0";
            string EOF_Purse_Add_Reverse_Vol = "0";
            string EOF_Purse_Add_Reverse_Val = "0";
            string EOF_Purse_Use_Reverse_Vol = "0";
            string EOF_Purse_Use_Reverse_Val = "0";
            string EOF_Purse_Add_Vol = "0";
            string EOF_Purse_Add_Val = "0";
            string EOF_Purse_Use_Vol = "0";
            string EOF_Purse_Use_Val = "0";
            string EMPORIUM_Purse_Add_Reverse_Vol = "0";
            string EMPORIUM_Purse_Add_Reverse_Val = "0";
            string EMPORIUM_Purse_Use_Reverse_Vol = "0";
            string EMPORIUM_Purse_Use_Reverse_Val = "0";
            string EMPORIUM_Purse_Add_Vol = "0";
            string EMPORIUM_Purse_Add_Val = "0";
            string EMPORIUM_Purse_Use_Vol = "0";
            string EMPORIUM_Purse_Use_Val = "0";
            string MBK_Purse_Add_Reverse_Vol = "0";
            string MBK_Purse_Add_Reverse_Val = "0";
            string MBK_Purse_Use_Reverse_Vol = "0";
            string MBK_Purse_Use_Reverse_Val = "0";
            string MBK_Purse_Add_Vol = "0";
            string MBK_Purse_Add_Val = "0";
            string MBK_Purse_Use_Vol = "0";
            string MBK_Purse_Use_Val = "0";
            string FOOD_STREET_Purse_Add_Reverse_Vol = "0";
            string FOOD_STREET_Purse_Add_Reverse_Val = "0";
            string FOOD_STREET_Purse_Use_Reverse_Vol = "0";
            string FOOD_STREET_Purse_Use_Reverse_Val = "0";
            string FOOD_STREET_Purse_Add_Vol = "0";
            string FOOD_STREET_Purse_Add_Val = "0";
            string FOOD_STREET_Purse_Use_Vol = "0";
            string FOOD_STREET_Purse_Use_Val = "0";
            string THAPRA_Purse_Add_Reverse_Vol = "0";
            string THAPRA_Purse_Add_Reverse_Val = "0";
            string THAPRA_Purse_Use_Reverse_Vol = "0";
            string THAPRA_Purse_Use_Reverse_Val = "0";
            string THAPRA_Purse_Add_Vol = "0";
            string THAPRA_Purse_Add_Val = "0";
            string THAPRA_Purse_Use_Vol = "0";
            string THAPRA_Purse_Use_Val = "0";
            string LEGACY_Purse_Add_Reverse_Val = "0";
            string LEGACY_Purse_Add_Reverse_Vol = "0";
            string LEGACY_Purse_Add_Vol = "0";
            string LEGACY_Purse_Add_Val = "0";
            string BCC_Card_Initialise = "0";
            string BANGKAE_Purse_Add_Vol = "0";
            string BANGKAE_Purse_Add_Val = "0";
            string BANGKAE_Purse_Use_Vol = "0";
            string BANGKAE_Purse_Use_Val = "0";
            string BANGKAE_Purse_Add_Reverse_Vol = "0";
            string BANGKAE_Purse_Add_Reverse_Val = "0";
            string BANGKAE_Purse_Use_Reverse_Vol = "0";
            string BANGKAE_Purse_Use_Reverse_Val = "0";
            string FSEKAMAI_Purse_Add_Vol = "0";
            string FSEKAMAI_Purse_Add_Val = "0";
            string FSEKAMAI_Purse_Use_Vol = "0";
            string FSEKAMAI_Purse_Use_Val = "0";
            string FSEKAMAI_Purse_Add_Reverse_Vol = "0";
            string FSEKAMAI_Purse_Add_Reverse_Val = "0";
            string FSEKAMAI_Purse_Use_Reverse_Vol = "0";
            string FSEKAMAI_Purse_Use_Reverse_Val = "0";
            string IMPERIAL4F_Purse_Add_Vol = "0";
            string IMPERIAL4F_Purse_Add_Val = "0";
            string IMPERIAL4F_Purse_Use_Vol = "0";
            string IMPERIAL4F_Purse_Use_Val = "0";
            string IMPERIAL4F_Purse_Add_Reverse_Vol = "0";
            string IMPERIAL4F_Purse_Add_Reverse_Val = "0";
            string IMPERIAL4F_Purse_Use_Reverse_Vol = "0";
            string IMPERIAL4F_Purse_Use_Reverse_Val = "0";
            string IMPERIALBF_Purse_Add_Vol = "0";
            string IMPERIALBF_Purse_Add_Val = "0";
            string IMPERIALBF_Purse_Use_Vol = "0";
            string IMPERIALBF_Purse_Use_Val = "0";
            string IMPERIALBF_Purse_Add_Reverse_Vol = "0";
            string IMPERIALBF_Purse_Add_Reverse_Val = "0";
            string IMPERIALBF_Purse_Use_Reverse_Vol = "0";
            string IMPERIALBF_Purse_Use_Reverse_Val = "0";
            string SET_Purse_Add_Vol = "0";
            string SET_Purse_Add_Val = "0";
            string SET_Purse_Add_Reverse_Vol = "0";
            string SET_Purse_Add_Reverse_Val = "0";
            string SETLOC1_Purse_Use_Vol = "0";
            string SETLOC1_Purse_Use_Val = "0";
            string SETLOC1_Purse_Use_Reverse_Vol = "0";
            string SETLOC1_Purse_Use_Reverse_Val = "0";
            string SETLOC2_Purse_Use_Vol = "0";
            string SETLOC2_Purse_Use_Val = "0";
            string SETLOC2_Purse_Use_Reverse_Vol = "0";
            string SETLOC2_Purse_Use_Reverse_Val = "0";
            string SETLOC3_Purse_Use_Vol = "0";
            string SETLOC3_Purse_Use_Val = "0";
            string SETLOC3_Purse_Use_Reverse_Vol = "0";
            string SETLOC3_Purse_Use_Reverse_Val = "0";
            string SETLOC4_Purse_Use_Vol = "0";
            string SETLOC4_Purse_Use_Val = "0";
            string SETLOC4_Purse_Use_Reverse_Vol = "0";
            string SETLOC4_Purse_Use_Reverse_Val = "0";
            string SETLOC5_Purse_Use_Vol = "0";
            string SETLOC5_Purse_Use_Val = "0";
            string SETLOC5_Purse_Use_Reverse_Vol = "0";
            string SETLOC5_Purse_Use_Reverse_Val = "0";
            string SETLOC6_Purse_Use_Vol = "0";
            string SETLOC6_Purse_Use_Val = "0";
            string SETLOC6_Purse_Use_Reverse_Vol = "0";
            string SETLOC6_Purse_Use_Reverse_Val = "0";
            string SETLOC7_Purse_Use_Vol = "0";
            string SETLOC7_Purse_Use_Val = "0";
            string SETLOC7_Purse_Use_Reverse_Vol = "0";
            string SETLOC7_Purse_Use_Reverse_Val = "0";
            string MallNgamwongwan_Purse_Use_Vol = "0";
            string MallNgamwongwan_Purse_Use_Val = "0";
            string MallNgamwongwan_Purse_Use_Reverse_Vol = "0";
            string MallNgamwongwan_Purse_Use_Reverse_Val = "0";
            string MallNgamwongwan_Purse_Add_Vol = "0";
            string MallNgamwongwan_Purse_Add_Val = "0";
            string MallNgamwongwan_Purse_Add_Reverse_Vol = "0";
            string MallNgamwongwan_Purse_Add_Reverse_Val = "0";
            string CYBERWORLD_Purse_Use_Vol = "0";
            string CYBERWORLD_Purse_Use_Val = "0";
            string CYBERWORLD_Purse_Use_Reverse_Vol = "0";
            string CYBERWORLD_Purse_Use_Reverse_Val = "0";
            string CYBERWORLD_Purse_Add_Vol = "0";
            string CYBERWORLD_Purse_Add_Val = "0";
            string CYBERWORLD_Purse_Add_Reverse_Vol = "0";
            string CYBERWORLD_Purse_Add_Reverse_Val = "0";


            int dm = 0;
            DataTable ds2 = new DataTable();
            ds2 = fn.connect_purse();
            foreach (DataRow drCurrent in ds2.Rows)
            {
                RPT_GRP[dm] = drCurrent["RPT_GRP"].ToString();
                TXN_TYPE[dm] = drCurrent["TXN_TYPE"].ToString();
                if (drCurrent["TXN_VOL"].ToString().Length > 3) { TXN_VOL[dm] = Convert.ToDecimal(drCurrent["TXN_VOL"]).ToString("0,000"); }
                if (drCurrent["TXN_VOL"].ToString().Length <= 3) { TXN_VOL[dm] = Convert.ToDecimal(drCurrent["TXN_VOL"]).ToString(); }
                if (drCurrent["TXN_VALUE"].ToString().Length > 3) TXN_VALUE[dm] = Convert.ToDecimal(drCurrent["TXN_VALUE"]).ToString("0,000");
                if (drCurrent["TXN_VALUE"].ToString().Length <= 3) TXN_VALUE[dm] = Convert.ToDecimal(drCurrent["TXN_VALUE"]).ToString();

                dm++;
            }

            for (dm = 0; dm < TXN_TYPE.Length; dm++)
            {
                if (RPT_GRP[dm] == "BBL")
                {
                    if (TXN_TYPE[dm] == "Purse Add")
                    {
                        BBL_Purse_Add_Val = TXN_VALUE[dm];
                        BBL_Purse_Add_Vol = TXN_VOL[dm];
                    }
                    if (TXN_TYPE[dm] == "Purse Use")
                    {
                        BBL_Purse_Use_Val = TXN_VALUE[dm];
                        BBL_Purse_Use_Vol = TXN_VOL[dm];
                    }
                    if (TXN_TYPE[dm] == "Purse Use Reverse")
                    {
                        BBL_Purse_Use_Reverse = TXN_VOL[dm];
                    }
                }
                if (RPT_GRP[dm] == "BCC")
                {
                    if (TXN_TYPE[dm] == "Card Refund") { BCC_Card_Refund = TXN_VOL[dm]; }
                    if (TXN_TYPE[dm] == "Purse Refund") { BCC_Purse_Refund_Val = TXN_VALUE[dm]; BCC_Purse_Refund_Vol = TXN_VOL[dm]; }
                    if (TXN_TYPE[dm] == "Card Initialise") { BCC_Card_Initialise = TXN_VOL[dm]; }
                    if (TXN_TYPE[dm] == "Card Issue") { BCC_Card_Issue = TXN_VOL[dm]; }
                    if (TXN_TYPE[dm] == "Card Surrender") { BCC_Card_Surrender = TXN_VOL[dm]; }
                    if (TXN_TYPE[dm] == "Card Refund Deferre") { BCC_Card_Refund_Deferred = TXN_VOL[dm]; }
                    if (TXN_TYPE[dm] == "Purse Add") { BCC_Purse_Add_Vol = TXN_VOL[dm]; BCC_Purse_Add_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Issue") { BCC_Purse_Issue_Vol = TXN_VOL[dm]; BCC_Purse_Issue_Val = TXN_VALUE[dm]; }
                }
                if (RPT_GRP[dm] == "BTS")
                {
                    if (TXN_TYPE[dm] == "Card Issue") { BTS_Card_Issue = TXN_VOL[dm]; }
                    if (TXN_TYPE[dm] == "Card Refund") { BTS_Card_Refund = TXN_VOL[dm]; }
                    if (TXN_TYPE[dm] == "Card Use Failed") { BTS_Card_Use_Failed = TXN_VOL[dm]; }
                    if (TXN_TYPE[dm] == "Multiride Issue") { BTS_Multiride_Issue_Vol = TXN_VOL[dm]; BTS_Multiride_Issue_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Multiride Issue Reverse") { BTS_Multiride_Issue_Reverse = TXN_VOL[dm]; }
                    if (TXN_TYPE[dm] == "Multiride Use On Entry") { BTS_Multiride_Use_On_Entry_Vol = TXN_VOL[dm]; BTS_Multiride_Use_On_Entry_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Multiride Use On Exit") { BTS_Multiride_Use_On_Exit_Vol = TXN_VOL[dm]; BTS_Multiride_Use_On_Exit_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Add") { BTS_Purse_Add_Vol = TXN_VOL[dm]; BTS_Purse_Add_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Add Reverse") { BTS_Purse_Add_Reverse = TXN_VOL[dm]; }
                    if (TXN_TYPE[dm] == "Purse Compensation Fare") { BTS_Purse_Compensation_Fare_Vol = TXN_VOL[dm]; BTS_Purse_Compensation_Fare_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Issue") { BTS_Purse_Issue_Vol = TXN_VOL[dm]; BTS_Purse_Issue_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Refund") { BTS_Purse_Refund = TXN_VOL[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use") { BTS_Purse_Use = TXN_VOL[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use On Entry") { BTS_Purse_Use_On_Entry_Vol = TXN_VOL[dm]; BTS_Purse_Use_On_Entry_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use On Exit") { BTS_Purse_Use_On_Exit_Vol = TXN_VOL[dm]; BTS_Purse_Use_On_Exit_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use Reverse") { BTS_Purse_Use_Reverse = TXN_VOL[dm]; }
                    if (TXN_TYPE[dm] == "Card Block") { BTS_Card_Block = TXN_VOL[dm]; }

                }
                if (RPT_GRP[dm] == "BRT")
                {

                    if (TXN_TYPE[dm] == "Card Issue") { BRT_Card_Issue = TXN_VOL[dm]; }
                    if (TXN_TYPE[dm] == "Card Use Failed") { BRT_Card_Use_Failed = TXN_VOL[dm]; }
                    if (TXN_TYPE[dm] == "Purse Add") { BRT_Purse_Add_Vol = TXN_VOL[dm]; BRT_Purse_Add_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Compensation Fare") { BRT_Purse_Compensation_Fare_Vol = TXN_VOL[dm]; BRT_Purse_Compensation_Fare_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Issue") { BRT_Purse_Issue_Vol = TXN_VOL[dm]; BRT_Purse_Issue_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use On Entry") { BRT_Purse_Use_On_Entry_Vol = TXN_VOL[dm]; BRT_Purse_Use_On_Entry_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use On Exit") { BRT_Purse_Use_On_Exit_Vol = TXN_VOL[dm]; BRT_Purse_Use_On_Exit_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Card Refund") { BRT_Card_Refund = TXN_VOL[dm]; }
                    if (TXN_TYPE[dm] == "Card Block") { BRT_Card_Block = TXN_VOL[dm]; }
                }
                if (RPT_GRP[dm] == "BIGC")
                {
                    if (TXN_TYPE[dm] == "Purse Add") { BIGC_Purse_Add_Vol = TXN_VOL[dm]; BIGC_Purse_Add_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use") { BIGC_Purse_Use_Vol = TXN_VOL[dm]; BIGC_Purse_Use_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use Reverse") { BIGC_Purse_Use_Reverse = TXN_VOL[dm]; }
                    if (TXN_TYPE[dm] == "Purse Add Reverse") { BIGC_Purse_Add_Reverse = TXN_VOL[dm]; }
                }
                if (RPT_GRP[dm] == "CARR")
                {
                    if (TXN_TYPE[dm] == "Purse Add") { CARR_Purse_Add_Vol = TXN_VOL[dm]; CARR_Purse_Add_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Card Use Failed") { CARR_Card_Use_Failed = TXN_VOL[dm]; }

                }
                if (RPT_GRP[dm] == "EMPORIUM")
                {
                    if (TXN_TYPE[dm] == "Purse Add") { EMPORIUM_Purse_Add_Vol = TXN_VOL[dm]; EMPORIUM_Purse_Add_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use") { EMPORIUM_Purse_Use_Vol = TXN_VOL[dm]; EMPORIUM_Purse_Use_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Add Reverse") { EMPORIUM_Purse_Add_Reverse_Vol = TXN_VOL[dm]; EMPORIUM_Purse_Add_Reverse_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use Reverse") { EMPORIUM_Purse_Use_Reverse_Vol = TXN_VOL[dm]; EMPORIUM_Purse_Use_Reverse_Val = TXN_VALUE[dm]; }
                }
                if (RPT_GRP[dm] == "EOF")
                {
                    if (TXN_TYPE[dm] == "Purse Add") { EOF_Purse_Add_Vol = TXN_VOL[dm]; EOF_Purse_Add_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Add Reverse") { EOF_Purse_Add_Reverse_Vol = TXN_VOL[dm]; EOF_Purse_Add_Reverse_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use") { EOF_Purse_Use_Vol = TXN_VOL[dm]; EOF_Purse_Use_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use Reverse") { EOF_Purse_Use_Reverse_Vol = TXN_VOL[dm]; EOF_Purse_Use_Reverse_Val = TXN_VALUE[dm]; }
                }
                if (RPT_GRP[dm] == "FOOD")
                {
                    if (TXN_TYPE[dm] == "Purse Add") { FOOD_Purse_Add_Vol = TXN_VOL[dm]; FOOD_Purse_Add_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use") { FOOD_Purse_Use_Vol = TXN_VOL[dm]; FOOD_Purse_Use_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Add Reverse") { FOOD_Purse_Add_Reverse_Vol = TXN_VOL[dm]; FOOD_Purse_Add_Reverse_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use Reverse") { FOOD_Purse_Use_Reverse_Vol = TXN_VOL[dm]; FOOD_Purse_Use_Reverse_Val = TXN_VALUE[dm]; }
                }
                if (RPT_GRP[dm] == "FOOD STREET")
                {
                    if (TXN_TYPE[dm] == "Purse Add") { FOOD_STREET_Purse_Add_Vol = TXN_VOL[dm]; FOOD_STREET_Purse_Add_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use") { FOOD_STREET_Purse_Use_Vol = TXN_VOL[dm]; FOOD_STREET_Purse_Use_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Add Reverse") { FOOD_STREET_Purse_Add_Reverse_Vol = TXN_VOL[dm]; FOOD_STREET_Purse_Add_Reverse_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use Reverse") { FOOD_STREET_Purse_Use_Reverse_Vol = TXN_VOL[dm]; FOOD_STREET_Purse_Use_Reverse_Val = TXN_VALUE[dm]; }
                }
                if (RPT_GRP[dm] == "LEGACY")
                {
                    if (TXN_TYPE[dm] == "Purse Add") { LEGACY_Purse_Add_Vol = TXN_VOL[dm]; LEGACY_Purse_Add_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Add Reverse") { LEGACY_Purse_Add_Reverse_Vol = TXN_VOL[dm]; LEGACY_Purse_Add_Reverse_Val = TXN_VALUE[dm]; }

                }
                if (RPT_GRP[dm] == "MBK")
                {
                    if (TXN_TYPE[dm] == "Purse Add") { MBK_Purse_Add_Vol = TXN_VOL[dm]; MBK_Purse_Add_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use") { MBK_Purse_Use_Vol = TXN_VOL[dm]; MBK_Purse_Use_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Add Reverse") { MBK_Purse_Add_Reverse_Vol = TXN_VOL[dm]; MBK_Purse_Add_Reverse_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use Reverse") { MBK_Purse_Use_Reverse_Vol = TXN_VOL[dm]; MBK_Purse_Use_Reverse_Val = TXN_VALUE[dm]; }
                }
                if (RPT_GRP[dm] == "MCDIRECT")
                {
                    if (TXN_TYPE[dm] == "Card Block") { MCDIRECT_Card_Block = TXN_VOL[dm]; }
                    if (TXN_TYPE[dm] == "Purse Add") { MCDIRECT_Purse_Add_Vol = TXN_VOL[dm]; MCDIRECT_Purse_Add_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use") { MCDIRECT_Purse_Use_Vol = TXN_VOL[dm]; MCDIRECT_Purse_Use_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Card Block") { MCDIRECT_Card_Block = TXN_VOL[dm]; }
                }
                if (RPT_GRP[dm] == "MPAY")
                {
                    if (TXN_TYPE[dm] == "Card Personalise") { MPAY_Card_Personalise = TXN_VOL[dm]; }
                    if (TXN_TYPE[dm] == "Purse Add") { MPAY_Purse_Add_Vol = (Convert.ToDecimal(MPAY_Purse_Add_Vol) + Convert.ToDecimal(TXN_VOL[dm])).ToString(); MPAY_Purse_Add_Val = (Convert.ToDecimal(MPAY_Purse_Add_Val) + Convert.ToDecimal(TXN_VALUE[dm])).ToString(); }

                }
                if (RPT_GRP[dm] == "TESCO")
                {
                    if (TXN_TYPE[dm] == "Purse Add") { TESCO_Purse_Add_Vol = TXN_VOL[dm]; TESCO_Purse_Add_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use") { TESCO_Purse_Use_Vol = TXN_VOL[dm]; TESCO_Purse_Use_Val = TXN_VALUE[dm]; }
                }
                if (RPT_GRP[dm] == "THAPRA")
                {
                    if (TXN_TYPE[dm] == "Purse Add") { THAPRA_Purse_Add_Vol = TXN_VOL[dm]; THAPRA_Purse_Add_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use") { THAPRA_Purse_Use_Vol = TXN_VOL[dm]; THAPRA_Purse_Use_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Add Reverse") { THAPRA_Purse_Add_Reverse_Vol = TXN_VOL[dm]; THAPRA_Purse_Add_Reverse_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use Reverse") { THAPRA_Purse_Use_Reverse_Vol = TXN_VOL[dm]; THAPRA_Purse_Use_Reverse_Val = TXN_VALUE[dm]; }
                }

                if (RPT_GRP[dm] == "BANGKAE")
                {
                    if (TXN_TYPE[dm] == "Purse Add") { BANGKAE_Purse_Add_Vol = TXN_VOL[dm]; BANGKAE_Purse_Add_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use") { BANGKAE_Purse_Use_Vol = TXN_VOL[dm]; BANGKAE_Purse_Use_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Add Reverse") { BANGKAE_Purse_Add_Reverse_Vol = TXN_VOL[dm]; BANGKAE_Purse_Add_Reverse_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use Reverse") { BANGKAE_Purse_Use_Reverse_Vol = TXN_VOL[dm]; BANGKAE_Purse_Use_Reverse_Val = TXN_VALUE[dm]; }
                }

                if (RPT_GRP[dm] == "FSEKAMAI")
                {
                    if (TXN_TYPE[dm] == "Purse Add") { FSEKAMAI_Purse_Add_Vol = TXN_VOL[dm]; FSEKAMAI_Purse_Add_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use") { FSEKAMAI_Purse_Use_Vol = TXN_VOL[dm]; FSEKAMAI_Purse_Use_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Add Reverse") { FSEKAMAI_Purse_Add_Reverse_Vol = TXN_VOL[dm]; FSEKAMAI_Purse_Add_Reverse_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use Reverse") { FSEKAMAI_Purse_Use_Reverse_Vol = TXN_VOL[dm]; FSEKAMAI_Purse_Use_Reverse_Val = TXN_VALUE[dm]; }
                }
                if (RPT_GRP[dm] == "ATU")
                {
                    if (TXN_TYPE[dm] == "Autoload Change of Service") { ATU_Change_of_Service = TXN_VOL[dm]; }
                    if (TXN_TYPE[dm] == "Product Autoload Disable") { ATU_Product_Autoload_Disable = TXN_VOL[dm]; }
                    if (TXN_TYPE[dm] == "Product Autoload Enable") { ATU_Product_Autoload_Enable = TXN_VOL[dm]; }

                }
                if (RPT_GRP[dm] == "IMPERIAL 4F")
                {
                    if (TXN_TYPE[dm] == "Purse Add") { IMPERIAL4F_Purse_Add_Vol = TXN_VOL[dm]; IMPERIAL4F_Purse_Add_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use") { IMPERIAL4F_Purse_Use_Vol = TXN_VOL[dm]; IMPERIAL4F_Purse_Use_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Add Reverse") { IMPERIAL4F_Purse_Add_Reverse_Vol = TXN_VOL[dm]; IMPERIAL4F_Purse_Add_Reverse_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use Reverse") { IMPERIAL4F_Purse_Use_Reverse_Vol = TXN_VOL[dm]; IMPERIAL4F_Purse_Use_Reverse_Val = TXN_VALUE[dm]; }
                }
                if (RPT_GRP[dm] == "IMPERIAL BF")
                {
                    if (TXN_TYPE[dm] == "Purse Add") { IMPERIALBF_Purse_Add_Vol = TXN_VOL[dm]; IMPERIALBF_Purse_Add_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use") { IMPERIALBF_Purse_Use_Vol = TXN_VOL[dm]; IMPERIALBF_Purse_Use_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Add Reverse") { IMPERIALBF_Purse_Add_Reverse_Vol = TXN_VOL[dm]; IMPERIALBF_Purse_Add_Reverse_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use Reverse") { IMPERIALBF_Purse_Use_Reverse_Vol = TXN_VOL[dm]; IMPERIALBF_Purse_Use_Reverse_Val = TXN_VALUE[dm]; }
                }
                if (RPT_GRP[dm] == "SET")
                {
                    if (TXN_TYPE[dm] == "Purse Add") { SET_Purse_Add_Vol = TXN_VOL[dm]; SET_Purse_Add_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Add Reverse") { SET_Purse_Add_Reverse_Vol = TXN_VOL[dm]; SET_Purse_Add_Reverse_Val = TXN_VALUE[dm]; }
                }
                if (RPT_GRP[dm] == "SET LOC 1")
                {
                    if (TXN_TYPE[dm] == "Purse Use") { SETLOC1_Purse_Use_Vol = TXN_VOL[dm]; SETLOC1_Purse_Use_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use Reverse") { SETLOC1_Purse_Use_Reverse_Vol = TXN_VOL[dm]; SETLOC1_Purse_Use_Reverse_Val = TXN_VALUE[dm]; }
                }
                if (RPT_GRP[dm] == "SET LOC 2")
                {
                    if (TXN_TYPE[dm] == "Purse Use") { SETLOC2_Purse_Use_Vol = TXN_VOL[dm]; SETLOC2_Purse_Use_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use Reverse") { SETLOC2_Purse_Use_Reverse_Vol = TXN_VOL[dm]; SETLOC2_Purse_Use_Reverse_Val = TXN_VALUE[dm]; }
                }
                if (RPT_GRP[dm] == "SET LOC 3")
                {
                    if (TXN_TYPE[dm] == "Purse Use") { SETLOC3_Purse_Use_Vol = TXN_VOL[dm]; SETLOC3_Purse_Use_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use Reverse") { SETLOC3_Purse_Use_Reverse_Vol = TXN_VOL[dm]; SETLOC3_Purse_Use_Reverse_Val = TXN_VALUE[dm]; }
                }
                if (RPT_GRP[dm] == "SET LOC 4")
                {
                    if (TXN_TYPE[dm] == "Purse Use") { SETLOC4_Purse_Use_Vol = TXN_VOL[dm]; SETLOC4_Purse_Use_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use Reverse") { SETLOC4_Purse_Use_Reverse_Vol = TXN_VOL[dm]; SETLOC4_Purse_Use_Reverse_Val = TXN_VALUE[dm]; }
                }
                if (RPT_GRP[dm] == "SET LOC 5")
                {
                    if (TXN_TYPE[dm] == "Purse Use") { SETLOC5_Purse_Use_Vol = TXN_VOL[dm]; SETLOC5_Purse_Use_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use Reverse") { SETLOC5_Purse_Use_Reverse_Vol = TXN_VOL[dm]; SETLOC5_Purse_Use_Reverse_Val = TXN_VALUE[dm]; }
                }
                if (RPT_GRP[dm] == "SET LOC 6")
                {
                    if (TXN_TYPE[dm] == "Purse Use") { SETLOC6_Purse_Use_Vol = TXN_VOL[dm]; SETLOC6_Purse_Use_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use Reverse") { SETLOC6_Purse_Use_Reverse_Vol = TXN_VOL[dm]; SETLOC6_Purse_Use_Reverse_Val = TXN_VALUE[dm]; }
                }
                if (RPT_GRP[dm] == "SET LOC 7")
                {
                    if (TXN_TYPE[dm] == "Purse Use") { SETLOC7_Purse_Use_Vol = TXN_VOL[dm]; SETLOC7_Purse_Use_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use Reverse") { SETLOC7_Purse_Use_Reverse_Vol = TXN_VOL[dm]; SETLOC7_Purse_Use_Reverse_Val = TXN_VALUE[dm]; }
                }
                if (RPT_GRP[dm] == "MALL NGAMWONGWAN")
                {
                    if (TXN_TYPE[dm] == "Purse Add") { MallNgamwongwan_Purse_Add_Vol = TXN_VOL[dm]; MallNgamwongwan_Purse_Add_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use") { MallNgamwongwan_Purse_Use_Vol = TXN_VOL[dm]; MallNgamwongwan_Purse_Use_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Add Reverse") { MallNgamwongwan_Purse_Add_Reverse_Vol = TXN_VOL[dm]; MallNgamwongwan_Purse_Add_Reverse_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use Reverse") { MallNgamwongwan_Purse_Use_Reverse_Vol = TXN_VOL[dm]; MallNgamwongwan_Purse_Use_Reverse_Val = TXN_VALUE[dm]; }
                }
                if (RPT_GRP[dm] == "CW TOWER")
                {
                    if (TXN_TYPE[dm] == "Purse Add") { CYBERWORLD_Purse_Add_Vol = TXN_VOL[dm]; CYBERWORLD_Purse_Add_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use") { CYBERWORLD_Purse_Use_Vol = TXN_VOL[dm]; CYBERWORLD_Purse_Use_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Add Reverse") { CYBERWORLD_Purse_Add_Reverse_Vol = TXN_VOL[dm]; CYBERWORLD_Purse_Add_Reverse_Val = TXN_VALUE[dm]; }
                    if (TXN_TYPE[dm] == "Purse Use Reverse") { CYBERWORLD_Purse_Use_Reverse_Vol = TXN_VOL[dm]; CYBERWORLD_Purse_Use_Reverse_Val = TXN_VALUE[dm]; }
                }
            }

            Microsoft.Office.Interop.Excel.Application excelfile = new Microsoft.Office.Interop.Excel.Application();

            if (excelfile == null)
            {
                Console.WriteLine("Excel is not properly installed!!");
                return;
            }


            object misValue = System.Reflection.Missing.Value;
            var excelworkbook = excelfile.Workbooks.Add(misValue);
            var excelworksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkbook.Worksheets.get_Item(1);
            excelworksheet.get_Range("a1", "ag133").NumberFormat = "#,###,##0";
            excelworksheet.get_Range("b1", "ag133").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
          //  excelworksheet.get_Range("f106", "f107").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            excelworksheet.get_Range("a1").ColumnWidth = 18;
            excelworksheet.get_Range("b1").ColumnWidth = 30;
            excelworksheet.get_Range("c1", "ag1").ColumnWidth = 10;
            excelworksheet.get_Range("b16", "ag19").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(189,215,238));
            //1
            var borders = excelworksheet.get_Range("b16", "ag19").Borders;
            borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            excelworksheet.get_Range("c19", "ag19").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            excelworksheet.get_Range("c19", "ag19").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //2
            excelworksheet.get_Range("b21", "r24").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(189,215,238));
            borders = excelworksheet.get_Range("b21", "r24").Borders;
            borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            excelworksheet.get_Range("c24", "r24").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            excelworksheet.get_Range("c24", "r24").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //3
            excelworksheet.get_Range("b26", "g29").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(189,215,238));
            borders = excelworksheet.get_Range("b26", "g29").Borders;
            borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            excelworksheet.get_Range("c29", "g29").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            excelworksheet.get_Range("c29", "g29").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //4
            excelworksheet.get_Range("b31", "h34").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(189,215,238));
            borders = excelworksheet.get_Range("b31", "h34").Borders;
            borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            excelworksheet.get_Range("c34", "h34").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            excelworksheet.get_Range("c34", "h34").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //5
            excelworksheet.get_Range("b36", "i39").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(189,215,238));
            borders = excelworksheet.get_Range("b36", "i39").Borders;
            borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            excelworksheet.get_Range("c39", "i39").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            excelworksheet.get_Range("c39", "i39").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //6
            excelworksheet.get_Range("b41", "g44").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(189,215,238));
            borders = excelworksheet.get_Range("b41", "g44").Borders;
            borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            excelworksheet.get_Range("c44", "g44").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            excelworksheet.get_Range("c44", "g44").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //7
            excelworksheet.get_Range("b46", "i49").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(189,215,238));
            borders = excelworksheet.get_Range("b46", "i49").Borders;
            borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            excelworksheet.get_Range("c49", "i49").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            excelworksheet.get_Range("c49", "i49").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //8
            excelworksheet.get_Range("b51", "l54").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(189,215,238));
            borders = excelworksheet.get_Range("b51", "l54").Borders;
            borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            excelworksheet.get_Range("c54", "l54").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            excelworksheet.get_Range("c54", "l54").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //9
            excelworksheet.get_Range("b56", "g59").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(189,215,238));
            borders = excelworksheet.get_Range("b56", "g59").Borders;
            borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            excelworksheet.get_Range("c59", "g59").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            excelworksheet.get_Range("c59", "g59").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //10
            excelworksheet.get_Range("b61", "h64").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(189,215,238));
            borders = excelworksheet.get_Range("b61", "h64").Borders;
            borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            excelworksheet.get_Range("c64", "h64").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            excelworksheet.get_Range("c64", "h64").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //11
            excelworksheet.get_Range("b66", "l69").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(189,215,238));
            borders = excelworksheet.get_Range("b66", "l69").Borders;
            borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            excelworksheet.get_Range("c69", "l69").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            excelworksheet.get_Range("c69", "l69").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //12
            excelworksheet.get_Range("b71", "l74").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(189,215,238));
            borders = excelworksheet.get_Range("b71", "l74").Borders;
            borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            excelworksheet.get_Range("c74", "l74").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            excelworksheet.get_Range("c74", "l74").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //13
            excelworksheet.get_Range("b76", "l79").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(189,215,238));
            borders = excelworksheet.get_Range("b76", "l79").Borders;
            borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            excelworksheet.get_Range("c79", "l79").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            excelworksheet.get_Range("c79", "l79").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //14
            excelworksheet.get_Range("b81", "l84").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(189,215,238));
            borders = excelworksheet.get_Range("b81", "l84").Borders;
            borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            excelworksheet.get_Range("c84", "l84").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            excelworksheet.get_Range("c84", "l84").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //15
            excelworksheet.get_Range("b86", "l89").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(189,215,238));
            borders = excelworksheet.get_Range("b86", "l89").Borders;
            borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            excelworksheet.get_Range("c89", "l89").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            excelworksheet.get_Range("c89", "l89").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //16
            excelworksheet.get_Range("b91", "l94").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(189,215,238));
            borders = excelworksheet.get_Range("b91", "l94").Borders;
            borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            excelworksheet.get_Range("c94", "l94").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            excelworksheet.get_Range("c94", "l94").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //17
            excelworksheet.get_Range("b96", "l99").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(189,215,238));
            borders = excelworksheet.get_Range("b96", "l99").Borders;
            borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            excelworksheet.get_Range("c99", "l99").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            excelworksheet.get_Range("c99", "l99").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //18
            excelworksheet.get_Range("b101", "l104").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(189, 215, 238));
            borders = excelworksheet.get_Range("b101", "l104").Borders;
            borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            excelworksheet.get_Range("c104", "l104").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            excelworksheet.get_Range("c104", "l104").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //19
            excelworksheet.get_Range("b106", "l109").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(189, 215, 238));
            borders = excelworksheet.get_Range("b106", "l109").Borders;
            borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            excelworksheet.get_Range("c109", "l109").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            excelworksheet.get_Range("c109", "l109").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //20
            excelworksheet.get_Range("b111", "l114").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(189, 215, 238));
            borders = excelworksheet.get_Range("b111", "l114").Borders;
            borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            excelworksheet.get_Range("c114", "l114").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            excelworksheet.get_Range("c114", "l114").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //21
            excelworksheet.get_Range("b116", "l119").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(189, 215, 238));
            borders = excelworksheet.get_Range("b116", "l119").Borders;
            borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            excelworksheet.get_Range("c119", "l119").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            excelworksheet.get_Range("c119", "l119").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //22
            excelworksheet.get_Range("b121", "f124").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(189,215,238));
            borders = excelworksheet.get_Range("b121", "f124").Borders;
            borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            excelworksheet.get_Range("c124", "f124").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            excelworksheet.get_Range("c124", "f124").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //Exception CCH
            excelworksheet.get_Range("b126", "d133").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(189,215,238));
            borders = excelworksheet.get_Range("b126", "d133").Borders;
            borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            excelworksheet.get_Range("d127", "d133").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            excelworksheet.get_Range("d127", "d133").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //Exception ISS
            excelworksheet.get_Range("h126", "l131").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(189,215,238));
            borders = excelworksheet.get_Range("h126", "l131").Borders;
            borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            excelworksheet.get_Range("l127", "l131").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            excelworksheet.get_Range("l127", "l131").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //eod
            excelworksheet.get_Range("b8", "b10").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(189,215,238));
            borders = excelworksheet.get_Range("b8", "c10").Borders;
            borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            excelworksheet.get_Range("b12", "b14").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(189,215,238));
            borders = excelworksheet.get_Range("b12", "c14").Borders;
            borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;


            excelworksheet.get_Range("c3", "c4").NumberFormat = "dd/mm/yyyy";
            excelworksheet.get_Range("c6", "c6").NumberFormat = "dd/mm/yyyy";


            // Microsoft.Office.Interop.Excel.Borders border = format.Borders;
            // border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            // border.Weight = 2d;
            //   excelworksheet.Cells[2,2].ColumnWidth = 20;

            excelworksheet.Cells[1, 1] = "System Day End Summary Report v1.0";
            excelworksheet.Cells[1, 1].Font.Bold = true;
            excelworksheet.Cells[1, 1].Font.Size = 18;
            excelworksheet.Cells[3, 1] = "Operator Detail";
            excelworksheet.Cells[3, 2] = "Date of check :";
            excelworksheet.Cells[4, 1] = "(Automatic System)";
           
            excelworksheet.Cells[3, 3] = DateTime.Now.ToString(CultureInfo.CreateSpecificCulture("en-US"));
            excelworksheet.get_Range("c3", "d3").Merge(false);
            excelworksheet.get_Range("c4", "d4").Merge(false);
            excelworksheet.get_Range("c6", "d6").Merge(false);
            excelworksheet.Cells[6, 1] = "Availability";
            excelworksheet.Cells[6, 2] = "Settlement Date :";
            excelworksheet.Cells[6, 3] = settlement_date;
            excelworksheet.Cells[8, 2] = "EOD Status";
            excelworksheet.Cells[8, 3] = status;
            excelworksheet.Cells[9, 2] = "SOD Status";
            excelworksheet.Cells[9, 3] = status;
            excelworksheet.Cells[10, 2] = "CD Distributed";
            excelworksheet.Cells[10, 3] = status;
            excelworksheet.Cells[12, 2] = "Report Distribution";
            excelworksheet.Cells[12, 3] = AUTOREPORT;
            excelworksheet.Cells[13, 2] = "Citrix Server";
            excelworksheet.Cells[13, 3] = CITRIX;
            excelworksheet.Cells[14, 2] = "BOI";
            excelworksheet.Cells[14, 3] = BOI;
            excelworksheet.Cells[16, 1] = "Integrity";
            excelworksheet.Cells[16, 2] = "1";
            excelworksheet.get_Range("b16", "b18").Merge(false);
            excelworksheet.Cells[16, 3] = "Files Level";
            excelworksheet.get_Range("c16", "d16").Merge(false);
            excelworksheet.Cells[16, 5] = "Card";
            excelworksheet.get_Range("e16", "i16").Merge(false);
            excelworksheet.Cells[16, 10] = "Purse";
            excelworksheet.get_Range("j16", "v16").Merge(false);
            excelworksheet.Cells[16, 23] = "Faulty use";
            excelworksheet.get_Range("w16", "x16").Merge(false);
            excelworksheet.Cells[16, 25] = "STP";
            excelworksheet.get_Range("Y16", "ag16").Merge(false);
            excelworksheet.Cells[17, 3] = "Received";
            excelworksheet.Cells[17, 4] = "Failed";
            excelworksheet.Cells[17, 5] = "Initialise";
            excelworksheet.Cells[17, 6] = "Issue(BTS)";
            excelworksheet.Cells[17, 7] = "Issue(BCC)";
            excelworksheet.Cells[17, 8] = "Surrender";
            excelworksheet.Cells[17, 9] = "Refund";
            excelworksheet.Cells[17, 10] = "Compensation Fare";
            excelworksheet.get_Range("j17", "k17").Merge(false);
            excelworksheet.Cells[17, 12] = "Purse Add";
            excelworksheet.Cells[17, 13] = "Purse Issue(BTS)";
            excelworksheet.get_Range("m17", "n17").Merge(false);
            excelworksheet.Cells[17, 15] = "Purse Issue(BCC)";
            excelworksheet.get_Range("o17", "p17").Merge(false);
            excelworksheet.Cells[17, 17] = "Purse Add";
            excelworksheet.get_Range("q17", "r17").Merge(false);
            excelworksheet.Cells[17, 19] = "Purse Entry";
            excelworksheet.get_Range("s17", "t17").Merge(false);
            excelworksheet.Cells[17, 21] = "Purse Exit";
            excelworksheet.get_Range("u17", "v17").Merge(false);
            excelworksheet.Cells[17, 23] = "Blocked";
            excelworksheet.Cells[17, 24] = "Rejected";
            excelworksheet.Cells[17, 25] = "Compensation Fare";
            excelworksheet.get_Range("y17", "z17").Merge(false);
            excelworksheet.Cells[17, 27] = "STU Issue";
            excelworksheet.Cells[17, 28] = "STP Issue";
            excelworksheet.get_Range("ab17", "ac17").Merge(false);
            excelworksheet.Cells[17, 30] = "STP Entry";
            excelworksheet.get_Range("ad17", "ae17").Merge(false);
            excelworksheet.Cells[17, 32] = "STP Exit";
            excelworksheet.get_Range("af17", "ag17").Merge(false);
            excelworksheet.Cells[18, 3] = "Vol";
            excelworksheet.Cells[18, 4] = "Vol";
            excelworksheet.Cells[18, 5] = "Vol";
            excelworksheet.Cells[18, 6] = "Vol";
            excelworksheet.Cells[18, 7] = "Vol";
            excelworksheet.Cells[18, 8] = "Vol";
            excelworksheet.Cells[18, 9] = "Vol";
            excelworksheet.Cells[18, 10] = "Vol";
            excelworksheet.Cells[18, 11] = "Val";
            excelworksheet.Cells[18, 12] = "Reverse Vol";
            excelworksheet.Cells[18, 13] = "Vol";
            excelworksheet.Cells[18, 14] = "Val";
            excelworksheet.Cells[18, 15] = "Vol";
            excelworksheet.Cells[18, 16] = "Val";
            excelworksheet.Cells[18, 17] = "Vol";
            excelworksheet.Cells[18, 18] = "Val";
            excelworksheet.Cells[18, 19] = "Vol";
            excelworksheet.Cells[18, 20] = "Val";
            excelworksheet.Cells[18, 21] = "Vol";
            excelworksheet.Cells[18, 22] = "Val";
            excelworksheet.Cells[18, 23] = "Vol";
            excelworksheet.Cells[18, 24] = "Vol";
            excelworksheet.Cells[18, 25] = "Val";
            excelworksheet.Cells[18, 26] = "Vol";
            excelworksheet.Cells[18, 27] = "Reverse Vol";
            excelworksheet.Cells[18, 28] = "Vol";
            excelworksheet.Cells[18, 29] = "Val";
            excelworksheet.Cells[18, 30] = "Vol";
            excelworksheet.Cells[18, 31] = "Val";
            excelworksheet.Cells[18, 32] = "Vol";
            excelworksheet.Cells[18, 33] = "Val";
            excelworksheet.Cells[19, 2] = "BTS/BCC Transactions";
            excelworksheet.Cells[19, 3] = (Convert.ToDecimal(bts_num_received_file) + Convert.ToDecimal(bcc_num_received_file));
            excelworksheet.Cells[19, 4] = (Convert.ToDecimal(bts_num_failed_file) + Convert.ToDecimal(bcc_num_failed_file));
            excelworksheet.Cells[19, 5] = BCC_Card_Initialise;
            excelworksheet.Cells[19, 6] = BTS_Card_Issue;
            excelworksheet.Cells[19, 7] = BCC_Card_Issue;
            excelworksheet.Cells[19, 8] = BCC_Card_Surrender;
            excelworksheet.Cells[19, 9] = (Convert.ToDecimal(BCC_Card_Refund) + Convert.ToDecimal(BTS_Card_Refund) + Convert.ToDecimal(BCC_Card_Refund_Deferred));
            excelworksheet.Cells[19, 10] = BTS_Purse_Compensation_Fare_Vol;
            excelworksheet.Cells[19, 11] = BTS_Purse_Compensation_Fare_Val;
            excelworksheet.Cells[19, 12] = BTS_Purse_Add_Reverse;
            excelworksheet.Cells[19, 13] = BTS_Purse_Issue_Vol;
            excelworksheet.Cells[19, 14] = BTS_Purse_Issue_Val;
            excelworksheet.Cells[19, 15] = BCC_Purse_Issue_Vol;
            excelworksheet.Cells[19, 16] = BCC_Purse_Issue_Val;
            excelworksheet.Cells[19, 17] = (Convert.ToDecimal(BCC_Purse_Add_Vol) + Convert.ToDecimal(BTS_Purse_Add_Vol));
            excelworksheet.Cells[19, 18] = (Convert.ToDecimal(BCC_Purse_Add_Val) + Convert.ToDecimal(BTS_Purse_Add_Val));
            excelworksheet.Cells[19, 19] = BTS_Purse_Use_On_Entry_Vol;
            excelworksheet.Cells[19, 20] = BTS_Purse_Use_On_Entry_Val;
            excelworksheet.Cells[19, 21] = BTS_Purse_Use_On_Exit_Vol;
            excelworksheet.Cells[19, 22] = BTS_Purse_Use_On_Exit_Val;
            excelworksheet.Cells[19, 23] = BTS_Card_Block;
            excelworksheet.Cells[19, 24] = BTS_Card_Use_Failed;
            excelworksheet.Cells[19, 25] = "0";
            excelworksheet.Cells[19, 26] = "0";
            excelworksheet.Cells[19, 27] = BTS_Multiride_Issue_Reverse;
            excelworksheet.Cells[19, 28] = BTS_Multiride_Issue_Vol;
            excelworksheet.Cells[19, 29] = BTS_Multiride_Issue_Val;
            excelworksheet.Cells[19, 30] = BTS_Multiride_Use_On_Entry_Vol;
            excelworksheet.Cells[19, 31] = BTS_Multiride_Use_On_Entry_Val;
            excelworksheet.Cells[19, 32] = BTS_Multiride_Use_On_Exit_Vol;
            excelworksheet.Cells[19, 33] = BTS_Multiride_Use_On_Exit_Val;
            // end BTS start BRT
            excelworksheet.Cells[21, 2] = "2";
            excelworksheet.get_Range("b21", "b23").Merge(false);
            excelworksheet.Cells[21, 3] = "Files Level";
            excelworksheet.get_Range("c21", "d21").Merge(false);
            excelworksheet.Cells[21, 5] = "Card";
            excelworksheet.get_Range("e21", "f21").Merge(false);
            excelworksheet.Cells[21, 10] = "Purse";
            excelworksheet.get_Range("g21", "p21").Merge(false);
            excelworksheet.Cells[21, 17] = "Faulty use";
            excelworksheet.get_Range("q21", "r21").Merge(false);
            excelworksheet.Cells[22, 3] = "Received";
            excelworksheet.Cells[22, 4] = "Failed";
            excelworksheet.Cells[22, 5] = "Issue";
            excelworksheet.Cells[22, 6] = "Refund";
            excelworksheet.Cells[22, 7] = "Compensation Fare";
            excelworksheet.get_Range("g22", "h22").Merge(false);
            excelworksheet.Cells[22, 9] = "Issue";
            excelworksheet.get_Range("i22", "j22").Merge(false);
            excelworksheet.Cells[22, 11] = "Purse Add";
            excelworksheet.get_Range("k22", "l22").Merge(false);
            excelworksheet.Cells[22, 13] = "Purse Entry";
            excelworksheet.get_Range("m22", "n22").Merge(false);
            excelworksheet.Cells[22, 15] = "Purse Exit";
            excelworksheet.get_Range("o22", "p22").Merge(false);
            excelworksheet.Cells[22, 17] = "Blocked";
            excelworksheet.Cells[22, 18] = "Rejected";
            excelworksheet.Cells[23, 3] = "Vol";
            excelworksheet.Cells[23, 4] = "Vol";
            excelworksheet.Cells[23, 5] = "Vol";
            excelworksheet.Cells[23, 6] = "Vol";
            excelworksheet.Cells[23, 7] = "Vol";
            excelworksheet.Cells[23, 8] = "Val";
            excelworksheet.Cells[23, 9] = "Vol";
            excelworksheet.Cells[23, 10] = "Val";
            excelworksheet.Cells[23, 11] = "Vol";
            excelworksheet.Cells[23, 12] = "Val";
            excelworksheet.Cells[23, 13] = "Vol";
            excelworksheet.Cells[23, 14] = "Val";
            excelworksheet.Cells[23, 15] = "Vol";
            excelworksheet.Cells[23, 16] = "Val";
            excelworksheet.Cells[23, 17] = "Vol";
            excelworksheet.Cells[23, 18] = "Vol";
            excelworksheet.Cells[24, 2] = "BRT Transactions";
            excelworksheet.Cells[24, 3] = brt_num_received_file;
            excelworksheet.Cells[24, 4] = brt_num_failed_file;
            excelworksheet.Cells[24, 5] = BRT_Card_Issue;
            excelworksheet.Cells[24, 6] = BRT_Card_Refund;
            excelworksheet.Cells[24, 7] = BRT_Purse_Compensation_Fare_Vol;
            excelworksheet.Cells[24, 8] = BRT_Purse_Compensation_Fare_Val;
            excelworksheet.Cells[24, 9] = BRT_Purse_Issue_Vol;
            excelworksheet.Cells[24, 10] = BRT_Purse_Issue_Val;
            excelworksheet.Cells[24, 11] = BRT_Purse_Add_Vol;
            excelworksheet.Cells[24, 12] = BRT_Purse_Add_Val;
            excelworksheet.Cells[24, 13] = BRT_Purse_Use_On_Entry_Vol;
            excelworksheet.Cells[24, 14] = BRT_Purse_Use_On_Entry_Val;
            excelworksheet.Cells[24, 15] = BRT_Purse_Use_On_Exit_Vol;
            excelworksheet.Cells[24, 16] = BRT_Purse_Use_On_Exit_Val;
            excelworksheet.Cells[24, 17] = BRT_Card_Block;
            excelworksheet.Cells[24, 18] = BRT_Card_Use_Failed;
            //END BRT start BBL
            excelworksheet.Cells[26, 2] = "3";
            excelworksheet.get_Range("b26", "b28").Merge(false);
            excelworksheet.Cells[26, 3] = "Files Level";
            excelworksheet.get_Range("c26", "d26").Merge(false);
            excelworksheet.Cells[26, 5] = "Purse";
            excelworksheet.get_Range("e26", "g26").Merge(false);
            excelworksheet.Cells[27, 3] = "Received";
            excelworksheet.Cells[27, 4] = "Failed";
            excelworksheet.Cells[27, 5] = "Reverse";
            excelworksheet.Cells[27, 6] = "Usage";
            excelworksheet.get_Range("f27", "g27").Merge(false);
            excelworksheet.Cells[28, 3] = "Vol";
            excelworksheet.Cells[28, 4] = "Vol";
            excelworksheet.Cells[28, 5] = "Vol";
            excelworksheet.Cells[28, 6] = "Vol";
            excelworksheet.Cells[28, 7] = "Val";
            excelworksheet.Cells[29, 2] = "Retail Transactions(BBL)";
            excelworksheet.Cells[29, 3] = bbl_num_received_file;
            excelworksheet.Cells[29, 4] = bbl_num_failed_file;
            excelworksheet.Cells[29, 5] = BBL_Purse_Use_Reverse;
            excelworksheet.Cells[29, 6] = BBL_Purse_Use_Vol;
            excelworksheet.Cells[29, 7] = Math.Round(Convert.ToDecimal(BBL_Purse_Use_Val));
            //end BBL start carrot
            excelworksheet.Cells[31, 2] = "4";
            excelworksheet.get_Range("b31", "b33").Merge(false);
            excelworksheet.Cells[31, 3] = "Files Level";
            excelworksheet.get_Range("c31", "d31").Merge(false);
            excelworksheet.Cells[31, 5] = "Purse";
            excelworksheet.get_Range("e31", "f31").Merge(false);
            excelworksheet.Cells[31, 7] = "Faulty Use";
            excelworksheet.get_Range("g31", "h31").Merge(false);
            excelworksheet.Cells[32, 3] = "Received";
            excelworksheet.Cells[32, 4] = "Failed";
            excelworksheet.Cells[32, 5] = "Add";
            excelworksheet.get_Range("e32", "f32").Merge(false);
            excelworksheet.Cells[32, 7] = "Blocked";
            excelworksheet.Cells[32, 8] = "Rejected";
            excelworksheet.Cells[33, 3] = "Vol";
            excelworksheet.Cells[33, 4] = "Vol";
            excelworksheet.Cells[33, 5] = "Vol";
            excelworksheet.Cells[33, 6] = "Val";
            excelworksheet.Cells[33, 7] = "Vol";
            excelworksheet.Cells[33, 8] = "Vol";
            excelworksheet.Cells[34, 2] = "Carrot Transations";
            excelworksheet.Cells[34, 3] = crr_num_received_file;
            excelworksheet.Cells[34, 4] = crr_num_failed_file;
            excelworksheet.Cells[34, 5] = CARR_Purse_Add_Vol;
            excelworksheet.Cells[34, 6] = CARR_Purse_Add_Val;
            excelworksheet.Cells[34, 7] = CARR_Card_Block;
            excelworksheet.Cells[34, 8] = CARR_Card_Use_Failed;
            //end carrot
            excelworksheet.Cells[36, 2] = "5";
            excelworksheet.get_Range("b36", "b38").Merge(false);
            excelworksheet.Cells[36, 3] = "Files Level";
            excelworksheet.get_Range("c36", "d36").Merge(false);
            excelworksheet.Cells[36, 5] = "Purse";
            excelworksheet.get_Range("e36", "i36").Merge(false);
            excelworksheet.Cells[37, 3] = "Received";
            excelworksheet.Cells[37, 4] = "Failed";
            excelworksheet.Cells[37, 5] = "Reverse";
            excelworksheet.Cells[37, 6] = "Add";
            excelworksheet.get_Range("f37", "g37").Merge(false);
            excelworksheet.Cells[37, 8] = "Usage";
            excelworksheet.get_Range("h37", "i37").Merge(false);
            excelworksheet.Cells[39, 3] = "Vol";
            excelworksheet.Cells[38, 4] = "Vol";
            excelworksheet.Cells[38, 5] = "Vol";
            excelworksheet.Cells[38, 6] = "Vol";
            excelworksheet.Cells[38, 7] = "Val";
            excelworksheet.Cells[38, 8] = "Vol";
            excelworksheet.Cells[38, 9] = "Val";
            excelworksheet.Cells[39, 2] = "Mcthai Direct";
            excelworksheet.Cells[39, 3] = maq_num_received_file;
            excelworksheet.Cells[39, 4] = maq_num_failed_file;
            excelworksheet.Cells[39, 5] = (MCDIRECT_Purse_Add_Reverse + MCDIRECT_Purse_Use_Reverse);
            excelworksheet.Cells[39, 6] = MCDIRECT_Purse_Add_Vol;
            excelworksheet.Cells[39, 7] = MCDIRECT_Purse_Add_Val;
            excelworksheet.Cells[39, 8] = MCDIRECT_Purse_Use_Vol;
            excelworksheet.Cells[39, 9] = MCDIRECT_Purse_Use_Val;

            //end mcthai
            excelworksheet.Cells[41, 2] = "6";
            excelworksheet.get_Range("b41", "b43").Merge(false);
            excelworksheet.Cells[41, 3] = "Files Level";
            excelworksheet.get_Range("c41", "d41").Merge(false);
            excelworksheet.Cells[41, 5] = "Purse";
            excelworksheet.get_Range("e41", "g41").Merge(false);
            excelworksheet.Cells[42, 3] = "Received";
            excelworksheet.Cells[42, 4] = "Failed";
            excelworksheet.Cells[42, 5] = "Reverse";
            excelworksheet.Cells[42, 6] = "Add";
            excelworksheet.get_Range("f42", "g42").Merge(false);
            excelworksheet.Cells[43, 3] = "Vol";
            excelworksheet.Cells[43, 4] = "Vol";
            excelworksheet.Cells[43, 5] = "Vol";
            excelworksheet.Cells[43, 6] = "Vol";
            excelworksheet.Cells[43, 7] = "Val";
            excelworksheet.Cells[44, 2] = "Mpay Mobile";
            excelworksheet.Cells[44, 3] = mpay_num_received_file;
            excelworksheet.Cells[44, 4] = mpay_num_failed_file;
            excelworksheet.Cells[44, 5] = MPAY_Purse_Add_Reverse;
            excelworksheet.Cells[44, 6] = MPAY_Purse_Add_Vol;
            excelworksheet.Cells[44, 7] = MPAY_Purse_Add_Val;

            //end mpay

            excelworksheet.Cells[46, 2] = "7";
            excelworksheet.get_Range("b46", "b48").Merge(false);
            excelworksheet.Cells[46, 3] = "Files Level";
            excelworksheet.get_Range("c46", "d46").Merge(false);
            excelworksheet.Cells[46, 5] = "Purse";
            excelworksheet.get_Range("e46", "i46").Merge(false);
            excelworksheet.Cells[47, 3] = "Received";
            excelworksheet.Cells[47, 4] = "Failed";
            excelworksheet.Cells[47, 5] = "Reverse";
            excelworksheet.Cells[47, 6] = "Add";
            excelworksheet.get_Range("f47", "g47").Merge(false);
            excelworksheet.Cells[47, 8] = "Usage";
            excelworksheet.get_Range("h47", "i47").Merge(false);
            excelworksheet.Cells[48, 3] = "Vol";
            excelworksheet.Cells[48, 4] = "Vol";
            excelworksheet.Cells[48, 5] = "Vol";
            excelworksheet.Cells[48, 6] = "Vol";
            excelworksheet.Cells[48, 7] = "Val";
            excelworksheet.Cells[48, 8] = "Vol";
            excelworksheet.Cells[48, 9] = "Val";
            excelworksheet.Cells[49, 2] = "BIG-C";
            excelworksheet.Cells[49, 3] = bigc_num_received_file;
            excelworksheet.Cells[49, 4] = bigc_num_failed_file;
            excelworksheet.Cells[49, 5] = (BIGC_Purse_Add_Reverse + BIGC_Purse_Use_Reverse);
            excelworksheet.Cells[49, 6] = BIGC_Purse_Add_Vol;
            excelworksheet.Cells[49, 7] = BIGC_Purse_Add_Val;
            excelworksheet.Cells[49, 8] = BIGC_Purse_Use_Vol;
            excelworksheet.Cells[49, 9] = Math.Round(Convert.ToDecimal(BIGC_Purse_Use_Val));


            //end BIG-C

            excelworksheet.Cells[51, 2] = "8";
            excelworksheet.get_Range("b51", "b53").Merge(false);
            excelworksheet.Cells[51, 3] = "Files Level";
            excelworksheet.get_Range("c51", "d51").Merge(false);
            excelworksheet.Cells[51, 5] = "Purse";
            excelworksheet.get_Range("e51", "l51").Merge(false);
            excelworksheet.Cells[52, 3] = "Received";
            excelworksheet.Cells[52, 4] = "Failed";
            excelworksheet.Cells[52, 5] = "Add Reverse";
            excelworksheet.get_Range("e52", "f52").Merge(false);
            excelworksheet.Cells[52, 7] = "Usage Reverse";
            excelworksheet.get_Range("g52", "h52").Merge(false);
            excelworksheet.Cells[52, 9] = "Add";
            excelworksheet.get_Range("i52", "j52").Merge(false);
            excelworksheet.Cells[52, 11] = "Usage";
            excelworksheet.get_Range("k52", "l52").Merge(false);
            excelworksheet.Cells[53, 3] = "Vol";
            excelworksheet.Cells[53, 4] = "Vol";
            excelworksheet.Cells[53, 5] = "Vol";
            excelworksheet.Cells[53, 6] = "Val";
            excelworksheet.Cells[53, 7] = "Vol";
            excelworksheet.Cells[53, 8] = "Val";
            excelworksheet.Cells[53, 9] = "Vol";
            excelworksheet.Cells[53, 10] = "Val";
            excelworksheet.Cells[53, 11] = "Vol";
            excelworksheet.Cells[53, 12] = "Val";
            excelworksheet.Cells[54, 2] = "PARAGON FOOD COURT";
            excelworksheet.Cells[54, 3] = prg_num_received_file;
            excelworksheet.Cells[54, 4] = prg_num_failed_file;
            excelworksheet.Cells[54, 5] = FOOD_Purse_Add_Reverse_Vol;
            excelworksheet.Cells[54, 6] = FOOD_Purse_Add_Reverse_Val;
            excelworksheet.Cells[54, 7] = FOOD_Purse_Use_Reverse_Vol;
            excelworksheet.Cells[54, 8] = FOOD_Purse_Use_Reverse_Val;
            excelworksheet.Cells[54, 9] = FOOD_Purse_Add_Vol;
            excelworksheet.Cells[54, 10] = FOOD_Purse_Add_Val;
            excelworksheet.Cells[54, 11] = FOOD_Purse_Use_Vol;
            excelworksheet.Cells[54, 12] = FOOD_Purse_Use_Val;


            //end Paragon food court
            excelworksheet.Cells[56, 2] = "9";
            excelworksheet.get_Range("b56", "b58").Merge(false);
            excelworksheet.Cells[56, 3] = "Purse";
            excelworksheet.get_Range("c56", "d56").Merge(false);
            excelworksheet.Cells[56, 5] = "Transaction Type";
            excelworksheet.get_Range("e56", "g56").Merge(false);
            excelworksheet.Cells[57, 3] = "Add";
            excelworksheet.get_Range("c57", "d57").Merge(false);
            excelworksheet.Cells[57, 5] = "Activate";
            excelworksheet.Cells[57, 6] = "Change of Service";
            excelworksheet.Cells[57, 7] = "Deativate";
            excelworksheet.Cells[58, 3] = "Vol";
            excelworksheet.Cells[58, 4] = "Val";
            excelworksheet.Cells[58, 5] = "Vol";
            excelworksheet.Cells[58, 6] = "Vol";
            excelworksheet.Cells[58, 7] = "Vol";
            excelworksheet.Cells[59, 2] = "Auto Top-up";
            excelworksheet.Cells[59, 3] = BBL_Purse_Add_Vol;
            excelworksheet.Cells[59, 4] = BBL_Purse_Add_Val;
            excelworksheet.Cells[59, 5] = ATU_Product_Autoload_Enable;
            excelworksheet.Cells[59, 6] = ATU_Product_Autoload_Disable;
            excelworksheet.Cells[59, 7]  = ATU_Change_of_Service;
            //end autotopup

            excelworksheet.Cells[61, 2] = "10";
            excelworksheet.get_Range("b61", "b63").Merge(false);
            excelworksheet.Cells[61, 3] = "Purese";
            excelworksheet.get_Range("c61", "h61").Merge(false);
            excelworksheet.Cells[62, 3] = "Usage Reverse";
            excelworksheet.get_Range("c62", "d62").Merge(false);
            excelworksheet.Cells[62, 5] = "Add";
            excelworksheet.get_Range("e62", "f62").Merge(false);
            excelworksheet.Cells[62, 7] = "Usage";
            excelworksheet.get_Range("g62", "h62").Merge(false);
            excelworksheet.Cells[63, 3] = "Vol";
            excelworksheet.Cells[63, 4] = "Val";
            excelworksheet.Cells[63, 5] = "Vol";
            excelworksheet.Cells[63, 6] = "Val";
            excelworksheet.Cells[63, 7] = "Vol";
            excelworksheet.Cells[63, 8] = "Val";
            excelworksheet.Cells[64, 2] = "TESCO EXPRESS";
            excelworksheet.Cells[64, 3] = TESCO_Purse_Use_Reverse_Vol;
            excelworksheet.Cells[64, 4] = TESCO_Purse_Use_Reverse_Val;
            excelworksheet.Cells[64, 5] = TESCO_Purse_Add_Vol;
            excelworksheet.Cells[64, 6] = TESCO_Purse_Add_Val;
            excelworksheet.Cells[64, 7] = TESCO_Purse_Use_Vol;
            excelworksheet.Cells[64, 8] = Math.Round(Convert.ToDecimal(TESCO_Purse_Use_Val));

            //end Tesco express


            excelworksheet.Cells[66, 2] = "11";
            excelworksheet.get_Range("b66", "b68").Merge(false);
            excelworksheet.Cells[66, 3] = "Files Level";
            excelworksheet.get_Range("c66", "d66").Merge(false);
            excelworksheet.Cells[66, 5] = "Purse";
            excelworksheet.get_Range("e66", "l66").Merge(false);
            excelworksheet.Cells[67, 3] = "Received";
            excelworksheet.Cells[67, 4] = "Failed";
            excelworksheet.Cells[67, 5] = "Add Reverse";
            excelworksheet.get_Range("e67", "f67").Merge(false);
            excelworksheet.Cells[67, 7] = "Usage Reverse";
            excelworksheet.get_Range("g67", "h67").Merge(false);
            excelworksheet.Cells[67, 9] = "Add";
            excelworksheet.get_Range("i67", "j67").Merge(false);
            excelworksheet.Cells[67, 11] = "Usage";
            excelworksheet.get_Range("k67", "l67").Merge(false);
            excelworksheet.Cells[68, 3] = "Vol";
            excelworksheet.Cells[68, 4] = "Vol";
            excelworksheet.Cells[68, 5] = "Vol";
            excelworksheet.Cells[68, 6] = "Val";
            excelworksheet.Cells[68, 7] = "Vol";
            excelworksheet.Cells[68, 8] = "Val";
            excelworksheet.Cells[68, 9] = "Vol";
            excelworksheet.Cells[68, 10] = "Val";
            excelworksheet.Cells[68, 11] = "Vol";
            excelworksheet.Cells[68, 12] = "Val";
            excelworksheet.Cells[69, 2] = "EMQUARTIER FOOD COURT";
            excelworksheet.Cells[69, 3] = eqf_num_received_file;
            excelworksheet.Cells[69, 4] = eqf_num_failed_file;
            excelworksheet.Cells[69, 5] = EOF_Purse_Add_Reverse_Vol;
            excelworksheet.Cells[69, 6] = EOF_Purse_Add_Reverse_Val;
            excelworksheet.Cells[69, 7] = EOF_Purse_Use_Reverse_Vol;
            excelworksheet.Cells[69, 8] = EOF_Purse_Use_Reverse_Val;
            excelworksheet.Cells[69, 9] = EOF_Purse_Add_Vol;
            excelworksheet.Cells[69, 10] = EOF_Purse_Add_Val;
            excelworksheet.Cells[69, 11] = EOF_Purse_Use_Vol;
            excelworksheet.Cells[69, 12] = EOF_Purse_Use_Val;


            //end Emquartier

            excelworksheet.Cells[71, 2] = "12";
            excelworksheet.get_Range("b71", "b73").Merge(false);
            excelworksheet.Cells[71, 3] = "Files Level";
            excelworksheet.get_Range("c71", "d71").Merge(false);
            excelworksheet.Cells[71, 5] = "Purse";
            excelworksheet.get_Range("e71", "l71").Merge(false);
            excelworksheet.Cells[72, 3] = "Received";
            excelworksheet.Cells[72, 4] = "Failed";
            excelworksheet.Cells[72, 5] = "Add Reverse";
            excelworksheet.get_Range("e72", "f72").Merge(false);
            excelworksheet.Cells[72, 7] = "Usage Reverse";
            excelworksheet.get_Range("g72", "h72").Merge(false);
            excelworksheet.Cells[72, 9] = "Add";
            excelworksheet.get_Range("i72", "j72").Merge(false);
            excelworksheet.Cells[72, 11] = "Usage";
            excelworksheet.get_Range("k72", "l72").Merge(false);
            excelworksheet.Cells[73, 3] = "Vol";
            excelworksheet.Cells[73, 4] = "Vol";
            excelworksheet.Cells[73, 5] = "Vol";
            excelworksheet.Cells[73, 6] = "Val";
            excelworksheet.Cells[73, 7] = "Vol";
            excelworksheet.Cells[73, 8] = "Val";
            excelworksheet.Cells[73, 9] = "Vol";
            excelworksheet.Cells[73, 10] = "Val";
            excelworksheet.Cells[73, 11] = "Vol";
            excelworksheet.Cells[73, 12] = "Val";
            excelworksheet.Cells[74, 2] = "EMPORIUM FOOD COURT";
            excelworksheet.Cells[74, 3] = epf_num_received_file;
            excelworksheet.Cells[74, 4] = epf_num_failed_file;
            excelworksheet.Cells[74, 5] = EMPORIUM_Purse_Add_Reverse_Vol;
            excelworksheet.Cells[74, 6] = EMPORIUM_Purse_Add_Reverse_Val;
            excelworksheet.Cells[74, 7] = EMPORIUM_Purse_Use_Reverse_Vol;
            excelworksheet.Cells[74, 8] = EMPORIUM_Purse_Use_Reverse_Val;
            excelworksheet.Cells[74, 9] = EMPORIUM_Purse_Add_Vol;
            excelworksheet.Cells[74, 10] = EMPORIUM_Purse_Add_Val;
            excelworksheet.Cells[74, 11] = EMPORIUM_Purse_Use_Vol;
            excelworksheet.Cells[74, 12] = EMPORIUM_Purse_Use_Val;
            // end emporium
            excelworksheet.Cells[76, 2] = "13";
            excelworksheet.get_Range("b76", "b78").Merge(false);
            excelworksheet.Cells[76, 3] = "Files Level";
            excelworksheet.get_Range("c76", "d76").Merge(false);
            excelworksheet.Cells[76, 5] = "Purse";
            excelworksheet.get_Range("e76", "l76").Merge(false);
            excelworksheet.Cells[77, 3] = "Received";
            excelworksheet.Cells[77, 4] = "Failed";
            excelworksheet.Cells[77, 5] = "Add Reverse";
            excelworksheet.get_Range("e77", "f77").Merge(false);
            excelworksheet.Cells[77, 7] = "Usage Reverse";
            excelworksheet.get_Range("g77", "h77").Merge(false);
            excelworksheet.Cells[77, 9] = "Add";
            excelworksheet.get_Range("i77", "j77").Merge(false);
            excelworksheet.Cells[77, 11] = "Usage";
            excelworksheet.get_Range("k77", "l77").Merge(false);
            excelworksheet.Cells[78, 3] = "Vol";
            excelworksheet.Cells[78, 4] = "Vol";
            excelworksheet.Cells[78, 5] = "Vol";
            excelworksheet.Cells[78, 6] = "Val";
            excelworksheet.Cells[78, 7] = "Vol";
            excelworksheet.Cells[78, 8] = "Val";
            excelworksheet.Cells[78, 9] = "Vol";
            excelworksheet.Cells[78, 10] = "Val";
            excelworksheet.Cells[78, 11] = "Vol";
            excelworksheet.Cells[78, 12] = "Val";
            excelworksheet.Cells[79, 2] = "MBK FOOD COURT";
            excelworksheet.Cells[79, 3] = mbf_num_received_file;
            excelworksheet.Cells[79, 4] = mbf_num_failed_file;
            excelworksheet.Cells[79, 5] = MBK_Purse_Add_Reverse_Vol;
            excelworksheet.Cells[79, 6] = MBK_Purse_Add_Reverse_Val;
            excelworksheet.Cells[79, 7] = MBK_Purse_Use_Reverse_Vol;
            excelworksheet.Cells[79, 8] = MBK_Purse_Use_Reverse_Val;
            excelworksheet.Cells[79, 9] = MBK_Purse_Add_Vol;
            excelworksheet.Cells[79, 10] = MBK_Purse_Add_Val;
            excelworksheet.Cells[79, 11] = MBK_Purse_Use_Vol;
            excelworksheet.Cells[79, 12] = MBK_Purse_Use_Val;



            //end MBK

            excelworksheet.Cells[81, 2] = "14";
            excelworksheet.get_Range("b81", "b83").Merge(false);
            excelworksheet.Cells[81, 3] = "Files Level";
            excelworksheet.get_Range("c81", "d81").Merge(false);
            excelworksheet.Cells[81, 5] = "Purse";
            excelworksheet.get_Range("e81", "l81").Merge(false);
            excelworksheet.Cells[82, 3] = "Received";
            excelworksheet.Cells[82, 4] = "Failed";
            excelworksheet.Cells[82, 5] = "Add Reverse";
            excelworksheet.get_Range("e82", "f82").Merge(false);
            excelworksheet.Cells[82, 7] = "Usage Reverse";
            excelworksheet.get_Range("g82", "h82").Merge(false);
            excelworksheet.Cells[82, 9] = "Add";
            excelworksheet.get_Range("i82", "j82").Merge(false);
            excelworksheet.Cells[82, 11] = "Usage";
            excelworksheet.get_Range("k82", "l82").Merge(false);
            excelworksheet.Cells[83, 3] = "Vol";
            excelworksheet.Cells[83, 4] = "Vol";
            excelworksheet.Cells[83, 5] = "Vol";
            excelworksheet.Cells[83, 6] = "Val";
            excelworksheet.Cells[83, 7] = "Vol";
            excelworksheet.Cells[83, 8] = "Val";
            excelworksheet.Cells[83, 9] = "Vol";
            excelworksheet.Cells[83, 10] = "Val";
            excelworksheet.Cells[83, 11] = "Vol";
            excelworksheet.Cells[83, 12] = "Val";
            excelworksheet.Cells[84, 2] = "FOOD STREET FOOD COURT";
            excelworksheet.Cells[84, 3] = FSF_NUM_RECEIVED_FILE;
            excelworksheet.Cells[84, 4] = FSF_NUM_FAILED_FILE;
            excelworksheet.Cells[84, 5] = FOOD_STREET_Purse_Add_Reverse_Vol;
            excelworksheet.Cells[84, 6] = FOOD_STREET_Purse_Add_Reverse_Val;
            excelworksheet.Cells[84, 7] = FOOD_STREET_Purse_Use_Reverse_Vol;
            excelworksheet.Cells[84, 8] = FOOD_STREET_Purse_Use_Reverse_Val;
            excelworksheet.Cells[84, 9] = FOOD_STREET_Purse_Add_Vol;
            excelworksheet.Cells[84, 10] = FOOD_STREET_Purse_Add_Val;
            excelworksheet.Cells[84, 11] = FOOD_STREET_Purse_Use_Vol;
            excelworksheet.Cells[84, 12] = FOOD_STREET_Purse_Use_Val;

            // end FOOD STREET FOOD COURT


            excelworksheet.Cells[86, 2] = "15";
            excelworksheet.get_Range("b86", "b88").Merge(false);
            excelworksheet.Cells[86, 3] = "Files Level";
            excelworksheet.get_Range("c86", "d86").Merge(false);
            excelworksheet.Cells[86, 5] = "Purse";
            excelworksheet.get_Range("e86", "l86").Merge(false);
            excelworksheet.Cells[87, 3] = "Received";
            excelworksheet.Cells[87, 4] = "Failed";
            excelworksheet.Cells[87, 5] = "Add Reverse";
            excelworksheet.get_Range("e87", "f87").Merge(false);
            excelworksheet.Cells[87, 7] = "Usage Reverse";
            excelworksheet.get_Range("g87", "h87").Merge(false);
            excelworksheet.Cells[87, 9] = "Add";
            excelworksheet.get_Range("i87", "j87").Merge(false);
            excelworksheet.Cells[87, 11] = "Usage";
            excelworksheet.get_Range("k87", "l87").Merge(false);
            excelworksheet.Cells[88, 3] = "Vol";
            excelworksheet.Cells[88, 4] = "Vol";
            excelworksheet.Cells[88, 5] = "Vol";
            excelworksheet.Cells[88, 6] = "Val";
            excelworksheet.Cells[88, 7] = "Vol";
            excelworksheet.Cells[88, 8] = "Val";
            excelworksheet.Cells[88, 9] = "Vol";
            excelworksheet.Cells[88, 10] = "Val";
            excelworksheet.Cells[88, 11] = "Vol";
            excelworksheet.Cells[88, 12] = "Val";
            excelworksheet.Cells[89, 2] = "THE MALL THAPRA FOOD COURT";
            excelworksheet.Cells[89, 3] = mtf_num_received_file;
            excelworksheet.Cells[89, 4] = mtf_num_failed_file;
            excelworksheet.Cells[89, 5] = THAPRA_Purse_Add_Reverse_Vol;
            excelworksheet.Cells[89, 6] = THAPRA_Purse_Add_Reverse_Val;
            excelworksheet.Cells[89, 7] = THAPRA_Purse_Use_Reverse_Vol;
            excelworksheet.Cells[89, 8] = THAPRA_Purse_Use_Reverse_Val;
            excelworksheet.Cells[89, 9] = THAPRA_Purse_Add_Vol;
            excelworksheet.Cells[89, 10] = THAPRA_Purse_Add_Val;
            excelworksheet.Cells[89, 11] = THAPRA_Purse_Use_Vol;
            excelworksheet.Cells[89, 12] = THAPRA_Purse_Use_Val;



            //end THAPRA FOOD COURT

            excelworksheet.Cells[91, 2] = "16";
            excelworksheet.get_Range("b91", "b93").Merge(false);
            excelworksheet.Cells[91, 3] = "Files Level";
            excelworksheet.get_Range("c91", "d91").Merge(false);
            excelworksheet.Cells[91, 5] = "Purse";
            excelworksheet.get_Range("e91", "l91").Merge(false);
            excelworksheet.Cells[92, 3] = "Received";
            excelworksheet.Cells[92, 4] = "Failed";
            excelworksheet.Cells[92, 5] = "Add Reverse";
            excelworksheet.get_Range("e92", "f92").Merge(false);
            excelworksheet.Cells[92, 7] = "Usage Reverse";
            excelworksheet.get_Range("g92", "h92").Merge(false);
            excelworksheet.Cells[92, 9] = "Add";
            excelworksheet.get_Range("i92", "j92").Merge(false);
            excelworksheet.Cells[92, 11] = "Usage";
            excelworksheet.get_Range("k92", "l92").Merge(false);
            excelworksheet.Cells[93, 3] = "Vol";
            excelworksheet.Cells[93, 4] = "Vol";
            excelworksheet.Cells[93, 5] = "Vol";
            excelworksheet.Cells[93, 6] = "Val";
            excelworksheet.Cells[93, 7] = "Vol";
            excelworksheet.Cells[93, 8] = "Val";
            excelworksheet.Cells[93, 9] = "Vol";
            excelworksheet.Cells[93, 10] = "Val";
            excelworksheet.Cells[93, 11] = "Vol";
            excelworksheet.Cells[93, 12] = "Val";
            excelworksheet.Cells[94, 2] = "THE MALL BANGKAE FOOD COURT";
            excelworksheet.Cells[94, 3] = bkf_num_received_file;
            excelworksheet.Cells[94, 4] = bkf_num_failed_file;
            excelworksheet.Cells[94, 5] = BANGKAE_Purse_Add_Reverse_Vol;
            excelworksheet.Cells[94, 6] = BANGKAE_Purse_Add_Reverse_Val;
            excelworksheet.Cells[94, 7] = BANGKAE_Purse_Use_Reverse_Vol;
            excelworksheet.Cells[94, 8] = BANGKAE_Purse_Use_Reverse_Val;
            excelworksheet.Cells[94, 9] = BANGKAE_Purse_Add_Vol;
            excelworksheet.Cells[94, 10] = BANGKAE_Purse_Add_Val;
            excelworksheet.Cells[94, 11] = BANGKAE_Purse_Use_Vol;
            excelworksheet.Cells[94, 12] = BANGKAE_Purse_Use_Val;
            //17

            excelworksheet.Cells[96, 2] = "17";
            excelworksheet.get_Range("b96", "b98").Merge(false);
            excelworksheet.Cells[96, 3] = "Files Level";
            excelworksheet.get_Range("c96", "d96").Merge(false);
            excelworksheet.Cells[96, 5] = "Purse";
            excelworksheet.get_Range("e96", "l96").Merge(false);
            excelworksheet.Cells[97, 3] = "Received";
            excelworksheet.Cells[97, 4] = "Failed";
            excelworksheet.Cells[97, 5] = "Add Reverse";
            excelworksheet.get_Range("e97", "f97").Merge(false);
            excelworksheet.Cells[97, 7] = "Usage Reverse";
            excelworksheet.get_Range("g97", "h97").Merge(false);
            excelworksheet.Cells[97, 9] = "Add";
            excelworksheet.get_Range("i97", "j97").Merge(false);
            excelworksheet.Cells[97, 11] = "Usage";
            excelworksheet.get_Range("k97", "l97").Merge(false);
            excelworksheet.Cells[98, 3] = "Vol";
            excelworksheet.Cells[98, 4] = "Vol";
            excelworksheet.Cells[98, 5] = "Vol";
            excelworksheet.Cells[98, 6] = "Val";
            excelworksheet.Cells[98, 7] = "Vol";
            excelworksheet.Cells[98, 8] = "Val";
            excelworksheet.Cells[98, 9] = "Vol";
            excelworksheet.Cells[98, 10] = "Val";
            excelworksheet.Cells[98, 11] = "Vol";
            excelworksheet.Cells[98, 12] = "Val";
            excelworksheet.Cells[99, 2] = "FOOD STREET EKAMAI FOOD COURT";
            excelworksheet.Cells[99, 3] = fef_num_received_file;
            excelworksheet.Cells[99, 4] = fef_num_failed_file;
            excelworksheet.Cells[99, 5] = FSEKAMAI_Purse_Add_Reverse_Vol;
            excelworksheet.Cells[99, 6] = FSEKAMAI_Purse_Add_Reverse_Val;
            excelworksheet.Cells[99, 7] = FSEKAMAI_Purse_Use_Reverse_Vol;
            excelworksheet.Cells[99, 8] = FSEKAMAI_Purse_Use_Reverse_Val;
            excelworksheet.Cells[99, 9] = FSEKAMAI_Purse_Add_Vol;
            excelworksheet.Cells[99, 10] = FSEKAMAI_Purse_Add_Val;
            excelworksheet.Cells[99, 11] = FSEKAMAI_Purse_Use_Vol;
            excelworksheet.Cells[99, 12] = FSEKAMAI_Purse_Use_Val;
            //18
            excelworksheet.Cells[101, 2] = "18";
            excelworksheet.get_Range("b101", "b103").Merge(false);
            excelworksheet.Cells[101, 3] = "Files Level";
            excelworksheet.get_Range("c101", "d101").Merge(false);
            excelworksheet.Cells[101, 5] = "Purse";
            excelworksheet.get_Range("e101", "l101").Merge(false);
            excelworksheet.Cells[102, 3] = "Received";
            excelworksheet.Cells[102, 4] = "Failed";
            excelworksheet.Cells[102, 5] = "Add Reverse";
            excelworksheet.get_Range("e102", "f102").Merge(false);
            excelworksheet.Cells[102, 7] = "Usage Reverse";
            excelworksheet.get_Range("g102", "h102").Merge(false);
            excelworksheet.Cells[102, 9] = "Add";
            excelworksheet.get_Range("i102", "j102").Merge(false);
            excelworksheet.Cells[102, 11] = "Usage";
            excelworksheet.get_Range("k102", "l102").Merge(false);
            excelworksheet.Cells[103, 3] = "Vol";
            excelworksheet.Cells[103, 4] = "Vol";
            excelworksheet.Cells[103, 5] = "Vol";
            excelworksheet.Cells[103, 6] = "Val";
            excelworksheet.Cells[103, 7] = "Vol";
            excelworksheet.Cells[103, 8] = "Val";
            excelworksheet.Cells[103, 9] = "Vol";
            excelworksheet.Cells[103, 10] = "Val";
            excelworksheet.Cells[103, 11] = "Vol";
            excelworksheet.Cells[103, 12] = "Val";
            excelworksheet.Cells[104, 2] = "SET FOOD COURT";
            excelworksheet.Cells[104, 3] = sef_num_received_file;
            excelworksheet.Cells[104, 4] = sef_num_failed_file;
            excelworksheet.Cells[104, 5] = SET_Purse_Add_Reverse_Vol;
            excelworksheet.Cells[104, 6] = SET_Purse_Add_Reverse_Val;
            excelworksheet.Cells[104, 7] = (Convert.ToDouble(SETLOC1_Purse_Use_Reverse_Vol) + Convert.ToDouble(SETLOC2_Purse_Use_Reverse_Vol) + Convert.ToDouble(SETLOC3_Purse_Use_Reverse_Vol) + Convert.ToDouble(SETLOC4_Purse_Use_Reverse_Vol) + Convert.ToDouble(SETLOC5_Purse_Use_Reverse_Vol) + Convert.ToDouble(SETLOC6_Purse_Use_Reverse_Vol) + Convert.ToDouble(SETLOC7_Purse_Use_Reverse_Vol)).ToString("0");
            excelworksheet.Cells[104, 8] = (Convert.ToDouble(SETLOC1_Purse_Use_Reverse_Val) + Convert.ToDouble(SETLOC2_Purse_Use_Reverse_Val) + Convert.ToDouble(SETLOC3_Purse_Use_Reverse_Val) + Convert.ToDouble(SETLOC4_Purse_Use_Reverse_Val) + Convert.ToDouble(SETLOC5_Purse_Use_Reverse_Val) + Convert.ToDouble(SETLOC6_Purse_Use_Reverse_Val) + Convert.ToDouble(SETLOC7_Purse_Use_Reverse_Val)).ToString("0");
            excelworksheet.Cells[104, 9] = SET_Purse_Add_Vol;
            excelworksheet.Cells[104, 10] = SET_Purse_Add_Val;
            excelworksheet.Cells[104, 11] = (Convert.ToDouble(SETLOC1_Purse_Use_Vol) + Convert.ToDouble(SETLOC2_Purse_Use_Vol) + Convert.ToDouble(SETLOC3_Purse_Use_Vol) + Convert.ToDouble(SETLOC4_Purse_Use_Vol) + Convert.ToDouble(SETLOC5_Purse_Use_Vol) + Convert.ToDouble(SETLOC6_Purse_Use_Vol) + Convert.ToDouble(SETLOC7_Purse_Use_Vol)).ToString("0");
            excelworksheet.Cells[104, 12] = (Convert.ToDouble(SETLOC1_Purse_Use_Val) + Convert.ToDouble(SETLOC2_Purse_Use_Val) + Convert.ToDouble(SETLOC3_Purse_Use_Val) + Convert.ToDouble(SETLOC4_Purse_Use_Val) + Convert.ToDouble(SETLOC5_Purse_Use_Val) + Convert.ToDouble(SETLOC6_Purse_Use_Val) + Convert.ToDouble(SETLOC7_Purse_Use_Val)).ToString("0");
            //19
            excelworksheet.Cells[106, 2] = "19";
            excelworksheet.get_Range("b106", "b108").Merge(false);
            excelworksheet.Cells[106, 3] = "Files Level";
            excelworksheet.get_Range("c106", "d106").Merge(false);
            excelworksheet.Cells[106, 5] = "Purse";
            excelworksheet.get_Range("e106", "l106").Merge(false);
            excelworksheet.Cells[107, 3] = "Received";
            excelworksheet.Cells[107, 4] = "Failed";
            excelworksheet.Cells[107, 5] = "Add Reverse";
            excelworksheet.get_Range("e107", "f107").Merge(false);
            excelworksheet.Cells[107, 7] = "Usage Reverse";
            excelworksheet.get_Range("g107", "h107").Merge(false);
            excelworksheet.Cells[107, 9] = "Add";
            excelworksheet.get_Range("i107", "j107").Merge(false);
            excelworksheet.Cells[107, 11] = "Usage";
            excelworksheet.get_Range("k107", "l107").Merge(false);
            excelworksheet.Cells[108, 3] = "Vol";
            excelworksheet.Cells[108, 4] = "Vol";
            excelworksheet.Cells[108, 5] = "Vol";
            excelworksheet.Cells[108, 6] = "Val";
            excelworksheet.Cells[108, 7] = "Vol";
            excelworksheet.Cells[108, 8] = "Val";
            excelworksheet.Cells[108, 9] = "Vol";
            excelworksheet.Cells[108, 10] = "Val";
            excelworksheet.Cells[108, 11] = "Vol";
            excelworksheet.Cells[108, 12] = "Val";
            excelworksheet.Cells[109, 2] = "IMPERIAL SAMRONG  FOOD COURT";
            excelworksheet.Cells[109, 3] = ips_num_received_file;
            excelworksheet.Cells[109, 4] = ips_num_failed_file;
            excelworksheet.Cells[109, 5] = (Convert.ToDouble(IMPERIAL4F_Purse_Add_Reverse_Vol) + Convert.ToDouble(IMPERIALBF_Purse_Add_Reverse_Vol)).ToString("0");
            excelworksheet.Cells[109, 6] = (Convert.ToDouble(IMPERIAL4F_Purse_Add_Reverse_Val) + Convert.ToDouble(IMPERIALBF_Purse_Add_Reverse_Val)).ToString("0");
            excelworksheet.Cells[109, 7] = (Convert.ToDouble(IMPERIAL4F_Purse_Use_Reverse_Vol) + Convert.ToDouble(IMPERIALBF_Purse_Use_Reverse_Vol)).ToString("0");
            excelworksheet.Cells[109, 8] = (Convert.ToDouble(IMPERIAL4F_Purse_Use_Reverse_Val) + Convert.ToDouble(IMPERIALBF_Purse_Use_Reverse_Val)).ToString("0");
            excelworksheet.Cells[109, 9] = (Convert.ToDouble(IMPERIAL4F_Purse_Add_Vol) + Convert.ToDouble(IMPERIALBF_Purse_Add_Vol)).ToString("0");
            excelworksheet.Cells[109, 10] = (Convert.ToDouble(IMPERIAL4F_Purse_Add_Val) + Convert.ToDouble(IMPERIALBF_Purse_Add_Val)).ToString("0");
            excelworksheet.Cells[109, 11] = (Convert.ToDouble(IMPERIAL4F_Purse_Use_Vol) + Convert.ToDouble(IMPERIALBF_Purse_Use_Vol)).ToString("0");
            excelworksheet.Cells[109, 12] = (Convert.ToDouble(IMPERIAL4F_Purse_Use_Val) + Convert.ToDouble(IMPERIALBF_Purse_Use_Val)).ToString("0");
            //20
            excelworksheet.Cells[111, 2] = "20";
            excelworksheet.get_Range("b111", "b113").Merge(false);
            excelworksheet.Cells[111, 3] = "Files Level";
            excelworksheet.get_Range("c111", "d111").Merge(false);
            excelworksheet.Cells[111, 5] = "Purse";
            excelworksheet.get_Range("e111", "l111").Merge(false);
            excelworksheet.Cells[112, 3] = "Received";
            excelworksheet.Cells[112, 4] = "Failed";
            excelworksheet.Cells[112, 5] = "Add Reverse";
            excelworksheet.get_Range("e112", "f112").Merge(false);
            excelworksheet.Cells[112, 7] = "Usage Reverse";
            excelworksheet.get_Range("g112", "h112").Merge(false);
            excelworksheet.Cells[112, 9] = "Add";
            excelworksheet.get_Range("i112", "j112").Merge(false);
            excelworksheet.Cells[112, 11] = "Usage";
            excelworksheet.get_Range("k112", "l112").Merge(false);
            excelworksheet.Cells[113, 3] = "Vol";
            excelworksheet.Cells[113, 4] = "Vol";
            excelworksheet.Cells[113, 5] = "Vol";
            excelworksheet.Cells[113, 6] = "Val";
            excelworksheet.Cells[113, 7] = "Vol";
            excelworksheet.Cells[113, 8] = "Val";
            excelworksheet.Cells[113, 9] = "Vol";
            excelworksheet.Cells[113, 10] = "Val";
            excelworksheet.Cells[113, 11] = "Vol";
            excelworksheet.Cells[113, 12] = "Val";
            excelworksheet.Cells[114, 2] = "THE MALL NGAMWONGWAN FOOD COURT";
            excelworksheet.Cells[114, 3] = mwf_num_received_file;
            excelworksheet.Cells[114, 4] = mwf_num_failed_file;
            excelworksheet.Cells[114, 5] = (Convert.ToDouble(MallNgamwongwan_Purse_Add_Reverse_Vol) + Convert.ToDouble(MallNgamwongwan_Purse_Add_Reverse_Vol)).ToString("0");
            excelworksheet.Cells[114, 6] = (Convert.ToDouble(MallNgamwongwan_Purse_Add_Reverse_Val) + Convert.ToDouble(MallNgamwongwan_Purse_Add_Reverse_Val)).ToString("0");
            excelworksheet.Cells[114, 7] = (Convert.ToDouble(MallNgamwongwan_Purse_Use_Reverse_Vol) + Convert.ToDouble(MallNgamwongwan_Purse_Use_Reverse_Vol)).ToString("0");
            excelworksheet.Cells[114, 8] = (Convert.ToDouble(MallNgamwongwan_Purse_Use_Reverse_Val) + Convert.ToDouble(MallNgamwongwan_Purse_Use_Reverse_Val)).ToString("0");
            excelworksheet.Cells[114, 9] = (Convert.ToDouble(MallNgamwongwan_Purse_Add_Vol) + Convert.ToDouble(MallNgamwongwan_Purse_Add_Vol)).ToString("0");
            excelworksheet.Cells[114, 10] = (Convert.ToDouble(MallNgamwongwan_Purse_Add_Val) + Convert.ToDouble(MallNgamwongwan_Purse_Add_Val)).ToString("0");
            excelworksheet.Cells[114, 11] = (Convert.ToDouble(MallNgamwongwan_Purse_Use_Vol) + Convert.ToDouble(MallNgamwongwan_Purse_Use_Vol)).ToString("0");
            excelworksheet.Cells[114, 12] = (Convert.ToDouble(MallNgamwongwan_Purse_Use_Val) + Convert.ToDouble(MallNgamwongwan_Purse_Use_Val)).ToString("0");
            //21
            excelworksheet.Cells[116, 2] = "21";
            excelworksheet.get_Range("b116", "b118").Merge(false);
            excelworksheet.Cells[116, 3] = "Files Level";
            excelworksheet.get_Range("c116", "d116").Merge(false);
            excelworksheet.Cells[116, 5] = "Purse";
            excelworksheet.get_Range("e116", "l116").Merge(false);
            excelworksheet.Cells[117, 3] = "Received";
            excelworksheet.Cells[117, 4] = "Failed";
            excelworksheet.Cells[117, 5] = "Add Reverse";
            excelworksheet.get_Range("e117", "f117").Merge(false);
            excelworksheet.Cells[117, 7] = "Usage Reverse";
            excelworksheet.get_Range("g117", "h117").Merge(false);
            excelworksheet.Cells[117, 9] = "Add";
            excelworksheet.get_Range("i117", "j117").Merge(false);
            excelworksheet.Cells[117, 11] = "Usage";
            excelworksheet.get_Range("k117", "l117").Merge(false);
            excelworksheet.Cells[118, 3] = "Vol";
            excelworksheet.Cells[118, 4] = "Vol";
            excelworksheet.Cells[118, 5] = "Vol";
            excelworksheet.Cells[118, 6] = "Val";
            excelworksheet.Cells[118, 7] = "Vol";
            excelworksheet.Cells[118, 8] = "Val";
            excelworksheet.Cells[118, 9] = "Vol";
            excelworksheet.Cells[118, 10] = "Val";
            excelworksheet.Cells[118, 11] = "Vol";
            excelworksheet.Cells[118, 12] = "Val";
            excelworksheet.Cells[119, 2] = "FS CW FOOD COURT";
            excelworksheet.Cells[119, 3] = cwf_num_received_file;
            excelworksheet.Cells[119, 4] = cwf_num_failed_file;
            excelworksheet.Cells[119, 5] = (Convert.ToDouble(CYBERWORLD_Purse_Add_Reverse_Vol) + Convert.ToDouble(CYBERWORLD_Purse_Add_Reverse_Vol)).ToString("0");
            excelworksheet.Cells[119, 6] = (Convert.ToDouble(CYBERWORLD_Purse_Add_Reverse_Val) + Convert.ToDouble(CYBERWORLD_Purse_Add_Reverse_Val)).ToString("0");
            excelworksheet.Cells[119, 7] = (Convert.ToDouble(CYBERWORLD_Purse_Use_Reverse_Vol) + Convert.ToDouble(CYBERWORLD_Purse_Use_Reverse_Vol)).ToString("0");
            excelworksheet.Cells[119, 8] = (Convert.ToDouble(CYBERWORLD_Purse_Use_Reverse_Val) + Convert.ToDouble(CYBERWORLD_Purse_Use_Reverse_Val)).ToString("0");
            excelworksheet.Cells[119, 9] = (Convert.ToDouble(CYBERWORLD_Purse_Add_Vol) + Convert.ToDouble(CYBERWORLD_Purse_Add_Vol)).ToString("0");
            excelworksheet.Cells[119, 10] = (Convert.ToDouble(CYBERWORLD_Purse_Add_Val) + Convert.ToDouble(CYBERWORLD_Purse_Add_Val)).ToString("0");
            excelworksheet.Cells[119, 11] = (Convert.ToDouble(CYBERWORLD_Purse_Use_Vol) + Convert.ToDouble(CYBERWORLD_Purse_Use_Vol)).ToString("0");
            excelworksheet.Cells[119, 12] = (Convert.ToDouble(CYBERWORLD_Purse_Use_Val) + Convert.ToDouble(CYBERWORLD_Purse_Use_Val)).ToString("0");
            //22
            excelworksheet.Cells[121, 2] = "22";
            excelworksheet.get_Range("b121", "b123").Merge(false);
            excelworksheet.Cells[121, 3] = "Purse";
            excelworksheet.get_Range("c121", "f121").Merge(false);
            excelworksheet.Cells[122, 3] = "Add Reverse";
            excelworksheet.get_Range("c122", "d122").Merge(false);
            excelworksheet.Cells[122, 5] = "Add";
            excelworksheet.get_Range("e122", "f122").Merge(false);
            excelworksheet.Cells[123, 3] = "Vol";
            excelworksheet.Cells[123, 4] = "Val";
            excelworksheet.Cells[123, 5] = "Vol";
            excelworksheet.Cells[123, 6] = "Val";
            excelworksheet.Cells[124, 2] = "LEGACY";
            excelworksheet.Cells[124, 3] = LEGACY_Purse_Add_Reverse_Vol;
            excelworksheet.Cells[124, 4] = LEGACY_Purse_Add_Reverse_Val;
            excelworksheet.Cells[124, 5] = LEGACY_Purse_Add_Vol;
            excelworksheet.Cells[124, 6] = LEGACY_Purse_Add_Val;

            //

            excelworksheet.Cells[126, 1] = "CCH Exception";
            excelworksheet.Cells[127, 1] = "(cchcpt)";
            excelworksheet.Cells[126, 4] = "Vol";
            excelworksheet.Cells[127, 2] = "ITG (Integrity Value Check)";
            excelworksheet.Cells[128, 2] = "MAC (Message Authentication Code)";
            excelworksheet.Cells[129, 2] = "DDT (Duplicate Device Transaction)";
            excelworksheet.Cells[130, 2] = "EXP (Expired Transaction)";
            excelworksheet.Cells[131, 2] = "IND (Invalid Device)";
            excelworksheet.Cells[132, 2] = "IVD (Invalid Date)";
            excelworksheet.Cells[133, 2] = "UNC (Unconfirmed Transaction)";
            excelworksheet.get_Range("b126", "c126").Merge(false);
            excelworksheet.get_Range("b127", "c127").Merge(false);
            excelworksheet.get_Range("b128", "c128").Merge(false);
            excelworksheet.get_Range("b129", "c129").Merge(false);
            excelworksheet.get_Range("b130", "c130").Merge(false);
            excelworksheet.get_Range("b131", "c131").Merge(false);
            excelworksheet.get_Range("b132", "c132").Merge(false);
            excelworksheet.get_Range("b133", "c133").Merge(false);
            excelworksheet.Cells[127, 4] = exc_itg;
            excelworksheet.Cells[128, 4] = exc_mac;
            excelworksheet.Cells[129, 4] = exc_ddt;
            excelworksheet.Cells[130, 4] = exc_exp;
            excelworksheet.Cells[131, 4] = exc_ind;
            excelworksheet.Cells[132, 4] = exc_ivd;
            excelworksheet.Cells[133, 4] = exc_unc;
            //
            excelworksheet.Cells[126, 6] = "ISS Exception";
            excelworksheet.Cells[127, 6] = "(isscpt)";
            excelworksheet.get_Range("f126", "g126").Merge(false);
            excelworksheet.get_Range("f127", "g127").Merge(false);
            excelworksheet.Cells[126, 12] = "Vol";
            excelworksheet.Cells[127, 8] = "IRV(Incorrect Remaining Value)";
            excelworksheet.Cells[128, 8] = "XTX(Expired Transaction Exception)";
            excelworksheet.Cells[129, 8] = "CMR(Card Master Record Not Present)";
            excelworksheet.Cells[130, 8] = "PMR(Product Master Record Not Present)";
            excelworksheet.Cells[131, 8] = "PSN(Duplicate Product Transaction Seq.No)";
            excelworksheet.get_Range("h126", "k126").Merge(false);
            excelworksheet.get_Range("h127", "k127").Merge(false);
            excelworksheet.get_Range("h128", "k128").Merge(false);
            excelworksheet.get_Range("h129", "k129").Merge(false);
            excelworksheet.get_Range("h130", "k130").Merge(false);
            excelworksheet.get_Range("h131", "k131").Merge(false);
            excelworksheet.Cells[127, 12] = exc_irv;
            excelworksheet.Cells[128, 12] = exc_xtx;
            excelworksheet.Cells[129, 12] = exc_cmr;
            excelworksheet.Cells[130, 12] = exc_pmr;
            excelworksheet.Cells[131, 12] = exc_psn;

            

    /*         StringBuilder vars = new StringBuilder();
            vars.AppendLine("<table  style=\"border-collapse:collapse;text-align:center;\" width=\"100%\">");
            for (int i = 0; i < 114; i++) {
                switch (i) {
                    case 0:
                        vars.AppendLine("<tr><td width=\"5.6%\"></td><td width=\"8.6%\"></td><td width=\"2.8%\"></td><td width=\"2.8%\"></td><td width=\"2.8%\"></td><td width=\"2.8%\"></td><td width=\"2.8%\"></td><td width=\"2.8%\"></td><td width=\"2.8%\"></td><td width=\"2.8%\"></td><td width=\"2.8%\"></td><td width=\"2.8%\"></td><td width=\"2.8%\"></td><td width=\"2.8%\"></td><td width=\"2.8%\"></td><td width=\"2.8%\"></td><td width=\"2.8%\"></td><td width=\"2.8%\"></td><td width=\"2.8%\"></td><td width=\"2.8%\"></td><td width=\"2.8%\"></td><td width=\"2.8%\"></td><td width=\"2.8%\"></td><td width=\"2.8%\"></td><td width=\"2.8%\"></td><td width=\"2.8%\"></td><td width=\"2.8%\"></td><td width=\"2.8%\"></td><td width=\"2.8%\"></td><td width=\"2.8%\"></td><td width=\"2.8%\"></td><td width=\"2.8%\"></td><td width=\"2.8%\"></td></tr>");
                        break;
                    case 1:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j < 34; j++)
                        {
                            string dummy = "";
                            if (j == 1) {
                                dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td colspan=\"3\"><h2>" + dummy + "</h2></td>");
                     
                            }
                           
                            if (j>3)
                            {
                                dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td>" + dummy + "</td>");
                            }
                        }
                        vars.AppendLine("</tr></p>");
                        break;
                    case 3:
                    case 6:
                  
                        vars.AppendLine("<tr>");
                        for (int j = 1; j < 34; j++)
                        {
                            if (j == 3)
                            {
                                var dummy = Convert.ToDateTime(excelworksheet.Cells[i, j].Value,CultureInfo.CreateSpecificCulture("en-US")).ToString("dd/MM/yyyy", CultureInfo.CreateSpecificCulture("en-US"));
                                vars.AppendLine("<td>" + dummy + "</td>");
                            }
                            else
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td>" + dummy + "</td>");
                            }
                           
                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 11:
                        vars.AppendLine("<tr><td></td></tr>");
                        break;
                    case 8:
                    case 9:
                    case 10:
                    case 12:
                    case 13:
                    case 14:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 3; j++)
                        {
                           
                            if (j == 1)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td>" + dummy + "</td>");
                            }
                            if (j == 2)
                            {
                              
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td  style = \"border:1px solid; background-color: #BDD7EE;\" >" + dummy + "</td>");
                            }
                            if (j == 3)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                
                                vars.AppendLine("<td  style =\"border:1px solid;text-align: right;\">" + dummy + "</td>");
                            }
                      
                        }
                        vars.AppendLine("</tr>");


                        break;
                    case 16:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 33; j++)
                        {
                          if (j == 1)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td>" + dummy + "</td>");
                            }
                            if (j == 2)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\" rowspan=\"3\">" + dummy + "</td>");
                            }
                            if (j == 3)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td  style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"2\">" + dummy + "</td>");
                            }

                            if (j == 5)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td  style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"5\">" + dummy + "</td>");
                            }
                            if (j == 10)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td  style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"13\">" + dummy + "</td>");
                            }
                            if (j == 23)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td  style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"2\">" + dummy + "</td>");
                            }
                            if (j == 25)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td  style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"9\">" + dummy + "</td>");
                            }

                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 17:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 33; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if ((j >2&&j<10)||j==23||j==24||j==27||j==12)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }
                            if (j ==10||j==13||j==15||j==17 || j == 19 || j == 21 || j == 25||j==28 || j == 30 || j == 32)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"2\">" + dummy + "</td>");
                            }

                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 18:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 33; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if (j > 2)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }

                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 19:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 33; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if (j == 2) {
                                var dummy = excelworksheet.Cells[i, j].Value;
                            vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }
                            if (j > 2)
                            {
                                var dummy = Convert.ToDecimal(excelworksheet.Cells[i, j].Value).ToString("#,##0");
                                vars.AppendLine("<td style =\"border:1px solid;text-align: right;\">" + dummy + "</td>");
                            }


                        }
                        vars.AppendLine("</tr><tr><td></td></tr>");
                        break;
//20
                    case 21:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 33; j++)
                        {
                            if (j == 1)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td>" + dummy + "</td>");
                            }
                            if (j == 2)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\" rowspan=\"3\">" + dummy + "</td>");
                            }
                            if (j == 3)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td  style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"2\">" + dummy + "</td>");
                            }

                            if (j == 5)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td  style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"2\">" + dummy + "</td>");
                            }
                            if (j == 7)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td  style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"10\">" + dummy + "</td>");
                            }
                            if (j == 17)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td  style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"2\">" + dummy + "</td>");
                            }
                     

                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 22:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 18; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if ((j > 2 && j < 7)||j==17||j==18)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }
                            if (j == 7 || j == 9 || j == 11 || j == 13||j==15)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"2\">" + dummy + "</td>");
                            }

                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 23:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 18; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if (j > 2)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }

                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 24:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 18; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if (j == 2)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }
                            if (j > 2)
                            {
                                var dummy = Convert.ToDecimal(excelworksheet.Cells[i, j].Value).ToString("#,##0");
                                vars.AppendLine("<td style =\"border:1px solid;text-align: right;\">" + dummy + "</td>");
                            }


                        }
                        vars.AppendLine("</tr><tr><td></td></tr>");
                        break;
                    ///24
                    case 26:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 7; j++)
                        {
                            if (j == 1)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td>" + dummy + "</td>");
                            }
                            if (j == 2)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\" rowspan=\"3\">" + dummy + "</td>");
                            }
                            if (j == 3)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td  style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"2\">" + dummy + "</td>");
                            }

                            if (j == 5)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td  style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"3\">" + dummy + "</td>");
                            }
                     

           
                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 27:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 7; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if (j > 2 && j <6)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }
                            if (j == 6)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"2\">" + dummy + "</td>");
                            }

                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 28:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 7; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if (j > 2)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }

                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 29:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 7; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if (j == 2)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }
                            if (j > 2)
                            {
                                var dummy = Convert.ToDecimal(excelworksheet.Cells[i, j].Value).ToString("#,##0");
                                vars.AppendLine("<td style =\"border:1px solid;text-align: right;\">" + dummy + "</td>");
                            }


                        }
                        vars.AppendLine("</tr><tr><td></td></tr>");
                        break;
                    //29

                    case 31:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 8; j++)
                        {
                            if (j == 1)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td>" + dummy + "</td>");
                            }
                            if (j == 2)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\" rowspan=\"3\">" + dummy + "</td>");
                            }
                            if (j == 3)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td  style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"2\">" + dummy + "</td>");
                            }

                            if (j == 5)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td  style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"2\">" + dummy + "</td>");
                            }
                            if (j == 7)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td  style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"2\">" + dummy + "</td>");
                            }
                         

                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 32:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 8; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if ((j > 2 && j < 5)||j==8||j==7)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }
                            if (j == 5)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"2\">" + dummy + "</td>");
                            }

                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 33:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 8; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if (j > 2)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }

                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 34:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 8; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if (j == 2)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }
                            if (j > 2)
                            {
                                var dummy = Convert.ToDecimal(excelworksheet.Cells[i, j].Value).ToString("#,##0");
                                vars.AppendLine("<td style =\"border:1px solid;text-align: right;\">" + dummy + "</td>");
                            }


                        }
                        vars.AppendLine("</tr><tr><td></td></tr>");
                        break;
                    //34

                    case 36:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 9; j++)
                        {
                            if (j == 1)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td>" + dummy + "</td>");
                            }
                            if (j == 2)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\" rowspan=\"3\">" + dummy + "</td>");
                            }
                            if (j == 3)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td  style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"2\">" + dummy + "</td>");
                            }

                            if (j == 5)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td  style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"5\">" + dummy + "</td>");
                            }
                          


                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 37:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 9; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if (j > 2 && j < 6)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }
                            if (j == 6 || j == 8)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"2\">" + dummy + "</td>");
                            }

                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 38:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 9; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if (j > 2)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }

                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 39:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 9; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if (j == 2)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }
                            if (j > 2)
                            {
                                var dummy = Convert.ToDecimal(excelworksheet.Cells[i, j].Value).ToString("#,##0");
                                vars.AppendLine("<td style =\"border:1px solid;text-align: right;\">" + dummy + "</td>");
                            }


                        }
                        vars.AppendLine("</tr><tr><td></td></tr>");
                        break;
                    //39

                    case 41:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 7; j++)
                        {
                            if (j == 1)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td>" + dummy + "</td>");
                            }
                            if (j == 2)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\" rowspan=\"3\">" + dummy + "</td>");
                            }
                            if (j == 3)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td  style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"2\">" + dummy + "</td>");
                            }

                            if (j == 5)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td  style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"3\">" + dummy + "</td>");
                            }
                         


                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 42:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 7; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if (j > 2 && j < 6)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }
                            if (j == 6)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"2\">" + dummy + "</td>");
                            }

                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 43:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 7; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if (j > 2)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }

                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 44:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 7; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if (j == 2)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }
                            if (j > 2)
                            {
                                var dummy = Convert.ToDecimal(excelworksheet.Cells[i, j].Value).ToString("#,##0");
                                vars.AppendLine("<td style =\"border:1px solid;text-align: right;\">" + dummy + "</td>");
                            }


                        }
                        vars.AppendLine("</tr><tr><td></td></tr>");
                        break;

                    case 46:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 9; j++)
                        {
                            if (j == 1)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td>" + dummy + "</td>");
                            }
                            if (j == 2)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\" rowspan=\"3\">" + dummy + "</td>");
                            }
                            if (j == 3)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td  style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"2\">" + dummy + "</td>");
                            }

                            if (j == 5)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td  style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"5\">" + dummy + "</td>");
                            }
                     


                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 47:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 9; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if (j > 2 && j < 6)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }
                            if (j == 6 || j == 8 )
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"2\">" + dummy + "</td>");
                            }

                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 48:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 9; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if (j > 2)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }

                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 49:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 9; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if (j == 2)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }
                            if (j > 2)
                            {
                                var dummy = Convert.ToDecimal(excelworksheet.Cells[i, j].Value).ToString("#,##0");
                                vars.AppendLine("<td style =\"border:1px solid;text-align: right;\">" + dummy + "</td>");
                            }


                        }
                        vars.AppendLine("</tr><tr><td></td></tr>");
                        break;
                    //49

               

                    //54

                    case 56:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 7; j++)
                        {
                            if (j == 1)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td>" + dummy + "</td>");
                            }
                            if (j == 2)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\" rowspan=\"3\">" + dummy + "</td>");
                            }
                            if (j == 3)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td  style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"2\">" + dummy + "</td>");
                            }

                            if (j == 5)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td  style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"3\">" + dummy + "</td>");
                            }
                  


                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 57:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 7; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if (j > 4)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }
                            if (j == 3)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"2\">" + dummy + "</td>");
                            }

                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 58:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 7; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if (j > 2)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }

                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 59:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 7; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if (j == 2)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }
                            if (j > 2)
                            {
                                var dummy = Convert.ToDecimal(excelworksheet.Cells[i, j].Value).ToString("#,##0");
                                vars.AppendLine("<td style =\"border:1px solid;text-align: right;\">" + dummy + "</td>");
                            }


                        }
                        vars.AppendLine("</tr><tr><td></td></tr>");
                        break;
                    //59
                    case 61:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 8; j++)
                        {
                            if (j == 1)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td>" + dummy + "</td>");
                            }
                            if (j == 2)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\" rowspan=\"3\">" + dummy + "</td>");
                            }
                            if (j == 3)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td  style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"6\">" + dummy + "</td>");
                            }

         
                      


                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 62:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 8; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                         
                            if (j ==3 || j == 5 || j == 7)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"2\">" + dummy + "</td>");
                            }

                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 63:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 8; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if (j > 2)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }

                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 64:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 8; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if (j == 2)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }
                            if (j > 2)
                            {
                                var dummy = Convert.ToDecimal(excelworksheet.Cells[i, j].Value).ToString("#,##0");
                                vars.AppendLine("<td style =\"border:1px solid;text-align: right;\">" + dummy + "</td>");
                            }


                        }
                        vars.AppendLine("</tr><tr><td></td></tr>");
                        break;
                    //64
                    case 51:
                    case 66:
                    case 71:
                    case 76:
                    case 81:
                    case 86:
                    case 91:
                    case 96:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 12; j++)
                        {
                            if (j == 1)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td>" + dummy + "</td>");
                            }
                            if (j == 2)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\" rowspan=\"3\">" + dummy + "</td>");
                            }
                            if (j == 3)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td  style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"2\">" + dummy + "</td>");
                            }

                            if (j == 5)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td  style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"8\">" + dummy + "</td>");
                            }
                        


                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 52:
                    case 67:
                    case 72:
                    case 77:
                    case 82:             
                    case 87:     
                    case 92:
                    case 97:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 12; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if (j > 2 && j < 5)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }
                            if (j == 5 || j == 7 || j == 9 || j == 11)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"2\">" + dummy + "</td>");
                            }

                        }
                        vars.AppendLine("</tr>");
                        break;
                   
                    case 53:
                    case 68:
                    case 73:
                    case 78:
                    case 83:
                    case 88:
                    case 93:
                    case 98:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 12; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if (j > 2)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }

                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 54:
                    case 69:
                    case 74:
                    case 79:
                    case 84:
                    case 89:
                    case 94:
                    case 99:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 12; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if (j == 2)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }
                            if (j > 2)
                            {
                                var dummy = Convert.ToDecimal(excelworksheet.Cells[i, j].Value).ToString("#,##0");
                                vars.AppendLine("<td style =\"border:1px solid;text-align: right;\">" + dummy + "</td>");
                            }


                        }
                        vars.AppendLine("</tr></tr>");
                        break;
                    //69
                    case 101:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 6; j++)
                        {
                            if (j == 1)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td>" + dummy + "</td>");
                            }
                            if (j == 2)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\" rowspan=\"3\">" + dummy + "</td>");
                            }
                            if (j == 3)
                            {

                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td  style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"4\">" + dummy + "</td>");
                            }



                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 102:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 6; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if (j == 3||j==5)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"2\">" + dummy + "</td>");
                            }

                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 103:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 6; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if (j > 2)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }

                        }
                        vars.AppendLine("</tr>");
                        break;
                    case 104:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 6; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if (j == 2)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }
                            if (j > 2)
                            {
                                var dummy = Convert.ToDecimal(excelworksheet.Cells[i, j].Value).ToString("#,##0");
                                vars.AppendLine("<td style =\"border:1px solid;text-align: right;\">" + dummy + "</td>");
                            }


                        }
                        vars.AppendLine("</tr><tr><td></td></tr>");
                        break;
                    case 106:
                    case 107:
                    case 108:
                    case 109:
                    case 110:
                    case 111:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <= 12; j++)
                        {
                            if (j == 1||j==5||j==6||j==7) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if ((j == 4&&i==106)||(j==12&&i==106))
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\">" + dummy + "</td>");
                            }
                            if ((j == 4 && i != 106) || (j == 12 && i != 106))
                            {
                                var dummy = Convert.ToDecimal(excelworksheet.Cells[i, j].Value).ToString("#,##0");
                                vars.AppendLine("<td style =\"border:1px solid;text-align: right;\">" + dummy + "</td>");
                            }
                            if (j == 2)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"2\">" + dummy + "</td>");
                            }
                            if (j == 8)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"4\">" + dummy + "</td>");
                            }
                        

                        }
                        vars.AppendLine("</tr>");
                        break;

                    case 112:
                    case 113:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j <=4 ; j++)
                        {
                            if (j == 1) { var dummy = excelworksheet.Cells[i, j].Value; vars.AppendLine("<td>" + dummy + "</td>"); }
                            if (j == 4  )
                            {
                                var dummy = Convert.ToDecimal(excelworksheet.Cells[i, j].Value).ToString("#,##0");
                                vars.AppendLine("<td style =\"border:1px solid;text-align: right;\">" + dummy + "</td>");
                            }
                            if (j == 2)
                            {
                                var dummy = excelworksheet.Cells[i, j].Value;
                                vars.AppendLine("<td style = \"border:1px solid; background-color: #BDD7EE;\" colspan=\"2\">" + dummy + "</td>");
                            }
                           

                        }
                        vars.AppendLine("</tr>");
                        break;

                    default:
                        vars.AppendLine("<tr>");
                        for (int j = 1; j < 34; j++)
                        {
                            var dummy = excelworksheet.Cells[i, j].Value;
                            vars.AppendLine("<td>" + dummy + "</td>");
                        }
                        vars.AppendLine("</tr>");
                        break;

                }
                }

          
            
            System.Net.Mime.ContentType ct = new System.Net.Mime.ContentType("image/png");
            var webClient = new WebClient();
            Stream slk1 = new MemoryStream(webClient.DownloadData(@"http://10.20.0.13/mrtg/10.20.9.124_10-day.png"));
            Stream slk2 = new MemoryStream(webClient.DownloadData(@"http://10.20.0.13/mrtg/10.20.9.124_10-week.png"));
            Stream slk3 = new MemoryStream(webClient.DownloadData(@"http://10.20.0.13/mrtg/10.20.9.124_10-month.png"));
            Stream slk4 = new MemoryStream(webClient.DownloadData(@"http://10.20.0.13/mrtg/10.20.9.124_10-year.png"));

            LinkedResource lk = new LinkedResource(slk1, ct);
            LinkedResource lk2 = new LinkedResource(slk2, ct);
            LinkedResource lk3 = new LinkedResource(slk3, ct);
            LinkedResource lk4 = new LinkedResource(slk4, ct);

            */
          //  vars.AppendLine("</table>");
         //   vars.AppendLine("</div><div><h3>`Daily' Graph (5 Minute Average)</h3><img src='cid:" + lk.ContentId + @"'></div><div><h3>`Weekly' Graph (30 Minute Average)</h3><img src='cid:" + lk2.ContentId + @"'></div><div><h3>`Monthly' Graph (2 Hour Average)</h3><img src='cid:" + lk3.ContentId + @"'></div><div><h3>`Yearly' Graph (1 Day Average)</h3><img src='cid:" + lk4.ContentId + @"'></div>");
            var a = "reportauto_" + Convert.ToDateTime(settlement_date, CultureInfo.CreateSpecificCulture("en-US")).AddDays(1).ToString("yyyyMMdd", CultureInfo.CreateSpecificCulture("en-US")) + ".xlsx";



            //  string tempPath = AppDomain.CurrentDomain.BaseDirectory + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second + DateTime.Now.Millisecond + "_temp";//date time added to be sure there are no name conflicts
            string tempPath = ConfigurationManager.AppSettings["path"] + "/" + a;
            excelworkbook.SaveAs(tempPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlUserResolution, true, misValue, misValue, misValue);
//            excelworkbook.SaveAs(b, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
    
            tempPath = excelworkbook.FullName;//name of the file with path and extension
            excelworkbook.Close(true, misValue, misValue);
            excelfile.Quit();
            /*
            byte[] attached = File.ReadAllBytes(tempPath);//change to byte[]

            File.Delete(tempPath);//delete temporary file
    
            Stream io = new MemoryStream(attached);
            Attachment attachment = new Attachment(io, a);

           




            MailMessage mailMsg = new MailMessage();
            mailMsg.From = new MailAddress("be@rabbit.co.th");
        
            if (ConfigurationManager.AppSettings["Emailto"] != "") {
                mailMsg.To.Add("panupongp@rabbit.co.th");
                string[] emailto = ConfigurationManager.AppSettings["Emailto"].Split('|');
                foreach (var dummyemail in emailto)
                {
                    mailMsg.To.Add(dummyemail);
                }
            }
           
            mailMsg.Subject = "CCH system daily report(morning) - " + DateTime.Now.ToString("MMM dd", CultureInfo.CreateSpecificCulture("en-US"));
            var view = AlternateView.CreateAlternateViewFromString(vars.ToString(), null, "text/html");
            view.LinkedResources.Add(lk);
            view.LinkedResources.Add(lk2);
            view.LinkedResources.Add(lk3);
            view.LinkedResources.Add(lk4);
            mailMsg.AlternateViews.Add(view);

         

         
            mailMsg.Attachments.Add(attachment);
            // Send the email.
            SmtpClient smtp = new SmtpClient();
            smtp.UseDefaultCredentials = true;
            smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
            // smtp.Credentials = new NetworkCredential("sm1405", "VmEJjdO8rdHqJTm");
            smtp.EnableSsl = false;
            smtp.Host = ConfigurationManager.AppSettings["host"];
            
            try
            {
                
                smtp.Port = Convert.ToInt32(ConfigurationManager.AppSettings["port"]);

            }
            catch {
                smtp.Port = 25;
            }
          
         
            smtp.Send(mailMsg);
            mailMsg.Dispose();
            mailMsg = null;


            Marshal.ReleaseComObject(excelworksheet);
            Marshal.ReleaseComObject(excelworkbook);
            Marshal.ReleaseComObject(excelfile);
            */
        }
    }
}
public class function
{
    public DataTable connect_file()
    {

        string cmd = "select * from bss_reports.daily_check_system_level where settlement_date = trunc(sysdate-" + ConfigurationManager.AppSettings["date"] + ") and status='0'";
        OracleConnection orcon = new OracleConnection();
        OracleCommand orCmd = new OracleCommand();
        OracleDataAdapter dtAdapter = new OracleDataAdapter();

        DataTable ds = new DataTable();

        try
        {
            orcon.ConnectionString = ConfigurationManager.ConnectionStrings["OrclDB"].ConnectionString;
            orcon.Open();
            orCmd.Connection = orcon;
            orCmd.CommandText = cmd;
            orCmd.CommandType = CommandType.Text;

            dtAdapter.SelectCommand = orCmd;

            dtAdapter.Fill(ds);

            dtAdapter = null;
        }
        catch (OracleException ex)
        {
            Console.WriteLine(ex);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex);
        }
        finally
        {
            if (orcon != null)
            {
                orcon.Close();
                orcon.Dispose();
                orcon = null;
            }
            if (orCmd != null)
            {
                orCmd.Dispose();
                orCmd = null;
            }


        }
        return ds;
    }

    public DataTable connect_purse()
    {
        string cmd = " SELECT settlement_date,'ATU' AS rpt_grp,reports.bkk_int_fun.getpassengertype(application_passenger_type, reports.bkk_int_fun.getdataversion(), 'en') AS card_type," +
"REPORTS.BKK_INT_FUN.GETTRANSACTIONDESC(UD_TYPE, UD_SUBTYPE) AS txn_type,NULL AS product_sale_desc,COUNT(card_serial_number)AS txn_vol,0 AS txn_value FROM CUT_PI_MAINTENANCE cut " +
"WHERE cut.ud_type = 3 AND cut.ud_subtype IN(9, 51, 59) AND cut.iss_txn_reflection = 'N' AND cut.cch_txn_approved = 'Y' AND cut.source_participant_id <> 88 AND cut.settlement_date = TRUNC(sysdate - 1)" +
" GROUP BY settlement_date,'ATU',reports.bkk_int_fun.getpassengertype(application_passenger_type, reports.bkk_int_fun.getdataversion(), 'en'),REPORTS.BKK_INT_FUN.GETTRANSACTIONDESC(UD_TYPE, UD_SUBTYPE)" +
" UNION SELECT settlement_date,'Hero' AS rpt_grp, reports.bkk_int_fun.getpassengertype(aa.PASSENGER_TYPE, reports.bkk_int_fun.getdataversion(), 'en') AS card_type," +
"REPORTS.BKK_INT_FUN.GETTRANSACTIONDESC(UD_TYPE, UD_SUBTYPE) AS txn_type,bss_reports.pkg_fun.getproductsaledesc(card_serial_number, card_issuer_id, card_life_cycle_count, card_type) AS product_sale_desc," +
"COUNT(bss_reports.pkg_fun.getproductsaledesc(card_serial_number, card_issuer_id, card_life_cycle_count, card_type)) AS txn_vol,0 AS txn_value FROM CUT_CI_MAINTENANCE LEFT JOIN application_account aa" +
" ON card_serial_number = aa.csc_serial_number AND card_life_cycle_count = aa.csc_lifecycle_count AND card_issuer_id = csc_issuer_id WHERE settlement_date = TRUNC(sysdate - 1) AND ud_type = 1 AND ud_subtype = 3" +
" AND iss_txn_reflection = 'N' AND aa.PASSENGER_TYPE    IN(21, 22, 23, 24) GROUP BY settlement_date, 'Hero', reports.bkk_int_fun.getpassengertype(aa.PASSENGER_TYPE, reports.bkk_int_fun.getdataversion(), 'en')," +
" REPORTS.BKK_INT_FUN.GETTRANSACTIONDESC(UD_TYPE, UD_SUBTYPE),bss_reports.pkg_fun.getproductsaledesc(card_serial_number, card_issuer_id, card_life_cycle_count, card_type) UNION SELECT settlement_date," +
" rpt_grp,NULL AS card_type,transaction_type AS txn_type,NULL AS product_sale_desc,txn_cnt AS txn_vol,txn_val AS txn_value FROM bss_reports.vw_system_daily_grp_rpt WHERE settlement_date = TRUNC(sysdate - " + ConfigurationManager.AppSettings["date"] + ") ORDER BY RPT_GRP";
        OracleConnection orcon = new OracleConnection();

        OracleCommand orCmd = new OracleCommand();
        OracleDataAdapter dtAdapter = new OracleDataAdapter();

        DataTable ds = new DataTable();

        try
        {
            orcon.ConnectionString = ConfigurationManager.ConnectionStrings["OrclDB"].ConnectionString;
            orcon.Open();
            orCmd.Connection = orcon;
            orCmd.CommandText = cmd;
            orCmd.CommandType = CommandType.Text;

            dtAdapter.SelectCommand = orCmd;

            dtAdapter.Fill(ds);

            dtAdapter = null;
        }
        catch (OracleException ex)
        {
            Console.WriteLine(ex);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex);
        }
        finally
        {
            if (orcon != null)
            {
                orcon.Close();
                orcon.Dispose();
                orcon = null;
            }
            if (orCmd != null)
            {
                orCmd.Dispose();
                orCmd = null;
            }


        }
        return ds;

    }

   

    public static Attachment CreateAttachment(Stream attachmentFile, string displayName)
    {
        var attachment = new Attachment(attachmentFile, displayName);
        attachment.ContentType = new ContentType("application/vnd.ms-excel");
        attachment.TransferEncoding = TransferEncoding.Base64;
        attachment.NameEncoding = Encoding.UTF8;
        string encodedAttachmentName = Convert.ToBase64String(Encoding.UTF8.GetBytes(displayName));
        encodedAttachmentName = SplitEncodedAttachmentName(encodedAttachmentName);
        attachment.Name = encodedAttachmentName;
        return attachment;
    }

    private static string SplitEncodedAttachmentName(string encoded)
    {
        const string encodingtoken = "=?UTF-8?B?";
        const string softbreak = "?=";
        const int maxChunkLength = 30;
        int splitLength = maxChunkLength - encodingtoken.Length - (softbreak.Length * 2);
        IEnumerable<string> parts = SplitByLength(encoded, splitLength);
        string encodedAttachmentName = encodingtoken;
        foreach (var part in parts)
        {
            encodedAttachmentName += part + softbreak + encodingtoken;
        }
        encodedAttachmentName = encodedAttachmentName.Remove(encodedAttachmentName.Length - encodingtoken.Length, encodingtoken.Length);
        return encodedAttachmentName;
    }

    private static IEnumerable<string> SplitByLength(string stringToSplit, int length)
    {
        while (stringToSplit.Length > length)
        {
            yield return stringToSplit.Substring(0, length);
            stringToSplit = stringToSplit.Substring(length);
        }
        if (stringToSplit.Length > 0)
        {
            yield return stringToSplit;
        }
    }
}

