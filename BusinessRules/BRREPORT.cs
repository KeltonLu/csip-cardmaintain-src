//******************************************************************
//*  作    者：yangyu(rosicky)
//*  功能說明：報表操作類
//*  創建日期：2009/09/21
//*  修改記錄：調整SP執行，增加 queryType 參數判斷是S：查詢或P：列印 by Ares Stanley 20220207
//*<author>            <time>            <TaskID>                <desc>
//*******************************************************************
using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using Framework.Data.OM;
using CSIPCardMaintain.EntityLayer;
using System.Data;
using Framework.Data.OM.Collections;
using Framework.Data.OM.Transaction;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using System.Configuration;
using Framework.Common.Message;
using CSIPCommonModel.EntityLayer;
using System.Web;
using Framework.Data;
using Framework.Common.Logging;
using Framework.Common.Utility;

namespace CSIPCardMaintain.BusinessRules
{
    public class BRReport : CSIPCommonModel.BusinessRules.BRBase<EntityReport>
    {


        #region 方法

        #region 信用卡卡人

        /// <summary>
        /// 卡人維護記錄查詢(Pull模式)
        /// </summary>
        /// <param name="strID">卡人ID</param>
        /// <param name="strPeople">維護員</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">結束日期</param>
        /// <param name="blnOld">是否查詢一年半以前資料</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strAgentID">業務員ID</param>
        /// <param name="blnHadRecord">是否有記錄</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否成功</returns>
        public static bool Report01010100_BySP(string strID, string strPeople, string strBeforeDate, string strEndDate, bool blnOld, string strSEQ, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {

            try
            {
                //*宣告變數
                DataHelper dh = new DataHelper();
                string strHistory = "";

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";
                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                if (strID != "")//*如果有輸入卡人
                {
                    strID = ConvertID(strID);
                }
                //*查詢一年半以前資料
                if (blnOld == true)
                {
                    strHistory = "Y";
                }
                else
                {
                    strHistory = "N";
                }

                PrintStoredProcedure("SP_RptMaintainLog", strID, strAgentID, strPeople, strBeforeDate, strEndDate, "", "", "", strHistory, strSEQ, "01010100", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();

                //Start 新增TimeOut  by Neal 20220531
                SqlCommand sqlcmd = new SqlCommand();
                sqlcmd.CommandText = "SP_RptMaintainLog";
                sqlcmd.CommandTimeout = int.Parse(UtilHelper.GetAppSettings("PageSqlCmdTimeoutMax"));
                sqlcmd.CommandType = CommandType.StoredProcedure;
                sqlcmd.Parameters.Add(new SqlParameter("@DB_CUST_ID", strID.ToString().Trim()));
                sqlcmd.Parameters.Add(new SqlParameter("@AGENT_ID", strAgentID));
                sqlcmd.Parameters.Add(new SqlParameter("@UPD_AGENT_ID", strPeople));
                sqlcmd.Parameters.Add(new SqlParameter("@MAINT_D_S", strBeforeDate));
                sqlcmd.Parameters.Add(new SqlParameter("@MAINT_D_E", strEndDate));
                sqlcmd.Parameters.Add(new SqlParameter("@FLD_NAME", ""));
                sqlcmd.Parameters.Add(new SqlParameter("@Value1", ""));
                sqlcmd.Parameters.Add(new SqlParameter("@Value2", ""));
                sqlcmd.Parameters.Add(new SqlParameter("@Query_History", strHistory));
                sqlcmd.Parameters.Add(new SqlParameter("@SEQ_Name", strSEQ));
                sqlcmd.Parameters.Add(new SqlParameter("@strAction", "01010100"));
                blnHadRecord = dh.ExecuteNonQuery(sqlcmd) > 0;
                //End  新增TimeOut  by Neal 20220531 

                //blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", strID, strAgentID, strPeople, strBeforeDate, strEndDate, "", "", "", strHistory, strSEQ, "01010100") > 0;
                sw.Stop();
                RecordSPExecuteTime("01010100", sw);

                return true;
            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }
        }


        //add by linhuanhuang 新增自扣資料查詢 2010/8/13 start
        /// <summary>
        /// 卡人自扣查詢(Pull模式)
        /// </summary>
        /// <param name="strID">卡人ID</param>
        /// <param name="strPeople">維護員</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">結束日期</param>
        /// <param name="blnOld">是否查詢一年半以前資料</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strAgentID">業務員ID</param>
        /// <param name="blnHadRecord">是否有記錄</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否成功</returns>
        public static bool Report01020000_BySP(string strID, string strPeople, string strBeforeDate, string strEndDate, bool blnOld, string strSEQ, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {

            try
            {
                //*宣告變數
                DataHelper dh = new DataHelper();
                string strHistory = "";

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";
                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                if (strID != "")//*如果有輸入卡人
                {
                    strID = ConvertID(strID);
                }
                //*查詢一年半以前資料
                if (blnOld == true)
                {
                    strHistory = "Y";
                }
                else
                {
                    strHistory = "N";
                }

                PrintStoredProcedure("SP_RptMaintainLog", strID, strAgentID, strPeople, strBeforeDate, strEndDate, "", "", "", strHistory, strSEQ, "01020000", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", strID, strAgentID, strPeople, strBeforeDate, strEndDate, "", "", "", strHistory, strSEQ, "01020000") > 0;
                sw.Stop();
                RecordSPExecuteTime("01020000", sw);

                return true;
            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }
        }
        //add by linhuanhuang 新增自扣資料查詢 2010/8/13 end


        /// <summary>
        /// 卡人維護員統計表查詢
        /// </summary>
        /// <param name="strID">卡人ID</param>
        /// <param name="strPeople">維護員</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">結束日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="blnOld">是否查詢一年半前資料</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report01010201_BySP(string strPeople, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, bool blnOld, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                //*宣告變數
                DataHelper dh = new DataHelper();
                string strHistory = "";

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";
                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                //*查詢一年半以前資料
                if (blnOld == true)
                {
                    strHistory = "Y";
                }
                else
                {
                    strHistory = "N";
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, "", "", "", strHistory, strSEQ, "01010201", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, "", "", "", strHistory, strSEQ, "01010201") > 0;
                sw.Stop();
                RecordSPExecuteTime("01010201", sw);
                return true;

            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }
        }

        /// <summary>
        /// 卡人>卡人與維護員關係表
        /// </summary>
        /// <param name="strID">卡人ID</param>
        /// <param name="strPeople">維護員</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">結束日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="blnOld">是否查詢一年半前資料</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report01010202_BySP(string strID, string strPeople, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, bool blnOld, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                //*宣告變數
                DataHelper dh = new DataHelper();
                string strHistory = "";

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";

                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                if (strID != "")//*如果有輸入卡人
                {
                    strID = ConvertID(strID);
                }

                //*查詢一年半以前資料
                if (blnOld == true)
                {
                    strHistory = "Y";
                }
                else
                {
                    strHistory = "N";
                }

                PrintStoredProcedure("SP_RptMaintainLog", strID, strAgentID, strPeople, strBeforeDate, strEndDate, "", "", "", strHistory, strSEQ, "01010202", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", strID, strAgentID, strPeople, strBeforeDate, strEndDate, "", "", "", strHistory, strSEQ, "01010202") > 0;
                sw.Stop();
                RecordSPExecuteTime("01010202", sw);
                return true;
            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }
        }

        /// <summary>
        /// 卡人維護欄位統計表查詢
        /// </summary>
        /// <param name="strFld">維護欄位</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">結束日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="blnOld">是否查詢一年半前資料</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report01010203_BySP(string strFld, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, bool blnOld, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                //*宣告變數
                DataHelper dh = new DataHelper();
                string strHistory = "";

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";

                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                //*查詢一年半以前資料
                if (blnOld == true)
                {
                    strHistory = "Y";
                }
                else
                {
                    strHistory = "N";
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, "", strBeforeDate, strEndDate, strFld, "", "", strHistory, strSEQ, "01010203", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, "", strBeforeDate, strEndDate, strFld, "", "", strHistory, strSEQ, "01010203") > 0;
                sw.Stop();
                RecordSPExecuteTime("01010203", sw);
                return true;

            }
            catch (Exception exp)
            {

                BRReport.SaveLog(exp);
                return false;
            }




        }

        /// <summary>
        /// 卡人調整統計表查詢
        /// </summary>
        /// <param name="strFld">維護欄位</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">結束日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="blnOld">是否查詢一年半前資料</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report01010204_BySP(string strFld, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, bool blnOld, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                //*宣告變數
                DataHelper dh = new DataHelper();
                string strHistory = "";

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";

                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                //*查詢一年半以前資料
                if (blnOld == true)
                {
                    strHistory = "Y";
                }
                else
                {
                    strHistory = "N";
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, "", strBeforeDate, strEndDate, strFld, "", "", strHistory, strSEQ, "01010204", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, "", strBeforeDate, strEndDate, strFld, "", "", strHistory, strSEQ, "01010204") > 0;
                sw.Stop();
                RecordSPExecuteTime("01010204", sw);
                return true;

            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }




        }

        /// <summary>
        /// 卡人調整固定額度查詢
        /// </summary>
        /// <param name="strPeople">維護員</param>
        /// <param name="strBeforeAmount">起始固定額度</param>
        /// <param name="strEndAmount">終止固定額度</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">終止日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="blnOld">是否查詢一年半前資料</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report01010301_BySP(string strPeople, string strBeforeAmount, string strEndAmount, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, bool blnOld, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                //*宣告變數
                DataHelper dh = new DataHelper();
                string strHistory = "";
                String strFldName = "CREDIT LINE PERM"; //固定額度

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";

                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                //*查詢一年半以前資料
                if (blnOld == true)
                {
                    strHistory = "Y";
                }
                else
                {
                    strHistory = "N";
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeAmount, strEndAmount, strHistory, strSEQ, "01010301", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeAmount, strEndAmount, strHistory, strSEQ, "01010301") > 0;
                sw.Stop();
                RecordSPExecuteTime("01010301", sw);
                return true;

            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }

        }

        /// <summary>
        /// 卡人調整臨時額度查詢
        /// </summary>
        /// <param name="strPeople">維護員</param>
        /// <param name="strBeforeAmount">起始固定額度</param>
        /// <param name="strEndAmount">終止固定額度</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">終止日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="blnOld">是否查詢一年半前資料</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report01010302_BySP(string strPeople, string strBeforeAmount, string strEndAmount, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, bool blnOld, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                //*宣告變數
                DataHelper dh = new DataHelper();
                string strHistory = "";
                String strFldName = "CREDIT LINE TEMP"; //臨時額度

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";

                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                //*查詢一年半以前資料
                if (blnOld == true)
                {
                    strHistory = "Y";
                }
                else
                {
                    strHistory = "N";
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeAmount, strEndAmount, strHistory, strSEQ, "01010302", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeAmount, strEndAmount, strHistory, strSEQ, "01010302") > 0;
                sw.Stop();
                RecordSPExecuteTime("01010302", sw);
                return true;

            }
            catch (Exception exp)
            {

                BRReport.SaveLog(exp);
                return false;
            }




        }

        /// <summary>
        /// 卡人新卡額度查詢
        /// </summary>
        /// <param name="strPeople">維護員</param>
        /// <param name="strBeforeAmount">起始固定額度</param>
        /// <param name="strEndAmount">終止固定額度</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">終止日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="blnOld">是否查詢一年半前資料</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report01010303_BySP(string strPeople, string strBeforeAmount, string strEndAmount, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, bool blnOld, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                //*宣告變數
                DataHelper dh = new DataHelper();
                string strHistory = "";
                String strFldName = "CR LINE CURR & PERM"; //新卡額度

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";

                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                //*查詢一年半以前資料
                if (blnOld == true)
                {
                    strHistory = "Y";
                }
                else
                {
                    strHistory = "N";
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeAmount, strEndAmount, strHistory, strSEQ, "01010303", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeAmount, strEndAmount, strHistory, strSEQ, "01010303") > 0;
                sw.Stop();
                RecordSPExecuteTime("01010303", sw);
                return true;

            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }
        }

        /// <summary>
        /// 卡人員工調整記錄查詢
        /// </summary>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">終止日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="blnOld">是否查詢一年半前資料</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report01010401_BySP(string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, bool blnOld, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                //*宣告變數
                DataHelper dh = new DataHelper();
                string strHistory = "";

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";

                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                //*查詢一年半以前資料
                if (blnOld == true)
                {
                    strHistory = "Y";
                }
                else
                {
                    strHistory = "N";
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, "", strBeforeDate, strEndDate, "", "", "", strHistory, strSEQ, "01010401", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, "", strBeforeDate, strEndDate, "", "", "", strHistory, strSEQ, "01010401") > 0;
                sw.Stop();
                RecordSPExecuteTime("01010401", sw);
                return true;

            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }
        }

        /// <summary>
        /// 卡人自扣帳戶ID與卡人ID不同者
        /// </summary>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">結束日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="blnOld">是否查詢一年半前資料</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report01010402_BySP(string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, bool blnOld, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                //*宣告變數
                DataHelper dh = new DataHelper();
                string strHistory = "";

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";

                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                //*查詢一年半以前資料
                if (blnOld == true)
                {
                    strHistory = "Y";
                }
                else
                {
                    strHistory = "N";
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, "", strBeforeDate, strEndDate, "", "", "", strHistory, strSEQ, "01010402", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, "", strBeforeDate, strEndDate, "", "", "", strHistory, strSEQ, "01010402") > 0;
                sw.Stop();
                RecordSPExecuteTime("01010402", sw);
                return true;
            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }
        }


        #endregion

        #region 信用卡卡片
        /// <summary>
        /// 卡片維護記錄查詢
        /// </summary>
        /// <param name="strID">卡號</param>
        /// <param name="strPeople">維護員</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">結束日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="blnOld">是否查詢一年半前資料</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report02010100_BySP(string strID, string strPeople, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, bool blnOld, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                //*宣告變數
                DataHelper dh = new DataHelper();
                string strHistory = "";

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";
                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }



                //*查詢一年半以前資料
                if (blnOld == true)
                {
                    strHistory = "Y";
                }
                else
                {
                    strHistory = "N";
                }

                PrintStoredProcedure("SP_RptMaintainLog", strID.ToString().Trim(), strAgentID, strPeople, strBeforeDate, strEndDate, "", "", "", strHistory, strSEQ, "02010100", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                //Start 新增TimeOut  by Neal 20220531
                SqlCommand sqlcmd = new SqlCommand();
                sqlcmd.CommandText = "SP_RptMaintainLog";
                sqlcmd.CommandTimeout = int.Parse(UtilHelper.GetAppSettings("PageSqlCmdTimeoutMax"));
                sqlcmd.CommandType = CommandType.StoredProcedure;
                sqlcmd.Parameters.Add(new SqlParameter("@DB_CUST_ID", strID.ToString().Trim()));
                sqlcmd.Parameters.Add(new SqlParameter("@AGENT_ID", strAgentID));
                sqlcmd.Parameters.Add(new SqlParameter("@UPD_AGENT_ID", strPeople));
                sqlcmd.Parameters.Add(new SqlParameter("@MAINT_D_S", strBeforeDate));
                sqlcmd.Parameters.Add(new SqlParameter("@MAINT_D_E", strEndDate));
                sqlcmd.Parameters.Add(new SqlParameter("@FLD_NAME", ""));
                sqlcmd.Parameters.Add(new SqlParameter("@Value1", ""));
                sqlcmd.Parameters.Add(new SqlParameter("@Value2", ""));
                sqlcmd.Parameters.Add(new SqlParameter("@Query_History", strHistory));
                sqlcmd.Parameters.Add(new SqlParameter("@SEQ_Name", strSEQ));
                sqlcmd.Parameters.Add(new SqlParameter("@strAction", "02010100"));
                blnHadRecord = dh.ExecuteNonQuery(sqlcmd) > 0;
                //End  新增TimeOut  by Neal 20220531 
                sw.Stop();
                RecordSPExecuteTime("02010100", sw);
                return true;
            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }


        }

        /// <summary>
        /// 卡片維護員統計表查詢
        /// </summary>
        /// <param name="strPeople">維護員</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">結束日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="blnOld">是否查詢一年半前資料</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report02020100_BySP(string strPeople, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, bool blnOld, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                //*宣告變數
                DataHelper dh = new DataHelper();
                string strHistory = "";

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";
                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                //*查詢一年半以前資料
                if (blnOld == true)
                {
                    strHistory = "Y";
                }
                else
                {
                    strHistory = "N";
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, "", "", "", strHistory, strSEQ, "02020100", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, "", "", "", strHistory, strSEQ, "02020100") > 0;
                sw.Stop();
                RecordSPExecuteTime("02020100", sw);
                return true;

            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }




        }

        /// <summary>
        /// 卡片>卡片與維護員關係表

        /// </summary>
        /// <param name="strID">卡號</param>
        /// <param name="strPeople">維護員</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">結束日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="blnOld">是否查詢一年半前資料</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report02010202_BySP(string strID, string strPeople, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, bool blnOld, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                //*宣告變數
                DataHelper dh = new DataHelper();
                string strHistory = "";

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";
                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }



                //*查詢一年半以前資料
                if (blnOld == true)
                {
                    strHistory = "Y";
                }
                else
                {
                    strHistory = "N";
                }

                PrintStoredProcedure("SP_RptMaintainLog", strID.ToString().Trim(), strAgentID, strPeople, strBeforeDate, strEndDate, "", "", "", strHistory, strSEQ, "02010202", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", strID.ToString().Trim(), strAgentID, strPeople, strBeforeDate, strEndDate, "", "", "", strHistory, strSEQ, "02010202") > 0;
                sw.Stop();
                RecordSPExecuteTime("02010202", sw);
                return true;

            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }


        }

        /// <summary>
        /// 卡片維護欄位統計表查詢

        /// </summary>
        /// <param name="strFld">維護欄位</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">結束日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="blnOld">是否查詢一年半前資料</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report02010203_BySP(string strFld, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, bool blnOld, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                //*宣告變數
                DataHelper dh = new DataHelper();
                string strHistory = "";

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";

                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                //*查詢一年半以前資料
                if (blnOld == true)
                {
                    strHistory = "Y";
                }
                else
                {
                    strHistory = "N";
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, "", strBeforeDate, strEndDate, strFld, "", "", strHistory, strSEQ, "02010203", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, "", strBeforeDate, strEndDate, strFld, "", "", strHistory, strSEQ, "02010203") > 0;
                sw.Stop();
                RecordSPExecuteTime("02010203", sw);
                return true;
            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }




        }

        /// <summary>
        /// 卡片調整統計表查詢

        /// </summary>
        /// <param name="strFld">維護欄位</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">結束日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="blnOld">是否查詢一年半前資料</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report02010204_BySP(string strFld, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, bool blnOld, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                //*宣告變數
                DataHelper dh = new DataHelper();
                string strHistory = "";

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";

                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                //*查詢一年半以前資料
                if (blnOld == true)
                {
                    strHistory = "Y";
                }
                else
                {
                    strHistory = "N";
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, "", strBeforeDate, strEndDate, strFld, "", "", strHistory, strSEQ, "02010204", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, "", strBeforeDate, strEndDate, strFld, "", "", strHistory, strSEQ, "02010204") > 0;
                sw.Stop();
                RecordSPExecuteTime("02010204", sw);
                return true;


            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }




        }

        /// <summary>
        /// 卡片調整固定額度查詢
        /// </summary>
        /// <param name="strPeople">卡號</param>
        /// <param name="strBeforeAmount">起始固定額度</param>
        /// <param name="strEndAmount">終止固定額度</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">終止日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="blnOld">是否查詢一年半前資料</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report02010301_BySP(string strPeople, string strBeforeAmount, string strEndAmount, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, bool blnOld, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                //*宣告變數
                DataHelper dh = new DataHelper();
                string strHistory = "";
                String strFldName = "CREDIT LINE PERM"; //固定額度

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";

                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                //*查詢一年半以前資料
                if (blnOld == true)
                {
                    strHistory = "Y";
                }
                else
                {
                    strHistory = "N";
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeAmount, strEndAmount, strHistory, strSEQ, "02010301", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeAmount, strEndAmount, strHistory, strSEQ, "02010301") > 0;
                sw.Stop();
                RecordSPExecuteTime("02010301", sw);
                return true;

            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }




        }

        /// <summary>
        /// 卡片調整臨時額度查詢
        /// </summary>
        /// <param name="strPeople">卡號</param>
        /// <param name="strBeforeAmount">起始固定額度</param>
        /// <param name="strEndAmount">終止固定額度</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">終止日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="blnOld">是否查詢一年半前資料</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report02010302_BySP(string strPeople, string strBeforeAmount, string strEndAmount, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, bool blnOld, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                //*宣告變數
                DataHelper dh = new DataHelper();
                string strHistory = "";
                String strFldName = "CREDIT LINE TEMP"; //臨時額度

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";

                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                //*查詢一年半以前資料
                if (blnOld == true)
                {
                    strHistory = "Y";
                }
                else
                {
                    strHistory = "N";
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeAmount, strEndAmount, strHistory, strSEQ, "02010302", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();

                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeAmount, strEndAmount, strHistory, strSEQ, "02010302") > 0;
                sw.Stop();
                RecordSPExecuteTime("02010302", sw);
                return true;

            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }




        }

        /// <summary>
        /// 卡片新卡額度查詢
        /// </summary>
        /// <param name="strPeople">卡號</param>
        /// <param name="strBeforeAmount">起始固定額度</param>
        /// <param name="strEndAmount">終止固定額度</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">終止日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="blnOld">是否查詢一年半前資料</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report02010303_New(string strPeople, string strBeforeAmount, string strEndAmount, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, bool blnOld, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                //*宣告變數
                DataHelper dh = new DataHelper();
                string strHistory = "";
                String strFldName = "CR LINE CURR & PERM"; //新卡額度

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";

                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                //*查詢一年半以前資料
                if (blnOld == true)
                {
                    strHistory = "Y";
                }
                else
                {
                    strHistory = "N";
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeAmount, strEndAmount, strHistory, strSEQ, "02010303", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeAmount, strEndAmount, strHistory, strSEQ, "02010303") > 0;
                sw.Stop();
                RecordSPExecuteTime("02010303", sw);
                return true;

            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }




        }

        /// <summary>
        /// 卡片非流通BlockCode調整爲流通中
        /// </summary>
        /// <param name="strPeople">維護員</param>
        /// <param name="strBeforeContent">調整前內容</param>
        /// <param name="strEndContent">調整後內容</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">終止日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="blnOld">是否查詢一年半前資料</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report02010401_BySP(string strPeople, string strBeforeContent, string strEndContent, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, bool blnOld, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                //*宣告變數
                DataHelper dh = new DataHelper();
                string strHistory = "";
                String strFldName = "PRIMARY BLOCK CODE";
                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";

                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                //*查詢一年半以前資料
                if (blnOld == true)
                {
                    strHistory = "Y";
                }
                else
                {
                    strHistory = "N";
                }
                if (strBeforeContent != "")
                {
                    //20211207 註解避免多一層小括號造成storedprocedure錯誤 by Ares Stanley
                    //strBeforeContent = GetRange(strBeforeContent);//*輸入了調整前內容

                }

                if (strEndContent != "")
                {
                    //20211207 註解避免多一層小括號造成storedprocedure錯誤 by Ares Stanley
                    //strEndContent = GetRange(strEndContent);//*輸入了調整後內容
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeContent, strEndContent, strHistory, strSEQ, "02010401", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeContent, strEndContent, strHistory, strSEQ, "02010401") > 0;
                sw.Stop();
                RecordSPExecuteTime("02010401", sw);
                return true;
            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }




        }

        /// <summary>
        /// 信用卡卡片年費無優惠調整爲優惠條件
        /// </summary>
        /// <param name="strBeforeContent">調整前內容</param>
        /// <param name="strEndContent">調整後內容</param>
        /// <param name="strPeople">維護員</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">結束日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="blnOld">是否查詢一年半前資料</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report02010402_BySP(string strBeforeContent, string strEndContent, string strPeople, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, bool blnOld, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                //*宣告變數
                DataHelper dh = new DataHelper();
                string strHistory = "";
                String strFldName = "USER CODE 01";
                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";

                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                //*查詢一年半以前資料
                if (blnOld == true)
                {
                    strHistory = "Y";
                }
                else
                {
                    strHistory = "N";
                }
                if (strBeforeContent != "")
                {
                    //20211207 註解避免多一層小括號造成storedprocedure錯誤 by Ares Stanley
                    //strBeforeContent = GetRange(strBeforeContent);//*輸入了調整前內容

                }

                if (strEndContent != "")
                {
                    //20211207 註解避免多一層小括號造成storedprocedure錯誤 by Ares Stanley
                    //strEndContent = GetRange(strEndContent);//*輸入了調整後內容
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeContent, strEndContent, strHistory, strSEQ, "02010402", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeContent, strEndContent, strHistory, strSEQ, "02010402") > 0;
                sw.Stop();
                RecordSPExecuteTime("02010402", sw);
                return true;
            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }
        }

        #endregion

        #region VD卡人
        /// <summary>
        /// VD卡人維護記錄查詢
        /// </summary>
        /// <param name="strID">卡人ID</param>
        /// <param name="strPeople">維護員</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">結束日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report03010100_BySP(string strID, string strPeople, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                //*宣告變數
                DataHelper dh = new DataHelper();

                string strDBID = "";

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";
                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                if (strID != "")//*如果有輸入卡人
                {
                    strDBID = ConvertID(strID);
                }

                PrintStoredProcedure("SP_RptMaintainLog", strDBID, strAgentID, strPeople, strBeforeDate, strEndDate, "", "", "", "N", strSEQ, "03010100", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                //20220530_Ares_Jack_新增TimeOut時間
                SqlCommand sqlcmd = new SqlCommand();
                sqlcmd.CommandType = CommandType.StoredProcedure;
                sqlcmd.CommandText = "SP_RptMaintainLog";
                sqlcmd.CommandTimeout = int.Parse(UtilHelper.GetAppSettings("PageSqlCmdTimeoutMax"));
                sqlcmd.Parameters.Add(new SqlParameter("@DB_CUST_ID", strDBID));
                sqlcmd.Parameters.Add(new SqlParameter("@AGENT_ID", strAgentID));
                sqlcmd.Parameters.Add(new SqlParameter("@UPD_AGENT_ID", strPeople));
                sqlcmd.Parameters.Add(new SqlParameter("@MAINT_D_S", strBeforeDate));
                sqlcmd.Parameters.Add(new SqlParameter("@MAINT_D_E", strEndDate));
                sqlcmd.Parameters.Add(new SqlParameter("@FLD_NAME", ""));
                sqlcmd.Parameters.Add(new SqlParameter("@Value1", ""));
                sqlcmd.Parameters.Add(new SqlParameter("@Value2", ""));
                sqlcmd.Parameters.Add(new SqlParameter("@Query_History", "N"));
                sqlcmd.Parameters.Add(new SqlParameter("@SEQ_Name", strSEQ));
                sqlcmd.Parameters.Add(new SqlParameter("@strAction", "03010100"));
                blnHadRecord = dh.ExecuteNonQuery(sqlcmd) > 0;

                //blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", strDBID, strAgentID, strPeople, strBeforeDate, strEndDate, "", "", "", "N", strSEQ, "03010100") > 0;
                sw.Stop();
                RecordSPExecuteTime("03010100", sw);
                return true;

            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }

        }

        /// <summary>
        /// VD卡人維護員統計表查詢
        /// </summary>
        /// <param name="strPeople">維護員</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">結束日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report03010201_BySP(string strPeople, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                //*宣告變數
                DataHelper dh = new DataHelper();

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";
                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, "", "", "", "N", strSEQ, "03010201", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, "", "", "", "N", strSEQ, "03010201") > 0;
                sw.Stop();
                RecordSPExecuteTime("03010201", sw);
                return true;
            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }




        }

        /// <summary>
        /// VD卡人>卡人與維護員關係表
        /// </summary>
        /// <param name="strID">卡人ID</param>
        /// <param name="strPeople">維護員</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">結束日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report03010202_BySP(string strID, string strPeople, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                //*宣告變數
                DataHelper dh = new DataHelper();
                string strDBID = "";

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";
                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                if (strID != "")//*如果有輸入卡人
                {
                    strDBID = ConvertID(strID);
                }

                PrintStoredProcedure("SP_RptMaintainLog", strDBID, strAgentID, strPeople, strBeforeDate, strEndDate, "", "", "", "N", strSEQ, "03010202", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", strDBID, strAgentID, strPeople, strBeforeDate, strEndDate, "", "", "", "N", strSEQ, "03010202") > 0;
                sw.Stop();
                RecordSPExecuteTime("03010202", sw);
                return true;

            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }

        }

        /// <summary>
        /// VD卡人維護欄位統計表查詢
        /// </summary>
        /// <param name="strFld">維護欄位</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">結束日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report03010203_BySP(string strFld, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                //*宣告變數
                DataHelper dh = new DataHelper();

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";

                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, "", strBeforeDate, strEndDate, strFld, "", "", "N", strSEQ, "03010203", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, "", strBeforeDate, strEndDate, strFld, "", "", "N", strSEQ, "03010203") > 0;
                sw.Stop();
                RecordSPExecuteTime("03010203", sw);
                return true;

            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }
        }

        /// <summary>
        /// VD卡人調整統計表查詢

        /// </summary>
        /// <param name="strFld">維護欄位</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">結束日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report03010204_BySP(string strFld, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                //*宣告變數
                DataHelper dh = new DataHelper();

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";

                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, "", strBeforeDate, strEndDate, strFld, "", "", "N", strSEQ, "03010204", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, "", strBeforeDate, strEndDate, strFld, "", "", "N", strSEQ, "03010204") > 0;
                sw.Stop();
                RecordSPExecuteTime("03010204", sw);
                return true;

            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }




        }

        /// <summary>
        /// VD卡人調整固定額度查詢
        /// </summary>
        /// <param name="strPeople">維護員</param>
        /// <param name="strBeforeAmount">起始固定額度</param>
        /// <param name="strEndAmount">終止固定額度</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">終止日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report03010301_BySP(string strPeople, string strBeforeAmount, string strEndAmount, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                //*宣告變數
                DataHelper dh = new DataHelper();
                String strFldName = "CREDIT LINE PERM"; //固定額度

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";

                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeAmount, strEndAmount, "N", strSEQ, "03010301", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeAmount, strEndAmount, "N", strSEQ, "03010301") > 0;
                sw.Stop();
                RecordSPExecuteTime("03010301", sw);
                return true;

            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }




        }

        /// <summary>
        /// VD卡人調整臨時額度查詢
        /// </summary>
        /// <param name="strPeople">維護員</param>
        /// <param name="strBeforeAmount">起始固定額度</param>
        /// <param name="strEndAmount">終止固定額度</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">終止日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report03010302_BySP(string strPeople, string strBeforeAmount, string strEndAmount, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                //*宣告變數
                DataHelper dh = new DataHelper();
                String strFldName = "CREDIT LINE TEMP"; //臨時額度

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";

                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeAmount, strEndAmount, "N", strSEQ, "03010302", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeAmount, strEndAmount, "N", strSEQ, "03010302") > 0;
                sw.Stop();
                RecordSPExecuteTime("03010302", sw);
                return true;

            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }




        }

        /// <summary>
        /// VD卡人新卡額度查詢
        /// </summary>
        /// <param name="strPeople">維護員</param>
        /// <param name="strBeforeAmount">起始固定額度</param>
        /// <param name="strEndAmount">終止固定額度</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">終止日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report03010303_BySP(string strPeople, string strBeforeAmount, string strEndAmount, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                //*宣告變數
                DataHelper dh = new DataHelper();
                String strFldName = "CR LINE CURR & PERM"; //新卡額度

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";

                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeAmount, strEndAmount, "N", strSEQ, "03010303", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeAmount, strEndAmount, "N", strSEQ, "03010303") > 0;
                sw.Stop();
                RecordSPExecuteTime("03010303", sw);
                return true;

            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }




        }

        /// <summary>
        /// VD卡人員工調整記錄查詢
        /// </summary>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">終止日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="blnOld">是否查詢一年半前資料</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report03010401_BySP(string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                //*宣告變數
                DataHelper dh = new DataHelper();
                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";

                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, "", strBeforeDate, strEndDate, "", "", "", "N", strSEQ, "03010401", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, "", strBeforeDate, strEndDate, "", "", "", "N", strSEQ, "03010401") > 0;
                sw.Stop();
                RecordSPExecuteTime("03010401", sw);
                return true;

            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }




        }

        /// <summary>
        /// VD卡人自扣帳戶ID與卡人ID不同者
        /// </summary>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">結束日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report03010402_BySP(string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                DataHelper dh = new DataHelper();
                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";

                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, "", strBeforeDate, strEndDate, "", "", "", "N", strSEQ, "03010402", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, "", strBeforeDate, strEndDate, "", "", "", "N", strSEQ, "03010402") > 0;
                sw.Stop();
                RecordSPExecuteTime("03010402", sw);
                return true;
            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }
        }
        #endregion

        #region VD卡片
        /// <summary>
        /// VD卡片維護記錄查詢
        /// </summary>
        /// <param name="strID">卡號</param>
        /// <param name="strPeople">維護員</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">終止日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report04010100_BySP(string strID, string strPeople, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                DataHelper dh = new DataHelper();

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";
                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                PrintStoredProcedure("SP_RptMaintainLog", strID.ToString().Trim(), strAgentID, strPeople, strBeforeDate, strEndDate, "", "", "", "N", strSEQ, "04010100", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                //20220530_Ares_Jack_新增TimeOut時間
                SqlCommand sqlcmd = new SqlCommand();
                sqlcmd.CommandText = "SP_RptMaintainLog";
                sqlcmd.CommandType = CommandType.StoredProcedure;
                sqlcmd.CommandTimeout = int.Parse(UtilHelper.GetAppSettings("PageSqlCmdTimeoutMax"));
                sqlcmd.Parameters.Add(new SqlParameter("@DB_CUST_ID", strID.ToString().Trim()));
                sqlcmd.Parameters.Add(new SqlParameter("@AGENT_ID", strAgentID));
                sqlcmd.Parameters.Add(new SqlParameter("@UPD_AGENT_ID", strPeople));
                sqlcmd.Parameters.Add(new SqlParameter("@MAINT_D_S", strBeforeDate));
                sqlcmd.Parameters.Add(new SqlParameter("@MAINT_D_E", strEndDate));
                sqlcmd.Parameters.Add(new SqlParameter("@FLD_NAME", ""));
                sqlcmd.Parameters.Add(new SqlParameter("@Value1", ""));
                sqlcmd.Parameters.Add(new SqlParameter("@Value2", ""));
                sqlcmd.Parameters.Add(new SqlParameter("@Query_History", "N"));
                sqlcmd.Parameters.Add(new SqlParameter("@SEQ_Name", strSEQ));
                sqlcmd.Parameters.Add(new SqlParameter("@strAction", "04010100"));
                blnHadRecord = dh.ExecuteNonQuery(sqlcmd) > 0;

                //blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", strID.ToString().Trim(), strAgentID, strPeople, strBeforeDate, strEndDate, "", "", "", "N", strSEQ, "04010100") > 0;
                sw.Stop();
                RecordSPExecuteTime("04010100", sw);

                return true;

            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }

        }

        /// <summary>
        /// VD卡片維護員統計表查詢
        /// </summary>
        /// <param name="strPeople">維護員</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">終止日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report04010201_BySP(string strPeople, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                DataHelper dh = new DataHelper();
                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";
                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, "", "", "", "N", strSEQ, "04010201", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, "", "", "", "N", strSEQ, "04010201") > 0;
                sw.Stop();
                RecordSPExecuteTime("04010201", sw);
                return true;

            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }




        }

        /// <summary>
        /// VD卡片>卡片與維護員關係表
        /// </summary>
        /// <param name="strID">卡號</param>
        /// <param name="strPeople">維護員</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">終止日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report04010202_BySP(string strID, string strPeople, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                DataHelper dh = new DataHelper();
                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";
                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                PrintStoredProcedure("SP_RptMaintainLog", strID.ToString().Trim(), strAgentID, strPeople, strBeforeDate, strEndDate, "", "", "", "N", strSEQ, "04010202", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", strID.ToString().Trim(), strAgentID, strPeople, strBeforeDate, strEndDate, "", "", "", "N", strSEQ, "04010202") > 0;
                sw.Stop();
                RecordSPExecuteTime("04010202", sw);

                return true;

            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }

        }

        /// <summary>
        /// VD卡片維護欄位統計表查詢
        /// </summary>
        /// <param name="strFld">維護欄位</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">終止日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report04010203_BySP(string strFld, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                DataHelper dh = new DataHelper();
                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";

                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, "", strBeforeDate, strEndDate, strFld, "", "", "N", strSEQ, "04010203", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, "", strBeforeDate, strEndDate, strFld, "", "", "N", strSEQ, "04010203") > 0;
                sw.Stop();
                RecordSPExecuteTime("04010203", sw);
                return true;

            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }

        }

        /// <summary>
        /// VD卡片調整統計表查詢

        /// </summary>
        /// <param name="strFld">維護欄位</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">終止日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report04010204_BySP(string strFld, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                DataHelper dh = new DataHelper();
                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";

                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, "", strBeforeDate, strEndDate, strFld, "", "", "N", strSEQ, "04010204", "P");
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, "", strBeforeDate, strEndDate, strFld, "", "", "N", strSEQ, "04010204") > 0;
                sw.Stop();
                RecordSPExecuteTime("04010204", sw);
                return true;
            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }




        }

        /// <summary>
        /// VD卡片調整固定額度查詢
        /// </summary>
        /// <param name="strPeople">維護員</param>
        /// <param name="strBeforeAmount">起始固定額度</param>
        /// <param name="strEndAmount">終止固定額度</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">終止日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report04010301_BySP(string strPeople, string strBeforeAmount, string strEndAmount, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                DataHelper dh = new DataHelper();
                String strFldName = "CREDIT LINE PERM"; //固定額度

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";

                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeAmount, strEndAmount, "N", strSEQ, "04010301", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeAmount, strEndAmount, "N", strSEQ, "04010301") > 0;
                sw.Stop();
                RecordSPExecuteTime("04010301", sw);
                return true;


            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }




        }

        /// <summary>
        /// VD卡片調整臨時額度查詢
        /// </summary>
        /// <param name="strPeople">維護員</param>
        /// <param name="strBeforeAmount">起始固定額度</param>
        /// <param name="strEndAmount">終止固定額度</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">終止日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report04010302_BySP(string strPeople, string strBeforeAmount, string strEndAmount, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                DataHelper dh = new DataHelper();
                String strFldName = "CREDIT LINE TEMP"; //臨時額度

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";

                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeAmount, strEndAmount, "N", strSEQ, "04010302", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeAmount, strEndAmount, "N", strSEQ, "04010302") > 0;
                sw.Stop();
                RecordSPExecuteTime("04010302", sw);
                return true;



            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }




        }

        /// <summary>
        /// VD卡片新卡額度查詢
        /// </summary>
        /// <param name="strPeople">維護員</param>
        /// <param name="strBeforeAmount">起始固定額度</param>
        /// <param name="strEndAmount">終止固定額度</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">終止日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report04010303_BySP(string strPeople, string strBeforeAmount, string strEndAmount, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                DataHelper dh = new DataHelper();
                String strFldName = "CR LINE CURR & PERM"; //新卡額度

                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";

                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeAmount, strEndAmount, "N", strSEQ, "04010303", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeAmount, strEndAmount, "N", strSEQ, "04010303") > 0;
                sw.Stop();
                RecordSPExecuteTime("04010303", sw);
                return true;

            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }
        }


        /// <summary>
        /// VD卡片非流通BlockCode調整爲流通中
        /// </summary>
        /// <param name="strPeople">維護員</param>
        /// <param name="strBeforeContent">調整前內容</param>
        /// <param name="strEndContent">調整後內容</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">終止日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report04010401_BySP(string strPeople, string strBeforeContent, string strEndContent, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                DataHelper dh = new DataHelper();
                String strFldName = "PRIMARY BLOCK CODE";
                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";

                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                if (strBeforeContent != "")
                {
                    //20211213 註解避免多一層小括號造成storedprocedure錯誤 by Ares Stanley
                    //strBeforeContent = GetRange(strBeforeContent);//*輸入了調整前內容

                }

                if (strEndContent != "")
                {
                    //20211213 註解避免多一層小括號造成storedprocedure錯誤 by Ares Stanley
                    //strEndContent = GetRange(strEndContent);//*輸入了調整後內容
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeContent, strEndContent, "N", strSEQ, "04010401", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeContent, strEndContent, "N", strSEQ, "04010401") > 0;
                sw.Stop();
                RecordSPExecuteTime("04010401", sw);
                return true;

            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }




        }


        /// <summary>
        /// VD卡片年費無優惠調整爲優惠條件
        /// </summary>
        /// <param name="strBeforeContent">調整前內容</param>
        /// <param name="strEndContent">調整後內容</param>
        /// <param name="strPeople">維護員</param>
        /// <param name="strBeforeDate">起始日期</param>
        /// <param name="strEndDate">結束日期</param>
        /// <param name="strSEQ">排序欄位</param>
        /// <param name="strMsgID">返回信息</param>
        /// <param name="rptResult">返回記錄集</param>
        /// <param name="strName">操作人員</param>
        /// <param name="blnOld">是否查詢一年半前資料</param>
        /// <param name="queryType">查詢類型S：查詢/P：列印</param>
        /// <returns>是否有記錄</returns>
        public static bool Report04010402_BySP(string strBeforeContent, string strEndContent, string strPeople, string strBeforeDate, string strEndDate, string strSEQ, string strMsgID, string strName, bool blnOld, string strAgentID, ref bool blnHadRecord, string queryType = "S")
        {
            try
            {
                DataHelper dh = new DataHelper();
                String strFldName = "USER CODE 01";
                //*維護日期起和維護日期迄
                if (strBeforeDate == "")
                {
                    strBeforeDate = "00000000";

                }
                if (strEndDate == "")
                {
                    strEndDate = "99999999";
                }

                if (strBeforeContent != "")
                {
                    //20211213 註解避免多一層小括號造成storedprocedure錯誤 by Ares Stanley
                    //strBeforeContent = GetRange(strBeforeContent);//*輸入了調整前內容

                }

                if (strEndContent != "")
                {
                    //20211213 註解避免多一層小括號造成storedprocedure錯誤 by Ares Stanley
                    //strEndContent = GetRange(strEndContent);//*輸入了調整後內容
                }

                PrintStoredProcedure("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeContent, strEndContent, "N", strSEQ, "04010402", queryType);
                Stopwatch sw = new Stopwatch();
                sw.Start();
                blnHadRecord = dh.ExecuteNonQuery("SP_RptMaintainLog", "", strAgentID, strPeople, strBeforeDate, strEndDate, strFldName, strBeforeContent, strEndContent, "N", strSEQ, "04010402") > 0;
                sw.Stop();
                RecordSPExecuteTime("04010402", sw);
                return true;

            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                return false;
            }
        }

        #endregion

        #region 公共方法
        /// <summary>
        /// 格式化ID
        /// </summary>
        /// <param name="strSource">需要轉化的ID</param>
        /// <returns></returns>
        private static string ConvertID(string strSource)
        {
            bool bYes = false;
            string l_sTemp = null;

            string l_sTemp2 = null;

            string l_sResult = null;

            bYes = false;
            switch (strSource.Length)
            {
                case 8:
                    //* 0< 數字ID < 100000000
                    //* 共8碼: 取9~16位
                    if (IsNumber(strSource))
                    {
                        l_sResult = "00000000" + strSource;
                        bYes = true;
                    }
                    break;
                case 10:
                    if (!IsNumber(strSource))
                    {
                        //*190000000000<數字ID <200000000000
                        //*共10碼
                        //*1.  第1~8碼: 取第5~12位
                        //*2.  第9~10碼:
                        //*2)  若第13或15位非0, 取第(13,14)(15,16)位分別轉成A~Z
                        if (IsNumber(strSource.Substring(0, 1)))
                        {
                            l_sResult = strSource.Substring(0, strSource.Length - 2);
                            l_sTemp = strSource.Substring(strSource.Length - 2);
                            l_sTemp = System.Convert.ToString(System.Convert.ToInt32(l_sTemp.Substring(0, 1).ToUpper()[0]) - 55) + System.Convert.ToString(System.Convert.ToInt32(l_sTemp.Substring(l_sTemp.Length - 1).ToUpper()[0]) - 55);
                            l_sResult = "0000" + l_sResult + l_sTemp;
                        }
                        //*10000000000<數字ID <36000000000
                        //*共10碼
                        //*1.  第1碼: 取第6,7位轉成A~Z
                        //*2.  第2~10碼: 取第8~16
                        else
                        {
                            l_sResult = "00000" + System.Convert.ToString(System.Convert.ToInt32(strSource.Substring(0, 1).ToUpper()[0]) - 55) + strSource.Substring(strSource.Length - (strSource.Length - 1));
                        }
                        bYes = true;
                    }
                    //*190000000000<數字ID <200000000000
                    //*共10碼
                    //*1.  第1~8碼: 取第5~12位
                    //*2.  第9~10碼:
                    //*1)  若第13及15位為0, 取第14,16位
                    else
                    {
                        l_sResult = "0000" + strSource.Substring(0, 8) + "0" + strSource.Substring(8, 1) + "0" + strSource.Substring(9, 1);
                        bYes = true;
                    }
                    break;
                case 11:
                    if (!IsNumber(strSource))
                    {
                        if (!IsNumber(strSource.Substring(1, 1)))
                        //*900000000000<數字ID <1000000000000
                        //*共11碼
                        //*1.  第1碼: 取第5位
                        //*2.  第2碼: 取第6,7位轉成A~Z
                        //*3.  第3~11碼: 取第8~16
                        {
                            l_sTemp = System.Convert.ToString(System.Convert.ToInt32(strSource.Substring(1, 1).ToUpper()[0]) - 55);
                            l_sResult = "0000" + strSource.Substring(0, 1) + l_sTemp + strSource.Substring(strSource.Length - (strSource.Length - 2));
                            bYes = true;
                        }
                        //*1000000000000<數字ID <3600000000000
                        //*共11碼
                        //*1.  第1碼: 取第4,5位轉成A~Z
                        //*2.  第2~10碼: 取第6~14位
                        //*3.  第11碼:
                        //*1)  若第15位為0, 取第16位
                        //*2)  若第15位非0, 取第15,16位轉成A~Z
                        else if (!IsNumber(strSource.Substring(0, 1)))
                        {
                            l_sTemp = System.Convert.ToString(System.Convert.ToInt32(strSource.Substring(0, 1).ToUpper()[0]) - 55);
                            l_sResult = strSource.Substring(1, 9);
                            l_sTemp2 = strSource.Substring(strSource.Length - 1);
                            if (!IsNumber(l_sTemp2))
                            {
                                l_sTemp2 = System.Convert.ToString(System.Convert.ToInt32(l_sTemp2.ToUpper()[0]) - 55);
                            }
                            else
                            {
                                l_sTemp2 = "0" + l_sTemp2;
                            }
                            l_sResult = "000" + l_sTemp + l_sResult + l_sTemp2;
                            bYes = true;
                        }
                        //*19000000000000<數字ID <20000000000000
                        //*共11碼
                        //*1.  第1~8碼: 取第3~10位
                        //*2.  第9~10碼: 取第(11,12)(13,14)位分別轉成A~Z
                        //*3.  第11碼:
                        //*1)  若第15位為0, 取第16位
                        //*2)  若第15位非0, 取第15,16位轉成A~Z
                        else if (!IsNumber(strSource.Substring(8, 1)))
                        {
                            l_sResult = strSource.Substring(0, 8);
                            l_sTemp = strSource.Substring(8, 2);
                            l_sTemp = System.Convert.ToString(System.Convert.ToInt32(l_sTemp.Substring(0, 1).ToUpper()[0]) - 55) + System.Convert.ToString(System.Convert.ToInt32(l_sTemp.Substring(l_sTemp.Length - 1).ToUpper()[0]) - 55);
                            l_sTemp2 = strSource.Substring(strSource.Length - 1);
                            if (!IsNumber(l_sTemp2))
                            {
                                l_sTemp2 = System.Convert.ToString(System.Convert.ToInt32(l_sTemp2.ToUpper()[0]) - 55);
                            }
                            else
                            {
                                l_sTemp2 = "0" + l_sTemp2;
                            }
                            l_sResult = "00" + l_sResult + l_sTemp + l_sTemp2;
                            bYes = true;
                        }
                        //*9000000000000<數字ID <10000000000000
                        //*共11碼
                        //*1.  第1碼: 取第4位
                        //*2.  第2~9碼: 取第5~12位
                        //*3.  第10~11碼:
                        //*2)  若第13或15位非0, 取第(13,14)(15,16)位分別轉成A~Z
                        else if (!IsNumber(strSource.Substring(9, 1)))
                        {
                            l_sResult = strSource.Substring(0, 9);
                            l_sTemp = strSource.Substring(strSource.Length - 2);
                            l_sTemp = System.Convert.ToString(System.Convert.ToInt32(l_sTemp.Substring(0, 1).ToUpper()[0]) - 55) + System.Convert.ToString(System.Convert.ToInt32(l_sTemp.Substring(l_sTemp.Length - 1).ToUpper()[0]) - 55);
                            l_sResult = "000" + l_sResult + l_sTemp;
                            bYes = true;
                        }
                    }
                    //*9000000000000<數字ID <10000000000000
                    //*共11碼
                    //*1.  第1碼: 取第4位
                    //*2.  第2~9碼: 取第5~12位
                    //*3.  第10~11碼:
                    //*   1)  若第13及15位為0, 取第14,16位
                    else
                    {
                        l_sResult = "000" + strSource.Substring(0, 9) + "0" + strSource.Substring(9, 1) + "0" + strSource.Substring(10, 1);
                        bYes = true;
                    }
                    break;
                case 12:
                    if (!IsNumber(strSource))
                    {
                        //*90000000000000<數字ID <100000000000000
                        //*共12碼
                        //*1.  第1碼: 取第3位
                        //*2.  第2碼: 取第4,5位轉成A~Z
                        //*3.  第3~10碼: 取第6~14位
                        //*4.  第12碼:
                        //*1)  若第15位為0, 取第16位
                        //*2)  若第15位非0, 取第15,16位轉成A~Z
                        if (!IsNumber(strSource.Substring(1, 1)))
                        {
                            l_sTemp = System.Convert.ToString(System.Convert.ToInt32(strSource.Substring(1, 1).ToUpper()[0]) - 55);
                            l_sTemp2 = strSource.Substring(strSource.Length - 1);
                            if (!IsNumber(l_sTemp2))
                            {
                                l_sTemp2 = System.Convert.ToString(System.Convert.ToInt32(l_sTemp2.ToUpper()[0]) - 55);
                            }
                            else
                            {
                                l_sTemp2 = "0" + l_sTemp2;
                            }
                            l_sResult = "00" + strSource.Substring(0, 1) + l_sTemp + strSource.Substring(2, 9) + l_sTemp2;
                            bYes = true;
                        }
                        //*900000000000000<數字ID <1000000000000000
                        //*共12碼
                        //*1.  第1碼: 取第2位
                        //*2.  第2~9碼: 取第3~10位
                        //*3.  第10~11碼: 取第(11,12)(13,14)位分別轉成A~Z
                        //*4.  第12碼:
                        //*1)  若第15位為0, 取第16位
                        //*2)  若第15位非0, 取第15,16位轉成A~Z
                        else if (!IsNumber(strSource.Substring(9, 1)))
                        {
                            l_sResult = strSource.Substring(0, 9);
                            l_sTemp = strSource.Substring(9, 2);
                            l_sTemp = System.Convert.ToString(System.Convert.ToInt32(l_sTemp.Substring(0, 1).ToUpper()[0]) - 55) + System.Convert.ToString(System.Convert.ToInt32(l_sTemp.Substring(l_sTemp.Length - 1).ToUpper()[0]) - 55);
                            l_sTemp2 = strSource.Substring(strSource.Length - 1);
                            if (!IsNumber(l_sTemp2))
                            {
                                l_sTemp2 = System.Convert.ToString(System.Convert.ToInt32(l_sTemp2.ToUpper()[0]) - 55);
                            }
                            else
                            {
                                l_sTemp2 = "0" + l_sTemp2;
                            }
                            l_sResult = "0" + l_sResult + l_sTemp + l_sTemp2;
                            bYes = true;
                        }
                    }

                    break;
                default:
                    if (strSource.Length > 16)
                    {
                        l_sResult = strSource.Substring(0, 16);
                    }
                    else
                    {
                        l_sResult = strSource;
                    }
                    bYes = true;
                    break;
            }
            if (bYes)
            {
                return l_sResult;
            }
            else
            {
                return "";
            }

        }

        /// <summary>
        /// 判斷是否為數字
        /// </summary>
        /// <param name="strNumber">傳入的字串</param>
        /// <returns></returns>
        public static bool IsNumber(String strNumber)
        {
            Regex objNotNumberPattern = new Regex("[^0-9.-]");
            return !objNotNumberPattern.IsMatch(strNumber);
        }

        /// <summary>
        /// 判斷開始日期早於18個月
        /// </summary>
        /// <param name="strBeforeDate">傳入日期</param>
        /// <returns></returns>
        public static bool IsDateTime(String strBeforeDate)
        {
            DateTime startTime = DateTime.ParseExact(strBeforeDate, "yyyyMMdd", null);
            DateTime endTime = DateTime.Now;
            int i = (endTime.Year - startTime.Year) * 12 + (endTime.Month - startTime.Month);
            if (i > 18)
            {
                return true;
            }
            return false;
        }

        /// <summary>
        /// 檢核輸入的日期
        /// </summary>
        /// <param name="dtmBeforeData">起始日期</param>
        /// <param name="dtmEndData">結束日期</param>
        /// <returns></returns>
        public static bool CheckDataTime(DateTime dtmBeforeData, DateTime dtmEndData)
        {
            try
            {
                System.TimeSpan diff1 = dtmBeforeData.Subtract(dtmEndData);
                int iDay = diff1.Days;
                if (iDay > 0)
                {
                    return false;
                }
                return true;

            }
            catch (Exception exp)
            {
                Logging.Log(exp);
                return false;
            }


        }

        /// <summary>
        /// 取得調整內容范圍
        /// </summary>
        /// <param name="strSource">需要轉化的調整內容</param>
        /// <returns></returns>
        private static string GetRange(string strSource)
        {
            int iFirst = 1;
            int iLast = 1;
            string strKey = null;
            string l_sTemp = "'";
            while (!(iLast == 0))
            {
                strKey = "";
                iLast = (strSource.IndexOf(".", iFirst - 1) + 1);
                if (iLast != 1 & iLast != strSource.Length & iLast != 0 & strSource.Trim(' ') != "")
                {
                    strKey = System.Convert.ToString(strSource.Substring(iFirst - 1, iLast - iFirst)).Trim(' ');
                    if (strKey == "%")
                    {
                        strKey = " ";
                    }

                    l_sTemp = l_sTemp + strKey + "','";
                    iLast = iLast + 1;
                    iFirst = iLast;
                }
                else if (iLast == strSource.Length && strSource.Trim(' ') != "")
                {
                    strKey = System.Convert.ToString(strSource.Substring(iFirst - 1, iLast - iFirst)).Trim(' ');
                    if (strKey == "%")
                    {
                        strKey = " ";
                    }

                    l_sTemp = l_sTemp + strKey + "'";
                    break;
                }
            }
            strKey = "";
            if (l_sTemp.Length >= 1 && iLast != strSource.Length)
            {
                if (strSource.Substring(iFirst - 1).Trim(' ') != "")
                {
                    strKey = strSource.Substring(iFirst - 1).Trim(' ');
                    if (strKey == "%")
                    {
                        strKey = " ";
                    }

                    l_sTemp = l_sTemp + strKey + "'";
                }
                else if (strSource.Substring(0).Trim(' ') != "")
                {
                    strKey = strSource.Substring(0).Trim(' ');
                    if (strKey == "%")
                    {
                        strKey = " ";
                    }

                    l_sTemp = l_sTemp + strKey + "'";
                }
                else if (l_sTemp != "'")
                {
                    strKey = System.Convert.ToString(l_sTemp.Substring(0, l_sTemp.Length - 2)).Trim(' ');
                    if (strKey == "%")
                    {
                        strKey = " ";
                    }

                    l_sTemp = strKey;
                }
            }
            return l_sTemp;
        }

        /// <summary>
        /// 作者：Ares Stanley
        /// 功能說明：印出Storedprocedure
        /// 創建日期：2022/01/28
        /// 修改紀錄：
        /// </summary>
        /// <param name="storedprocedureName">SP名稱</param>
        /// <param name="DB_CUST_ID">卡人ID</param>
        /// <param name="AGENT_ID">業務員ID</param>
        /// <param name="UPD_AGENT_ID">維護員</param>
        /// <param name="MAINT_D_S">起始日期</param>
        /// <param name="MAINT_D_E">結束日期</param>
        /// <param name="FLD_NAME">欄位名稱</param>
        /// <param name="Value1">調整前內容</param>
        /// <param name="Value2">調整後內容</param>
        /// <param name="Query_History">是否查詢一年半資料</param>
        /// <param name="SEQ_Name">排序欄位</param>
        /// <param name="strAction">報表代碼</param>
        public static void PrintStoredProcedure(string storedprocedureName, string DB_CUST_ID, string AGENT_ID, string UPD_AGENT_ID, string MAINT_D_S, string MAINT_D_E, string FLD_NAME, string Value1, string Value2, string Query_History, string SEQ_Name, string strAction, string queryType)
        {
            string strQType = string.Empty;
            string strQTable = string.Empty;
            string ADJUST_S = string.Empty;
            string ADJUST_E = string.Empty;
            string BeforeContent = string.Empty;
            string EndContent = string.Empty;
            string strInsSQL = string.Empty;
            string strSQL = string.Empty;
            bool isError = false;
            string reportName = GetReportName(strAction);
            string queryTypeName = queryType == "S" ? "查詢" : "列印";
            int stepNO = 1;
            Logging.Log($"====================報表 {reportName} ({strAction}) [{queryTypeName}]  紀錄開始====================", LogState.Info, LogLayer.Util);
            Logging.Log($"呼叫{storedprocedureName}", LogState.Info, LogLayer.Util);
            StringBuilder totalSQL = new StringBuilder();
            try
            {
                switch (strAction.Substring(0, 2))
                {
                    case "01":
                        strQType = "C";
                        strQTable = "CPMAST";
                        ADJUST_S = Value1;
                        ADJUST_E = Value2;
                        break;
                    case "02":
                        strQType = "H";
                        strQTable = "CPMAST";
                        BeforeContent = Value1;
                        EndContent = Value2;
                        break;
                    case "03":
                        strQType = "C";
                        strQTable = "CPMAST4";
                        ADJUST_S = Value1;
                        ADJUST_E = Value2;
                        break;
                    case "04":
                        strQType = "H";
                        strQTable = "CPMAST4";
                        BeforeContent = Value1;
                        EndContent = Value2;
                        break;
                    default:
                        break;
                }

                if (isError)
                {
                    Logging.Log("取得StoredProcedure發生錯誤", LogState.Info, LogLayer.Util);
                }

                strInsSQL = "INSERT INTO Rpt_CPMAST (CUST_ID ,FLD_NAME ,BEFOR_UPD ,AFTER_UPD ,MAINT_D ,MAINT_T ,USER_ID ,CSIPAgentID ,CSIPDatetime )";

                totalSQL.Append($"\n Step{stepNO} 刪除報表暫存檔：DELETE Rpt_CPMAST WHERE [CSIPAgentID] = '{AGENT_ID}' OR [CSIPDatetime] < DATEADD(day,-2,GETDATE())");

                if (strAction == "01010401" || strAction == "03010401")
                {
                    
                    totalSQL.AppendLine("新增員工對應的暫存檔：");
                    totalSQL.AppendLine($@"
                    INSERT INTO Rpt_Emp_ID ( ID, NAME, DEPRT, ID_Tmp, CSIPAgentID, CSIPDatetime )
                    SELECT
                        ID,
                        NAME,
                        DEPRT,
                        ID_Tmp,
                        '{AGENT_ID}',
                        GETDATE( ) 
                    FROM
	                    Emp_ID");

                    strSQL = $"SELECT CUST_ID,FLD_NAME,BEFOR_UPD,AFTER_UPD,MAINT_D,MAINT_T,USER_ID, '{AGENT_ID}', GETDATE()";
                    strSQL = strSQL + $" FROM {strQTable} MainMAST ,Emp_ID Emp_ID ";
                    strSQL = strSQL + $" WHERE MainMAST.CUST_ID = Emp_ID.ID_Tmp AND MainMAST.Type='{strQType}'";
                    strSQL = strSQL + $" AND (MainMAST.MAINT_D >= '{MAINT_D_S}' AND MainMAST.MAINT_D <= '{MAINT_D_E}' ";
                }
                else
                {
                    strSQL = $"SELECT CUST_ID,FLD_NAME,BEFOR_UPD,AFTER_UPD,MAINT_D,MAINT_T,USER_ID, '{AGENT_ID}', GETDATE() ";
                    strSQL = strSQL + $" FROM {strQTable} WHERE Type='{strQType}' ";
                    strSQL = strSQL + $" AND (MAINT_D >='{MAINT_D_S}' AND MAINT_D <='{MAINT_D_E}' ";
                }

                if (strAction == "01020000")
                {
                    strSQL = strSQL + " AND FLD_NAME IN ( 'DD BANK ID ACCT', 'DD MAINT DATE', 'CO SOC.SEC/TAXID FLA', 'DIRECT DEBIT ID' ) ";
                }

                if (!string.IsNullOrEmpty(DB_CUST_ID))
                {
                    strSQL = strSQL + $" AND CUST_ID='{DB_CUST_ID}' ";
                }

                if (!string.IsNullOrEmpty(UPD_AGENT_ID))
                {
                    strSQL = strSQL + $" AND User_ID ='{UPD_AGENT_ID}' ";
                }

                if (!string.IsNullOrEmpty(FLD_NAME))
                {
                    strSQL = strSQL + $" AND FLD_NAME ='{FLD_NAME}' ";
                }

                if (!string.IsNullOrEmpty(ADJUST_S) && !string.IsNullOrEmpty(ADJUST_E))
                {
                    strSQL = strSQL + $" AND CONVERT(MONEY,AFTER_UPD) >='{ADJUST_S}' AND CONVERT(MONEY,AFTER_UPD) <='{ADJUST_E}' ";
                }
                else if (!string.IsNullOrEmpty(ADJUST_S))
                {
                    strSQL = strSQL + $" AND CONVERT(MONEY,AFTER_UPD)='{ADJUST_S}' ";
                }
                else if (!string.IsNullOrEmpty(ADJUST_E))
                {
                    strSQL = strSQL + $" AND CONVERT(MONEY,AFTER_UPD)='{ADJUST_E}' ";
                }

                if (!string.IsNullOrEmpty(BeforeContent))
                {
                    strSQL = strSQL + $" AND LTRIM(BEFOR_UPD) in ('{BeforeContent}') ";
                }

                if (!string.IsNullOrEmpty(EndContent))
                {
                    strSQL = strSQL + $" AND LTRIM(AFTER_UPD) in ('{EndContent}') ";
                }

                if (Query_History == "Y")
                {
                    strSQL = strSQL + " UNION " + strSQL.Replace("CPMAST", "CPMAST_H");
                }

                strSQL = strInsSQL + $" ({strSQL})";

                if (!string.IsNullOrEmpty(SEQ_Name))
                {
                    strSQL = strSQL + $" ORDER BY '{SEQ_Name}' ";
                }

                totalSQL.AppendLine($"\n Step{stepNO+=1} 從[{strQTable}]撈取條件紀錄並全部寫入[Rpt_CPMAST] \n 執行SQL：");
                totalSQL.AppendLine(strSQL);
                Logging.Log(totalSQL.ToString(), LogState.Info, LogLayer.Util);
                if (strQType == "C")
                {
                    StringBuilder updateRpt_CPMAST = new StringBuilder();

                    string updateRpt_CPMAST_SQL = $@"
UPDATE Rpt_CPMAST SET CUST_ID=
    CASE WHEN convert(float,Cust_ID) between 0 and 100000000 THEN substring(Cust_ID,9,8)
    WHEN convert(float,Cust_ID) between 10000000000 and 36000000000 THEN char(convert(float,substring(Cust_ID,6,2))+55)+substring(Cust_ID,8,9)
    WHEN convert(float,Cust_ID) between 900000000000 and 999999999999 THEN substring(Cust_ID,5,1)+ char(convert(int,substring(Cust_ID,6,2))+55)+substring(Cust_ID,8,9)
    WHEN convert(float,Cust_ID) between 1000000000000 and 3600000000000 THEN CASE  WHEN substring(Cust_ID,15,1) = '0'  THEN char(convert(int,substring(Cust_ID,4,2))+55)+substring(Cust_ID,6,9)+substring(Cust_ID,16,1)  ELSE char(convert(int,substring(Cust_ID,4,2))+55)+substring(Cust_ID,6,9)+ char(convert(int,substring(Cust_ID,15,2))+55) End
    WHEN convert(float,Cust_ID) between 90000000000000 and 100000000000000  THEN  CASE  WHEN substring(Cust_ID,15,1) = '0' THEN substring(Cust_ID,3,1)+char(convert(int,substring(Cust_ID,4,2))+55)+substring(Cust_ID,6,9)+substring(Cust_ID,16,1) ELSE substring(Cust_ID,3,1)+char(convert(int,substring(Cust_ID,4,2))+55)+substring(Cust_ID,6,9)+ char(convert(int,substring(Cust_ID,15,2))+55)  End
    WHEN convert(float,Cust_ID) between 190000000000 and 200000000000 THEN CASE WHEN  substring(Cust_ID,13,1) = '0' and substring(Cust_ID,15,1) = '0'  THEN substring(Cust_ID,5,8)+substring(Cust_ID,14,1)+substring(Cust_ID,16,1) ELSE substring(Cust_ID,5,8)+char(convert(int,substring(Cust_ID,13,2))+55)+ char(convert(int,substring(Cust_ID,15,2))+55) End 
    WHEN convert(float,Cust_ID) between 9000000000000 and 10000000000000 THEN  CASE WHEN  substring(Cust_ID,13,1) = '0' and substring(Cust_ID,15,1) = '0'  THEN substring(Cust_ID,4,1)+substring(Cust_ID,5,8)+substring(Cust_ID,14,1)+substring(Cust_ID,16,1) ELSE substring(Cust_ID,4,1)+substring(Cust_ID,5,8)+char(convert(int,substring(Cust_ID,13,2))+55)+ char(convert(int,substring(Cust_ID,15,2))+55) End
    WHEN convert(float,Cust_ID) between 19000000000000 and 20000000000000 THEN  CASE WHEN substring(Cust_ID,15,1) = '0'  THEN substring(Cust_ID,3,8)+char(convert(int,substring(Cust_ID,11,2))+55)+char(convert(int,substring(Cust_ID,13,2))+55)+substring(Cust_ID,16,1)  ELSE substring(Cust_ID,3,8)+char(convert(int,substring(Cust_ID,11,2))+55)+char(convert(int,substring(Cust_ID,13,2))+55)+ char(convert(int,substring(Cust_ID,15,2))+55) End
    WHEN convert(float,Cust_ID) between 900000000000000 and 1000000000000000 THEN  CASE WHEN substring(Cust_ID,15,1) = '0'  THEN substring(Cust_ID,2,1)+substring(Cust_ID,3,8)+char(convert(int,substring(Cust_ID,11,2))+55)+char(convert(int,substring(Cust_ID,13,2))+55)+substring(Cust_ID,16,1)  ELSE substring(Cust_ID,2,1)+substring(Cust_ID,3,8)+char(convert(int,substring(Cust_ID,11,2))+55)+char(convert(int,substring(Cust_ID,13,2))+55)+ char(convert(int,substring(Cust_ID,15,2))+55) End
    Else Cust_ID
    End
    WHERE [CSIPAgentID] = '{AGENT_ID}' ";
                    updateRpt_CPMAST.AppendLine($"\n Step{stepNO += 1} 轉換Rpt_CPMAST的 CUST_ID：");
                    updateRpt_CPMAST.AppendLine(updateRpt_CPMAST_SQL);
                    Logging.Log(updateRpt_CPMAST.ToString(), LogState.Info, LogLayer.Util);
                }
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                Logging.Log($"====================報表 {reportName} ({strAction}) [{queryTypeName}] StoredProcedure 紀錄結束====================", LogState.Info, LogLayer.Util);
            }

        }

        /// <summary>
        /// 作者：Ares Stanley
        /// 功能說明：取得報表名稱
        /// 創建日期：2022/01/28
        /// 修改紀錄：
        /// </summary>
        /// <param name="reportId"></param>
        /// <returns></returns>
        public static string GetReportName(string reportId)
        {
            string reportName = string.Empty;
            switch (reportId)
            {
                case "01010100":
                    reportName = "信用卡卡人資料查詢列印維護紀錄查詢";
                    break;
                case "01010201":
                    reportName = "信用卡卡人資料查詢列印統計表或關係表查詢維護員統計表";
                    break;
                case "01010202":
                    reportName = "信用卡卡人資料查詢列印統計表或關係表查詢卡人與維護員關係表";
                    break;
                case "01010203":
                    reportName = "信用卡卡人資料查詢列印統計表或關係表查詢維護欄位統計表";
                    break;
                case "01010204":
                    reportName = "信用卡卡人資料查詢列印統計表或關係表查詢卡人調整統計表";
                    break;
                case "01010301":
                    reportName = "信用卡卡人資料查詢列印額度查詢調整固定額度";
                    break;
                case "01010302":
                    reportName = "信用卡卡人資料查詢列印額度查詢調整臨時額度";
                    break;
                case "01010303":
                    reportName = "信用卡卡人資料查詢列印額度查詢新卡額度";
                    break;
                case "01010401":
                    reportName = "信用卡卡人資料查詢列印特殊額度員工調整紀錄";
                    break;
                case "01010402":
                    reportName = "信用卡卡人資料查詢列印特殊額度自扣帳戶ID與卡人ID不同者";
                    break;
                case "01020000":
                    reportName = "信用卡卡人自扣資料查詢";
                    break;
                case "01030000":
                    reportName = "信用卡卡人自扣申請書歸檔查詢";
                    break;
                case "02010100":
                    reportName = "信用卡卡片資料查詢列印維護紀錄查詢";
                    break;
                case "02020100":
                    reportName = "信用卡卡片資料查詢列印統計表或關係表查詢維護員統計表";
                    break;
                case "02010202":
                    reportName = "信用卡卡片資料查詢列印統計表或關係表查詢卡片與維護員關係表";
                    break;
                case "02010203":
                    reportName = "信用卡卡片資料查詢列印統計表或關係表查詢維護欄位統計表";
                    break;
                case "02010204":
                    reportName = "信用卡卡片資料查詢列印統計表或關係表查詢卡片調整統計表";
                    break;
                case "02010301":
                    reportName = "信用卡卡片資料查詢列印額度查詢調整固定額度";
                    break;
                case "02010302":
                    reportName = "信用卡卡片資料查詢列印額度查詢調整臨時額度";
                    break;
                case "02010303":
                    reportName = "信用卡卡片資料查詢列印額度查詢新卡額度";
                    break;
                case "02010401":
                    reportName = "信用卡卡片資料查詢列印特殊額度非流通BlockCode調整為流通中";
                    break;
                case "02010402":
                    reportName = "信用卡卡片資料查詢列印特殊額度年費無優惠調整為優惠條件";
                    break;
                case "03010100":
                    reportName = "VD卡人資料查詢列印維護記錄查詢";
                    break;
                case "03010201":
                    reportName = "VD卡人資料查詢列印統計表或關係表查詢維護員統計表";
                    break;
                case "03010202":
                    reportName = "VD卡人資料查詢列印統計表或關係表查詢卡人與維護員關係表";
                    break;
                case "03010203":
                    reportName = "VD卡人資料查詢列印統計表或關係表查詢維護欄位統計表";
                    break;
                case "03010204":
                    reportName = "VD卡人資料查詢列印統計表或關係表查詢卡人調整統計表";
                    break;
                case "03010301":
                    reportName = "VD卡人資料查詢列印額度查詢調整固定額度";
                    break;
                case "03010302":
                    reportName = "VD卡人資料查詢列印額度查詢調整臨時額度";
                    break;
                case "03010303":
                    reportName = "VD卡人資料查詢列印額度查詢新卡額度";
                    break;
                case "03010401":
                    reportName = "VD卡人資料查詢列印特殊查詢員工調整記錄";
                    break;
                case "03010402":
                    reportName = "VD卡人資料查詢列印特殊查詢自扣帳戶ID與卡人ID不同者";
                    break;
                case "04010100":
                    reportName = "VD卡片資料查詢列印維護記錄查詢";
                    break;
                case "04010201":
                    reportName = "VD卡片資料查詢列印統計表或關係表查詢維護員統計表";
                    break;
                case "04010202":
                    reportName = "VD卡片資料查詢列印統計表或關係表查詢卡片與維護員關係表";
                    break;
                case "04010203":
                    reportName = "VD卡片資料查詢列印統計表或關係表查詢維護欄位統計表";
                    break;
                case "04010204":
                    reportName = "VD卡片資料查詢列印統計表或關係表查詢卡片調整統計表";
                    break;
                case "04010301":
                    reportName = "VD卡片資料查詢列印額度查詢調整固定額度";
                    break;
                case "04010302":
                    reportName = "VD卡片資料查詢列印額度查詢調整臨時額度";
                    break;
                case "04010303":
                    reportName = "VD卡片資料查詢列印額度查詢新卡額度";
                    break;
                case "04010401":
                    reportName = "VD卡片資料查詢列印特殊額度非流通BlockCode調整為流通中";
                    break;
                case "04010402":
                    reportName = "VD卡片資料查詢列印特殊額度年費無優惠調整為優惠條件";
                    break;
                case "05010000":
                    reportName = "匯入紀錄查詢";
                    break;
                case "05010000_Detail":
                    reportName = "匯入紀錄查詢明細";
                    break;
                default:
                    reportName = reportId;
                    break;

            }
            return reportName;
        }

        /// <summary>
        /// 作者：Ares Stanley
        /// 功能說明：紀錄SP查詢耗時
        /// 創建日期：2022/01/28
        /// 修改紀錄：
        /// </summary>
        /// <param name="strAction"></param>
        /// <param name="sw"></param>
        public static void RecordSPExecuteTime(string strAction, Stopwatch sw)
        {
            try
            {
                string reportName = GetReportName(strAction);
                TimeSpan ts = sw.Elapsed;
                Logging.Log(string.Format("SP執行時間共：{0} ms", ts.TotalMilliseconds), LogState.Info, LogLayer.Util);
            }
            catch (Exception ex)
            {
                Logging.Log("記錄SP執行時間時發生錯誤：" + ex.ToString());
            }
        }

        /// <summary>
        /// 作者：Ares Stanley
        /// 功能說明：紀錄SQL查詢結果筆數
        /// 創建日期：2022/01/28
        /// 修改紀錄：
        /// </summary>
        /// <param name="strAction">報表代碼</param>
        /// <param name="dt">資料表</param>
        public static void RecordSQLDataCount(string strAction, DataTable dt)
        {
            try
            {
                string reportName = GetReportName(strAction);

                if (dt == null)
                {
                    return;
                }

                Logging.Log(string.Format("{0}：資料結果共{1}筆", reportName, dt.Rows.Count));
            }
            catch (Exception ex)
            {
                Logging.Log("記錄SQL結果筆數時發生錯誤：" + ex.ToString());
            }
        }
        #endregion

        #endregion
    }

}
