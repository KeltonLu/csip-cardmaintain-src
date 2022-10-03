//******************************************************************
//*  作    者：Ares Stanley
//*  功能說明：報表查詢、產出
//*  創建日期：2021/11/08
//*  修改記錄：
//*<author>            <time>            <TaskID>                <desc>
//*Ares Stanley　　 2022/05/11　20210058-CSIP作業服務平台現代化II　調整信用卡卡人、卡片、VD卡人、卡片之維護紀錄查詢超過6萬筆時，以CSV產出
//*******************************************************************
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using Framework.Common.Logging;
using Framework.Common.Utility;
using System.Data.SqlClient;
using Framework.Data;


namespace CSIPCardMaintain.BusinessRules
{
    public class BR_Excel_File
    {
        #region SQL Command

        public const string sqlComm_member = @"
        SELECT
	        CPMAST.BEFOR_UPD,
	        CPMAST.AFTER_UPD,
	        CPMAST.MAINT_D,
	        CPMAST.MAINT_T,
	        CPMAST.USER_ID,
	        CPMAST.FLD_NAME,
	        Emp_ID.name,
	        Emp_ID.Deprt,
	        Emp_ID.ID 
        FROM
	        [{0}].[dbo].Rpt_Emp_ID Emp_ID
	    INNER JOIN [{0}].[dbo].Rpt_CPMAST CPMAST ON Emp_ID.ID_Tmp = CPMAST.CUST_ID
        WHERE CPMAST.CSIPAgentID = @CSIPAgentID";

        public const string sqlComm_member1 = @"
        SELECT
	        CPMAST.CUST_ID,
	        CPMAST.FLD_NAME,
	        CPMAST.BEFOR_UPD,
	        CPMAST.AFTER_UPD,
	        CPMAST.MAINT_D,
	        CPMAST.MAINT_T,
	        CPMAST.USER_ID 
        FROM
	        [{0}].[dbo].Rpt_CPMAST CPMAST
	        WHERE CPMAST.CSIPAgentID = @CSIPAgentID";


        public const string sqlComm_fld = @"
        SELECT
	        CPMAST.FLD_NAME,
	        CPMAST.USER_ID,
	        CPMAST.MAINT_D 
        FROM
	        [{0}].[dbo].[Rpt_CPMAST] CPMAST
        WHERE
	        CPMAST.CSIPAgentID = @CSIPAgentID";

        public const string sqlComm_user = @"
        SELECT
	        CPMAST.USER_ID,
	        CPMAST.FLD_NAME 
        FROM
	        [{0}].[dbo].Rpt_CPMAST CPMAST
        WHERE
	        CPMAST.CSIPAgentID = @CSIPAgentID";

        public const string sqlComm_card_r = @"
        SELECT
	        CPMAST.CUST_ID,
	        CPMAST.USER_ID 
        FROM
	        [{0}].[dbo].Rpt_CPMAST CPMAST
        WHERE
	        CPMAST.CSIPAgentID = @CSIPAgentID";

        public const string sqlComm_card = @"
        SELECT
	        CPMAST.CUST_ID,
	        CPMAST.FLD_NAME 
        FROM
	        [{0}].[dbo].Rpt_CPMAST CPMAST
        WHERE
	        CPMAST.CSIPAgentID = @CSIPAgentID";

        /// <summary>
        /// 維護員統計表
        /// 共用：01010201、02020100、03010201、04010201
        /// </summary>
        public const string sqlComm_01010201 = @"
        SELECT
	        CPMAST.USER_ID,
	        CPMAST.FLD_NAME ,
	        COUNT(CPMAST.FLD_NAME) AS SUBTOTAL
        FROM
	        [{0}].[dbo].Rpt_CPMAST CPMAST 
	        WHERE CSIPAgentID = @CSIPAgentID
        GROUP BY
	        USER_ID, FLD_NAME";

        /// <summary>
        /// 卡人與維護員關係表
        /// 共用：01010202、02010202、03010202、04010202
        /// </summary>
        public const string sqlComm_01010202 = @"
        SELECT
	        '' AS TOTAL,
	        CPMAST.CUST_ID,
	        CPMAST.USER_ID,
	        COUNT ( CPMAST.USER_ID ) AS SUBTOTAL 
        FROM
	        [{0}].[dbo].Rpt_CPMAST CPMAST 
        WHERE
	        CSIPAgentID = @CSIPAgentID 
        GROUP BY
	        CPMAST.CUST_ID,
	        CPMAST.USER_ID";

        /// <summary>
        /// 卡人與維護關係表(群組數)
        /// 共用：01010202、02010202、03010202、04010202
        /// </summary>
        public const string sqlComm_01010202_Count = @"
        SELECT COUNT
	        ( DISTINCT CUST_ID ) 
        FROM
	        Rpt_CPMAST 
        WHERE
	        CSIPAgentID = @CSIPAgentID";

        /// <summary>
        /// 維護欄位統計表
        /// 共用：01010203、02010203、03010203、04010203
        /// </summary>
        public const string sqlComm_01010203 = @"
        SELECT
	        CPMAST.FLD_NAME,
	        CPMAST.USER_ID,
	        COUNT ( CPMAST.USER_ID ) AS SUBTOTAL
        FROM
	        [{0}].[dbo].Rpt_CPMAST CPMAST 
        WHERE
	        CSIPAgentID = @CSIPAgentID 
        GROUP BY
	        CPMAST.FLD_NAME,
	        CPMAST.USER_ID";

        /// <summary>
        /// 卡人、卡片調整統計表
        /// 共用：01010204、02010204、03010204、04010204
        /// </summary>
        public const string sqlComm_01010204 = @"
        SELECT
            '' AS TOTAL,
	        CPMAST.CUST_ID,
	        CPMAST.FLD_NAME,
	        COUNT ( CPMAST.FLD_NAME ) AS SUBTOTAL 
        FROM
	        [{0}].dbo.Rpt_CPMAST CPMAST 
        WHERE
	        CSIPAgentID = @CSIPAgentID 
        GROUP BY
	        CPMAST.CUST_ID,
	        CPMAST.FLD_NAME";

        /// <summary>
        /// 卡人、卡片調整統計表(群組數)
        /// 共用01010204、02010204、03010204、04010204
        /// </summary>
        public const string sqlComm_01010204_Count = @"
        SELECT COUNT
	        ( DISTINCT CUST_ID ) 
        FROM
	        Rpt_CPMAST 
        WHERE
	        CSIPAgentID = @CSIPAgentID";

        /// <summary>
        /// 匯入紀錄查詢(有輸入日期)
        /// </summary>
        public const string sqlComm_05010000_withCondition = @"select * from Import_Log where  UPPER(INDate)  >= @dateStart AND  UPPER(INDate)  <= @dateEnd and ( (FILENAME LIKE 'os06%') or (FILENAME LIKE 'ts06%') ) ORDER BY INDate DESC";

        /// <summary>
        /// 匯入紀錄查詢(無輸入日期)
        /// </summary>
        public const string sqlComm_05010000_withoutCondition = @"select * from Import_Log where ( (FILENAME LIKE 'os06%') or (FILENAME LIKE 'ts06%') ) ORDER BY INDate DESC";

        /// <summary>
        /// 匯入紀錄明細查詢
        /// </summary>
        public const string sqlComm_05010000Detail = @"SELECT * FROM {0} WHERE EXE_Name = @EXE_Name";

        #endregion SQL Command

        #region getData
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:取報表範本member1資料
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/10
        /// </summary>
        /// <returns></returns>
        public static DataTable getData_member1(string agentId, ref int totalCount, string strAction = "", int iPageIndex = 0, string queryType = "S", bool isOrderBy = false)
        {
            string connection = UtilHelper.GetConnectionStrings("Connection_System");
            string queryTypeName = queryType == "S" ? "查詢" : "列印";
            SqlConnection sql_conn = new SqlConnection(connection);
            DataTable dt = new DataTable();
            try
            {
                SqlCommand sqlComm = new SqlCommand();
                sqlComm.CommandText = string.Format(sqlComm_member1, UtilHelper.GetAppSettings("DB_CP_DBF"));
                sqlComm.Parameters.Add(new SqlParameter("@CSIPAgentID", agentId));

                //紀錄SQL
                Dictionary<string, string> commandParameters = new Dictionary<string, string>();
                commandParameters.Add("@CSIPAgentID", agentId);

                //Get report name
                string reportName = string.Empty;
                if (!string.IsNullOrEmpty(strAction))
                {
                    reportName = BRReport.GetReportName(strAction);
                }
                Stopwatch sw = new Stopwatch();
                DataHelper dh = new DataHelper();
                sw.Start();
                DataSet ds = dh.ExecuteDataSet(sqlComm);
                sw.Stop();
                if (ds.Tables.Count > 0)
                    dt = ds.Tables[0];

                if (dt != null)
                {
                    totalCount = dt.Rows.Count;
                }
                else
                {
                    totalCount = 0;
                }

                PrintSQL(sqlComm.CommandText, sw, commandParameters, reportName, dt, queryTypeName);

                if (isOrderBy)
                {
                    DataView dv = dt.DefaultView;
                    dv.Sort = "MAINT_D DESC";
                    dt = dv.ToTable();
                }

                // 判斷頁次//20220629_Ares_Jack_調整翻頁BUG
                for (int i = 0; i < int.Parse(UtilHelper.GetAppSettings("PageSize")) * (iPageIndex - 1); i++)
                {
                    dt.Rows.Remove(dt.Rows[0]);
                }

                return dt;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                sql_conn.Close();
                return dt;
            }
            finally
            {

            }
        }
        

        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:取報表範本member資料
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/10
        /// </summary>
        /// <returns></returns>
        public static DataTable getData_member(string agentId, ref int totalCount, string strAction = "", int iPageIndex = 0, string queryType = "S")
        {
            string connection = UtilHelper.GetConnectionStrings("Connection_System");
            string queryTypeName = queryType == "S" ? "查詢" : "列印";
            SqlConnection sql_conn = new SqlConnection(connection);
            DataTable dt = new DataTable();
            try
            {
                SqlCommand sqlComm = new SqlCommand();
                sqlComm.CommandText = string.Format(sqlComm_member, UtilHelper.GetAppSettings("DB_CP_DBF"));
                sqlComm.Parameters.Add(new SqlParameter("@CSIPAgentID", agentId));

                //Get report name
                string reportName = string.Empty;
                if (!string.IsNullOrEmpty(strAction))
                {
                    reportName = BRReport.GetReportName(strAction);
                }

                //紀錄SQL
                Dictionary<string, string> commandParameters = new Dictionary<string, string>();
                commandParameters.Add("@CSIPAgentID", agentId);

                DataHelper dh = new DataHelper();
                Stopwatch sw = new Stopwatch();
                sw.Start();
                DataSet ds = dh.ExecuteDataSet(sqlComm);
                sw.Stop();
                if (ds.Tables.Count > 0)
                    dt = ds.Tables[0];

                if (dt != null)
                {
                    totalCount = dt.Rows.Count;
                }
                else
                {
                    totalCount = 0;
                }

                PrintSQL(sqlComm.CommandText, sw, commandParameters, reportName, dt, queryTypeName);

                // 判斷頁次
                for (int i = 0; i < 10 * (iPageIndex - 1); i++)
                {
                    dt.Rows.Remove(dt.Rows[0]);
                }

                return dt;

            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                sql_conn.Close();
                return dt;
            }
            finally
            {

            }
        }

        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:取報表範本user資料
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/19
        /// </summary>
        /// <returns></returns>
        public static DataTable getData_user(string agentId, ref int totalCount, string strAction = "", int iPageIndex = 0, string queryType = "S")
        {
            string connection = UtilHelper.GetConnectionStrings("Connection_System");
            string queryTypeName = queryType == "S" ? "查詢" : "列印";
            SqlConnection sql_conn = new SqlConnection(connection);
            DataTable dt = new DataTable();
            try
            {
                SqlCommand sqlComm = new SqlCommand();
                sqlComm.CommandText = string.Format(sqlComm_user, UtilHelper.GetAppSettings("DB_CP_DBF"));
                sqlComm.Parameters.Add(new SqlParameter("@CSIPAgentID", agentId));

                //Get report name
                string reportName = string.Empty;
                if (!string.IsNullOrEmpty(strAction))
                {
                    reportName = BRReport.GetReportName(strAction);
                }

                //紀錄SQL
                Dictionary<string, string> commandParameters = new Dictionary<string, string>();
                commandParameters.Add("@CSIPAgentID", agentId);

                DataHelper dh = new DataHelper();
                Stopwatch sw = new Stopwatch();
                sw.Start();
                DataSet ds = dh.ExecuteDataSet(sqlComm);
                sw.Stop();
                if (ds.Tables.Count > 0)
                    dt = ds.Tables[0];

                if (dt != null)
                {
                    totalCount = dt.Rows.Count;
                }
                else
                {
                    totalCount = 0;
                }

                PrintSQL(sqlComm.CommandText, sw, commandParameters, reportName, dt, queryTypeName);

                // 判斷頁次
                for (int i = 0; i < 10 * (iPageIndex - 1); i++)
                {
                    dt.Rows.Remove(dt.Rows[0]);
                }

                return dt;

            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                sql_conn.Close();
                return dt;
            }
            finally
            {

            }
        }

        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:共用取得資料
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/19
        /// </summary>
        /// <returns></returns>
        public static DataTable getData_Comm(string agentId, string sqlCommText, ref int totalCount, string strAction = "", int iPageIndex = 0, bool isPrintSQL = true, string queryType = "S")
        {
            string connection = UtilHelper.GetConnectionStrings("Connection_System");
            string queryTypeName = queryType == "S" ? "查詢" : "列印";
            SqlConnection sql_conn = new SqlConnection(connection);
            DataTable dt = new DataTable();
            try
            {
                SqlCommand sqlComm = new SqlCommand();
                sqlComm.CommandText = sqlCommText;
                sqlComm.Parameters.Add(new SqlParameter("@CSIPAgentID", agentId));

                //Get report name
                string reportName = string.Empty;
                if (!string.IsNullOrEmpty(strAction))
                {
                    reportName = BRReport.GetReportName(strAction);
                }

                //紀錄SQL
                Dictionary<string, string> commandParameters = new Dictionary<string, string>();
                commandParameters.Add("@CSIPAgentID", agentId);

                DataHelper dh = new DataHelper();
                Stopwatch sw = new Stopwatch();
                sw.Start();
                DataSet ds = dh.ExecuteDataSet(sqlComm);
                sw.Stop();
                if (ds.Tables.Count > 0)
                    dt = ds.Tables[0];

                if (dt != null)
                {
                    totalCount = dt.Rows.Count;
                }
                else
                {
                    totalCount = 0;
                }

                if (isPrintSQL)
                {
                    PrintSQL(sqlComm.CommandText, sw, commandParameters, reportName, dt, queryTypeName);
                }

                // 判斷頁次
                for (int i = 0; i < 10 * (iPageIndex - 1); i++)
                {
                    dt.Rows.Remove(dt.Rows[0]);
                }

                return dt;

            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                sql_conn.Close();
                return dt;
            }
            finally
            {

            }
        }
        #endregion

        #region 信用卡卡人

        #region Report01010100 信用卡卡人-維護資料查詢

        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Excel_Report01010100(卡人/維護資料查詢)
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/10
        /// </summary>
        /// <param name="strRptID"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="strAgentName"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report01010100(string strRptID, string strRptBeforeDate, string strRptEndDate, string strAgentName, string agentId, ref string strPathFile, ref string strMsgID, ref bool isCSV)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "member1.xls";
                int totalCount = 0;
                DataTable dt = getData_member1(agentId, ref totalCount, "01010100", 0, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                if (dt.Rows.Count < 60000)
                {
                    //資料少於6萬筆，以Excel產出
                    FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                    HSSFWorkbook wb = new HSSFWorkbook(fs);
                    ISheet sheet = wb.GetSheet("工作表1");
                    #region 表頭
                    sheet.GetRow(0).GetCell(0).SetCellValue("維護記錄查詢");
                    sheet.GetRow(2).GetCell(0).SetCellValue(string.Format("卡人:{0}", strRptID));
                    sheet.GetRow(3).GetCell(0).SetCellValue(string.Format("維護日期 : {0} ~ {1}", strRptBeforeDate, strRptEndDate));
                    sheet.GetRow(3).GetCell(6).SetCellValue(DateTime.Now.ToString("yyyyMMdd"));//製表日
                    sheet.GetRow(4).GetCell(6).SetCellValue(strAgentName);//製表人
                    sheet.GetRow(7).GetCell(0).SetCellValue("卡人ID");
                    #endregion

                    //取得樣式
                    HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        sheet.CreateRow(sheet.LastRowNum + 1);
                        for (int b = 0; b < 7; b++)
                        {
                            sheet.GetRow(sheet.LastRowNum).CreateCell(b);
                            sheet.GetRow(sheet.LastRowNum).GetCell(b).CellStyle = contentFormat;
                        }
                        sheet.GetRow(sheet.LastRowNum).GetCell(0).SetCellValue(dt.Rows[i]["CUST_ID"].ToString().Trim());
                        sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue(dt.Rows[i]["FLD_NAME"].ToString().Trim());
                        sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(dt.Rows[i]["BEFOR_UPD"].ToString().Trim());
                        sheet.GetRow(sheet.LastRowNum).GetCell(3).SetCellValue(dt.Rows[i]["AFTER_UPD"].ToString().Trim());
                        sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(dt.Rows[i]["MAINT_D"].ToString().Trim());
                        sheet.GetRow(sheet.LastRowNum).GetCell(5).SetCellValue(dt.Rows[i]["MAINT_T"].ToString().Trim());
                        sheet.GetRow(sheet.LastRowNum).GetCell(6).SetCellValue(dt.Rows[i]["USER_ID"].ToString().Trim());
                    }
                    // 保存文件到運行目錄下
                    strPathFile = strPathFile + @"\ExcelFile_Report01010100" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                    FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                    wb.Write(fs1);
                    fs1.Close();
                    fs.Close();
                }
                else
                {
                    //資料多於6萬筆，以CSV產出
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine("卡人ID,欄位名稱,調整前內容,調整後內容,維護日期,維護時間,維護員");
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        sb.AppendLine(
                            string.Format("=\"{0}\"", dt.Rows[i]["CUST_ID"].ToString().Trim()) + "," +
                            string.Format("=\"{0}\"", dt.Rows[i]["FLD_NAME"].ToString().Trim()) + "," +
                            string.Format("=\"{0}\"", dt.Rows[i]["BEFOR_UPD"].ToString().Trim().Replace(",", "")) + "," +
                            string.Format("=\"{0}\"", dt.Rows[i]["AFTER_UPD"].ToString().Trim().Replace(",", "")) + "," +
                            string.Format("=\"{0}\"", dt.Rows[i]["MAINT_D"].ToString().Trim()) + "," +
                            string.Format("=\"{0}\"", dt.Rows[i]["MAINT_T"].ToString().Trim()) + "," +
                            string.Format("=\"{0}\"", dt.Rows[i]["USER_ID"].ToString().Trim())
                            );
                    }
                    strPathFile = strPathFile + @"\ExcelFile_Report01010100" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";
                    File.WriteAllText(strPathFile, sb.ToString(), Encoding.Default);
                    isCSV = true;
                }

                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }


        }
        #endregion

        #region Report01010301 信用卡卡人-調整固定額度
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Excel_Report01010301(卡人/額度查詢/調整固定額度)
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/10
        /// </summary>
        /// <param name="strRptPeople"></param>
        /// <param name="strRptBeforeAmount"></param>
        /// <param name="strRptEndAmount"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="strAgentName"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report01010301(string strRptPeople, string strRptBeforeAmount, string strRptEndAmount, string strRptBeforeDate, string strRptEndDate, string strAgentName, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "member1.xls";
                int totalCount = 0;
                DataTable dt = getData_member1(agentId, ref totalCount, "01010301", 0, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");
                #region 表頭
                sheet.GetRow(0).GetCell(0).SetCellValue("調整固定額度");
                sheet.GetRow(2).GetCell(0).SetCellValue(string.Format("維護員 : {0}", strRptPeople));
                sheet.GetRow(3).GetCell(0).SetCellValue(string.Format("額度 : {0} ~ {1}", strRptBeforeAmount, strRptEndAmount));
                sheet.GetRow(3).GetCell(6).SetCellValue(DateTime.Now.ToString("yyyyMMdd"));//製表日
                sheet.GetRow(4).GetCell(0).SetCellValue(string.Format("維護日期 : {0} ~ {1}", strRptBeforeDate, strRptEndDate));
                sheet.GetRow(4).GetCell(6).SetCellValue(strAgentName);//製表人
                sheet.GetRow(7).GetCell(0).SetCellValue("卡人ID");
                #endregion

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 資料
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sheet.CreateRow(sheet.LastRowNum + 1);
                    for (int b = 0; b < 7; b++)
                    {
                        sheet.GetRow(sheet.LastRowNum).CreateCell(b);
                        sheet.GetRow(sheet.LastRowNum).GetCell(b).CellStyle = contentFormat;
                    }
                    sheet.GetRow(sheet.LastRowNum).GetCell(0).SetCellValue(dt.Rows[i]["CUST_ID"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue(dt.Rows[i]["FLD_NAME"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(dt.Rows[i]["BEFOR_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(3).SetCellValue(dt.Rows[i]["AFTER_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(dt.Rows[i]["MAINT_D"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(5).SetCellValue(dt.Rows[i]["MAINT_T"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(6).SetCellValue(dt.Rows[i]["USER_ID"].ToString().Trim());
                }
                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report01010301" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }
        }
        #endregion

        #region Report01010302 信用卡卡人-調整臨時額度
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Excel_Report01010302(卡人/額度查詢/調整臨時額度)
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/10
        /// </summary>
        /// <param name="strRptPeople"></param>
        /// <param name="strRptBeforeAmount"></param>
        /// <param name="strRptEndAmount"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="strAgentName"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report01010302(string strRptPeople, string strRptBeforeAmount, string strRptEndAmount, string strRptBeforeDate, string strRptEndDate, string strAgentName, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "member1.xls";
                int totalCount = 0;
                DataTable dt = getData_member1(agentId, ref totalCount, "01010302", 0, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");
                #region 表頭
                sheet.GetRow(0).GetCell(0).SetCellValue("調整臨時額度");
                sheet.GetRow(2).GetCell(0).SetCellValue(string.Format("維護員 : {0}", strRptPeople));
                sheet.GetRow(3).GetCell(0).SetCellValue(string.Format("額度 : {0} ~ {1}", strRptBeforeAmount, strRptEndAmount));
                sheet.GetRow(3).GetCell(6).SetCellValue(DateTime.Now.ToString("yyyyMMdd"));//製表日
                sheet.GetRow(4).GetCell(0).SetCellValue(string.Format("維護日期 : {0} ~ {1}", strRptBeforeDate, strRptEndDate));
                sheet.GetRow(4).GetCell(6).SetCellValue(strAgentName);//製表人
                sheet.GetRow(7).GetCell(0).SetCellValue("卡人ID");
                #endregion

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 資料
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sheet.CreateRow(sheet.LastRowNum + 1);
                    for (int b = 0; b < 7; b++)
                    {
                        sheet.GetRow(sheet.LastRowNum).CreateCell(b);
                        sheet.GetRow(sheet.LastRowNum).GetCell(b).CellStyle = contentFormat;
                    }
                    sheet.GetRow(sheet.LastRowNum).GetCell(0).SetCellValue(dt.Rows[i]["CUST_ID"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue(dt.Rows[i]["FLD_NAME"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(dt.Rows[i]["BEFOR_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(3).SetCellValue(dt.Rows[i]["AFTER_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(dt.Rows[i]["MAINT_D"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(5).SetCellValue(dt.Rows[i]["MAINT_T"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(6).SetCellValue(dt.Rows[i]["USER_ID"].ToString().Trim());
                }
                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report01010302" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }
        }
        #endregion

        #region Report01010303 信用卡卡人-新卡額度
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Excel_Report01010303(卡人/額度查詢/新卡額度)
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/10
        /// </summary>
        /// <param name="strRptPeople"></param>
        /// <param name="strRptBeforeAmount"></param>
        /// <param name="strRptEndAmount"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="strAgentName"></param>
        /// <param name="agentId"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report01010303(string strRptPeople, string strRptBeforeAmount, string strRptEndAmount, string strRptBeforeDate, string strRptEndDate, string strAgentName, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "member1.xls";
                int totalCount = 0;
                DataTable dt = getData_member1(agentId, ref totalCount, "01010303", 0, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");
                #region 表頭
                sheet.GetRow(0).GetCell(0).SetCellValue("新卡額度");
                sheet.GetRow(2).GetCell(0).SetCellValue(string.Format("維護員 : {0}", strRptPeople));
                sheet.GetRow(3).GetCell(0).SetCellValue(string.Format("額度 : {0} ~ {1}", strRptBeforeAmount, strRptEndAmount));
                sheet.GetRow(3).GetCell(6).SetCellValue(DateTime.Now.ToString("yyyyMMdd"));//製表日
                sheet.GetRow(4).GetCell(0).SetCellValue(string.Format("維護日期 : {0} ~ {1}", strRptBeforeDate, strRptEndDate));
                sheet.GetRow(4).GetCell(6).SetCellValue(strAgentName);//製表人
                sheet.GetRow(7).GetCell(0).SetCellValue("卡人ID");
                #endregion

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 資料
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sheet.CreateRow(sheet.LastRowNum + 1);
                    for (int b = 0; b < 7; b++)
                    {
                        sheet.GetRow(sheet.LastRowNum).CreateCell(b);
                        sheet.GetRow(sheet.LastRowNum).GetCell(b).CellStyle = contentFormat;
                    }
                    sheet.GetRow(sheet.LastRowNum).GetCell(0).SetCellValue(dt.Rows[i]["CUST_ID"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue(dt.Rows[i]["FLD_NAME"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(dt.Rows[i]["BEFOR_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(3).SetCellValue(dt.Rows[i]["AFTER_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(dt.Rows[i]["MAINT_D"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(5).SetCellValue(dt.Rows[i]["MAINT_T"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(6).SetCellValue(dt.Rows[i]["USER_ID"].ToString().Trim());
                }
                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report01010303" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }
        }
        #endregion

        #region Report01010401 信用卡卡人-員工調整紀錄
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Excel_Report01010401(卡人/特殊查詢/員工調整紀錄)
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/10
        /// </summary>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="strAgentName"></param>
        /// <param name="agentId"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report01010401(string strRptBeforeDate, string strRptEndDate, string strAgentName, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "member.xls";
                int totalCount = 0;
                DataTable dt = getData_member(agentId, ref totalCount, "01010401", 0, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");
                #region 表頭
                sheet.GetRow(0).GetCell(0).SetCellValue("員工調整記錄");
                sheet.GetRow(3).GetCell(6).SetCellValue(DateTime.Now.ToString("yyyyMMdd"));//製表日
                sheet.GetRow(3).GetCell(0).SetCellValue(string.Format("維護日期 : {0} ~ {1}", strRptBeforeDate, strRptEndDate));
                sheet.GetRow(4).GetCell(6).SetCellValue(strAgentName);//製表人
                sheet.GetRow(7).GetCell(0).SetCellValue("卡人ID");
                #endregion

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 資料
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sheet.CreateRow(sheet.LastRowNum + 1);
                    for (int b = 0; b < 7; b++)
                    {
                        sheet.GetRow(sheet.LastRowNum).CreateCell(b);
                        sheet.GetRow(sheet.LastRowNum).GetCell(b).CellStyle = contentFormat;
                    }
                    sheet.GetRow(sheet.LastRowNum).GetCell(0).SetCellValue(dt.Rows[i]["ID"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue(dt.Rows[i]["FLD_NAME"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(dt.Rows[i]["BEFOR_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(3).SetCellValue(dt.Rows[i]["AFTER_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(dt.Rows[i]["MAINT_D"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(5).SetCellValue(dt.Rows[i]["MAINT_T"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(6).SetCellValue(dt.Rows[i]["USER_ID"].ToString().Trim());
                }
                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report01010401" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }
        }
        #endregion

        #region Report01010402 信用卡卡人-自扣帳戶ID與卡人ID不同者
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Excel_Report01010402(卡人/特殊查詢/自扣帳戶ID與卡人ID不同者)
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/10
        /// </summary>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="strAgentName"></param>
        /// <param name="agentId"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report01010402(string strRptBeforeDate, string strRptEndDate, string strAgentName, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "member1.xls";
                int totalCount = 0;
                DataTable dt = getData_member1(agentId, ref totalCount, "01010402", 0, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");
                #region 表頭
                sheet.GetRow(0).GetCell(0).SetCellValue("自扣ID與卡人ID不同者");
                sheet.GetRow(3).GetCell(0).SetCellValue("維護欄位 : DIRECT DEBIT ID");
                sheet.GetRow(3).GetCell(6).SetCellValue(DateTime.Now.ToString("yyyyMMdd"));//製表日
                sheet.GetRow(4).GetCell(0).SetCellValue(string.Format("維護日期 : {0} ~ {1}", strRptBeforeDate, strRptEndDate));
                sheet.GetRow(4).GetCell(6).SetCellValue(strAgentName);//製表人
                sheet.GetRow(7).GetCell(0).SetCellValue("卡人ID");
                #endregion

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 資料
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sheet.CreateRow(sheet.LastRowNum + 1);
                    for (int b = 0; b < 7; b++)
                    {
                        sheet.GetRow(sheet.LastRowNum).CreateCell(b);
                        sheet.GetRow(sheet.LastRowNum).GetCell(b).CellStyle = contentFormat;
                    }
                    sheet.GetRow(sheet.LastRowNum).GetCell(0).SetCellValue(dt.Rows[i]["CUST_ID"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue(dt.Rows[i]["FLD_NAME"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(dt.Rows[i]["BEFOR_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(3).SetCellValue(dt.Rows[i]["AFTER_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(dt.Rows[i]["MAINT_D"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(5).SetCellValue(dt.Rows[i]["MAINT_T"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(6).SetCellValue(dt.Rows[i]["USER_ID"].ToString().Trim());
                }
                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report01010402" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }
        }
        #endregion

        #region Report01020000 信用卡卡人-自扣資料查詢
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Excel_Report01020000(卡人/自扣資料查詢)
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/10
        /// </summary>
        /// <param name="strRptID"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="strAgentName"></param>
        /// <param name="agentId"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report01020000(string strRptID, string strRptBeforeDate, string strRptEndDate, string strAgentName, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "member1.xls";
                int totalCount = 0;
                DataTable dt = getData_member1(agentId, ref totalCount, "01020000", 0, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");
                #region 表頭
                sheet.GetRow(0).GetCell(0).SetCellValue("自扣資料查詢");
                sheet.GetRow(2).GetCell(0).SetCellValue(string.Format("卡人:{0}", strRptID));
                sheet.GetRow(3).GetCell(6).SetCellValue(DateTime.Now.ToString("yyyyMMdd"));//製表日
                sheet.GetRow(3).GetCell(0).SetCellValue(string.Format("維護日期 : {0} ~ {1}", strRptBeforeDate, strRptEndDate));
                sheet.GetRow(4).GetCell(6).SetCellValue(strAgentName);//製表人
                sheet.GetRow(7).GetCell(0).SetCellValue("卡人ID");
                #endregion

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 資料
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sheet.CreateRow(sheet.LastRowNum + 1);
                    for (int b = 0; b < 7; b++)
                    {
                        sheet.GetRow(sheet.LastRowNum).CreateCell(b);
                        sheet.GetRow(sheet.LastRowNum).GetCell(b).CellStyle = contentFormat;
                    }
                    sheet.GetRow(sheet.LastRowNum).GetCell(0).SetCellValue(dt.Rows[i]["CUST_ID"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue(dt.Rows[i]["FLD_NAME"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(dt.Rows[i]["BEFOR_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(3).SetCellValue(dt.Rows[i]["AFTER_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(dt.Rows[i]["MAINT_D"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(5).SetCellValue(dt.Rows[i]["MAINT_T"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(6).SetCellValue(dt.Rows[i]["USER_ID"].ToString().Trim());
                }
                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report01020000" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }

        }
        #endregion

        #region Report01010201 信用卡卡人-維護員統計表
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Excel_Report01010201(卡人/統計表或關係表查詢/維護員統計表)
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/10
        /// </summary>
        /// <param name="strRptID"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="strAgentName"></param>
        /// <param name="agentId"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report01010201(string strName, string strRptPeople, string strRptBeforeDate, string strRptEndDate, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "user.xls";
                int totalCount = 0;
                DataTable dt = getData_Comm(agentId, string.Format(sqlComm_01010201, UtilHelper.GetAppSettings("DB_CP_DBF")), ref totalCount, "01010201", 0, true, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 表頭
                sheet.GetRow(4).GetCell(0).SetCellValue("維護員：" + strRptPeople);//維護員
                sheet.GetRow(5).GetCell(0).SetCellValue("維護日期：" + strRptBeforeDate + " ~ " + strRptEndDate);//維護日期
                sheet.GetRow(4).GetCell(6).SetCellValue("製表日：" + DateTime.Now.ToString("yyyy/MM/dd"));//製表日
                sheet.GetRow(5).GetCell(6).SetCellValue("製表人：" + strName);//製表人

                #endregion


                #region 表身

                //資料去空白
                removeBlank(ref dt);

                ExportExcelForNPOI_SubTotal(dt, ref wb, 7, "工作表1");

                #region 合併相同維護員資料
                int startRow = 0;
                int endRow = 0;
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    if (row == sheet.LastRowNum)
                    {
                        if (startRow > 0)
                        {
                            endRow = row;
                        }
                        if (endRow - startRow >= 1)
                        {
                            sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                            sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                            //建立小計
                            NPOI_AddSubTotal(sheet, startRow, endRow, contentFormat);
                            startRow = 0;
                            endRow = 0;
                        }
                        break;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() == sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString())
                    {
                        if (startRow != 0)
                            continue;
                        startRow = row;
                        continue;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() != sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString() || (row == sheet.LastRowNum - 1 && startRow > 0))
                    {
                        if (startRow != 0)
                        {
                            endRow = row;

                            if (endRow - startRow >= 1)
                            {
                                sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                                sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                                // 建立小計
                                NPOI_AddSubTotal(sheet, startRow, endRow, contentFormat);
                                startRow = 0;
                                endRow = 0;
                                continue;
                            }
                        }

                        if (startRow == 0 && sheet.GetRow(row).GetCell(3).StringCellValue.ToString() != "小計")
                        {
                            // 建立小計
                            NPOI_AddSubTotal(sheet, row, row, contentFormat);
                        }
                    }
                }

                //尾列總計
                int sumValue = 0;
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    if (sheet.GetRow(row).GetCell(3).StringCellValue == "小計")
                    {
                        int result = 0;
                        bool tryParse = int.TryParse(sheet.GetRow(row).GetCell(4).NumericCellValue.ToString(), out result);
                        if (tryParse)
                        {
                            sumValue += result;
                        }
                    }
                }
                sheet.CreateRow(sheet.LastRowNum + 1);
                for (int col = 2; col < 5; col++)
                {
                    sheet.GetRow(sheet.LastRowNum).CreateCell(col);
                    sheet.GetRow(sheet.LastRowNum).GetCell(col).CellStyle = contentFormat;
                }
                sheet.AddMergedRegion(new CellRangeAddress(sheet.LastRowNum, sheet.LastRowNum, 2, 3));
                sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue("總計");
                sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(sumValue);
                #endregion

                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report01010201" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }

        }
        #endregion

        #region Report01010202 信用卡卡人-卡人與維護員關係表
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Excel_Report01010202(卡人/統計表或關係表查詢/卡人與維護員關係表)
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/30
        /// </summary>
        /// <param name="strName"></param>
        /// <param name="strRptPeople"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="agentId"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report01010202(string strName, string strRptPeople, string strRptBeforeDate, string strRptEndDate, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "card_r.xls";
                int totalCount = 0;
                DataTable dt = getData_Comm(agentId, string.Format(sqlComm_01010202, UtilHelper.GetAppSettings("DB_CP_DBF")), ref totalCount, "01010202", 0, true, "P");
                if (dt.Rows.Count <= 0)
                    return false;
                DataTable dt2 = getData_Comm(agentId, sqlComm_01010202_Count, ref totalCount, "01010202差異總計", 0, false);
                string totalDiffCount = "";
                if (dt2.Rows.Count > 0)
                {
                    totalDiffCount = dt2.Rows[0][0].ToString();
                }
                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 表頭
                sheet.GetRow(4).GetCell(0).SetCellValue("維護員：" + strRptPeople);//維護員
                sheet.GetRow(5).GetCell(0).SetCellValue("維護日期：" + strRptBeforeDate + " ~ " + strRptEndDate);//維護日期
                sheet.GetRow(4).GetCell(5).SetCellValue("製表日：" + DateTime.Now.ToString("yyyy/MM/dd"));//製表日
                sheet.GetRow(5).GetCell(5).SetCellValue("製表人：" + strName);//製表人

                #endregion


                #region 表身

                //資料去空白
                removeBlank(ref dt);

                //資料寫入
                ExportExcelForNPOI(dt, ref wb, 7, "工作表1", 1);

                #region 合併相同維護員資料
                int startRow = 0;
                int endRow = 0;
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    if (row == sheet.LastRowNum)
                    {
                        if (startRow > 0)
                        {
                            endRow = row;
                        }
                        if (endRow - startRow >= 1)
                        {
                            sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 1, 1));
                            sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                            sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                            startRow = 0;
                            endRow = 0;
                        }
                        break;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() == sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString())
                    {
                        if (startRow != 0)
                            continue;
                        startRow = row;
                        continue;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() != sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString() || (row == sheet.LastRowNum - 1 && startRow > 0))
                    {
                        if (startRow != 0)
                        {
                            endRow = row;

                            if (endRow - startRow >= 1)
                            {
                                sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 1, 1));
                                sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                                sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                                startRow = 0;
                                endRow = 0;
                                continue;
                            }
                        }
                    }
                }
                #endregion

                //增加尾列總計
                int sumValue = NPOI_ColumnSum(sheet, 7, sheet.LastRowNum);
                sheet.CreateRow(sheet.LastRowNum + 1);
                for (int c = 1; c < 6; c++)
                {
                    sheet.GetRow(sheet.LastRowNum).CreateCell(c);
                    sheet.GetRow(sheet.LastRowNum).GetCell(c).CellStyle = contentFormat;
                }
                sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue("總計");
                sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(totalDiffCount);
                sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(sumValue);

                //小計欄位合併
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    sheet.GetRow(row).CreateCell(5);
                    sheet.GetRow(row).GetCell(5).CellStyle = contentFormat;
                }
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    sheet.AddMergedRegion(new CellRangeAddress(row, row, 4, 5));
                    sheet.GetRow(row).GetCell(4).CellStyle = contentFormat;
                }
                //額外合併尾列維護員、小計
                sheet.AddMergedRegion(new CellRangeAddress(sheet.LastRowNum, sheet.LastRowNum, 3, 4));
                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report01010202" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }
        }
        #endregion

        #region Report01010203 信用卡卡人-維護欄位統計表
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Excel_Report01010203(卡人/統計表或關係表查詢/維護欄位統計表)
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/30
        /// </summary>
        /// <param name="strName"></param>
        /// <param name="strRptFld"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="agentId"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report01010203(string strName, string strRptFld, string strRptBeforeDate, string strRptEndDate, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "fld.xls";
                int totalCount = 0;
                DataTable dt = getData_Comm(agentId, string.Format(sqlComm_01010203, UtilHelper.GetAppSettings("DB_CP_DBF")), ref totalCount, "01010203", 0, true, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 表頭
                sheet.GetRow(4).GetCell(0).SetCellValue("維護欄位：" + strRptFld);//維護欄位
                sheet.GetRow(5).GetCell(0).SetCellValue("維護日期：" + strRptBeforeDate + " ~ " + strRptEndDate);//維護日期
                sheet.GetRow(4).GetCell(6).SetCellValue("製表日：" + DateTime.Now.ToString("yyyy/MM/dd"));//製表日
                sheet.GetRow(5).GetCell(6).SetCellValue("製表人：" + strName);//製表人

                #endregion


                #region 表身

                //資料去空白
                removeBlank(ref dt);

                ExportExcelForNPOI_SubTotal(dt, ref wb, 7, "工作表1");

                #region 合併相同維護員資料
                int startRow = 0;
                int endRow = 0;
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    if (row == sheet.LastRowNum)
                    {
                        if (startRow > 0)
                        {
                            endRow = row;
                        }
                        if (endRow - startRow >= 1)
                        {
                            sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                            sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                            //建立小計
                            NPOI_AddSubTotal(sheet, startRow, endRow, contentFormat);
                            startRow = 0;
                            endRow = 0;
                        }
                        break;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() == sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString())
                    {
                        if (startRow != 0)
                            continue;
                        startRow = row;
                        continue;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() != sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString() || (row == sheet.LastRowNum - 1 && startRow > 0))
                    {
                        if (startRow != 0)
                        {
                            endRow = row;

                            if (endRow - startRow >= 1)
                            {
                                sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                                sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                                // 建立小計
                                NPOI_AddSubTotal(sheet, startRow, endRow, contentFormat);
                                startRow = 0;
                                endRow = 0;
                                continue;
                            }
                        }

                        if (startRow == 0 && sheet.GetRow(row).GetCell(3).StringCellValue.ToString() != "小計")
                        {
                            // 建立小計
                            NPOI_AddSubTotal(sheet, row, row, contentFormat);
                        }
                    }
                }

                //尾列總計
                int sumValue = 0;
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    if (sheet.GetRow(row).GetCell(3).StringCellValue == "小計")
                    {
                        int result = 0;
                        bool tryParse = int.TryParse(sheet.GetRow(row).GetCell(4).NumericCellValue.ToString(), out result);
                        if (tryParse)
                        {
                            sumValue += result;
                        }
                    }
                }
                sheet.CreateRow(sheet.LastRowNum + 1);
                for (int col = 2; col < 5; col++)
                {
                    sheet.GetRow(sheet.LastRowNum).CreateCell(col);
                    sheet.GetRow(sheet.LastRowNum).GetCell(col).CellStyle = contentFormat;
                }
                sheet.AddMergedRegion(new CellRangeAddress(sheet.LastRowNum, sheet.LastRowNum, 2, 3));
                sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue("總計");
                sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(sumValue);
                #endregion

                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report01010203" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }

        }
        #endregion

        #region Report01010204 信用卡卡人-卡人調整統計表
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Excel_Report01010204(卡人/統計表或關係表查詢/卡人調整統計表)
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/30
        /// </summary>
        /// <param name="strName"></param>
        /// <param name="strRptFld"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="agentId"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report01010204(string strName, string strRptFld, string strRptBeforeDate, string strRptEndDate, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "card.xls";
                int totalCount = 0;
                DataTable dt = getData_Comm(agentId, string.Format(sqlComm_01010204, UtilHelper.GetAppSettings("DB_CP_DBF")), ref totalCount, "01010204", 0, true, "P");
                if (dt.Rows.Count <= 0)
                    return false;
                DataTable dt2 = getData_Comm(agentId, sqlComm_01010204_Count, ref totalCount, "01010204差異總計", 0, false);
                string totalDiffCount = "";
                if (dt2.Rows.Count > 0)
                {
                    totalDiffCount = dt2.Rows[0][0].ToString();
                }
                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 表頭
                sheet.GetRow(4).GetCell(0).SetCellValue("維護欄位：" + strRptFld);//維護欄位
                sheet.GetRow(5).GetCell(0).SetCellValue("維護日期：" + strRptBeforeDate + " ~ " + strRptEndDate);//維護日期
                sheet.GetRow(4).GetCell(5).SetCellValue("製表日：" + DateTime.Now.ToString("yyyy/MM/dd"));//製表日
                sheet.GetRow(5).GetCell(5).SetCellValue("製表人：" + strName);//製表人

                #endregion


                #region 表身

                //資料去空白
                removeBlank(ref dt);

                //資料寫入
                ExportExcelForNPOI(dt, ref wb, 7, "工作表1", 1);

                #region 合併相同資料
                int startRow = 0;
                int endRow = 0;
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    if (row == sheet.LastRowNum)
                    {
                        if (startRow > 0)
                        {
                            endRow = row;
                        }
                        if (endRow - startRow >= 1)
                        {
                            sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 1, 1));
                            sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                            sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                            startRow = 0;
                            endRow = 0;
                        }
                        break;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() == sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString())
                    {
                        if (startRow != 0)
                            continue;
                        startRow = row;
                        continue;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() != sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString() || (row == sheet.LastRowNum - 1 && startRow > 0))
                    {
                        if (startRow != 0)
                        {
                            endRow = row;

                            if (endRow - startRow >= 1)
                            {
                                sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 1, 1));
                                sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                                sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                                startRow = 0;
                                endRow = 0;
                                continue;
                            }
                        }
                    }
                }
                #endregion

                //增加尾列總計
                int sumValue = NPOI_ColumnSum(sheet, 7, sheet.LastRowNum);
                sheet.CreateRow(sheet.LastRowNum + 1);
                for (int c = 1; c < 6; c++)
                {
                    sheet.GetRow(sheet.LastRowNum).CreateCell(c);
                    sheet.GetRow(sheet.LastRowNum).GetCell(c).CellStyle = contentFormat;
                }
                sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue("總計");
                sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(totalDiffCount);
                sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(sumValue);

                //小計欄位合併
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    sheet.GetRow(row).CreateCell(5);
                    sheet.GetRow(row).GetCell(5).CellStyle = contentFormat;
                }
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    sheet.AddMergedRegion(new CellRangeAddress(row, row, 4, 5));
                    sheet.GetRow(row).GetCell(4).CellStyle = contentFormat;
                }
                //額外合併尾列維護員、小計
                sheet.AddMergedRegion(new CellRangeAddress(sheet.LastRowNum, sheet.LastRowNum, 3, 4));
                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report01010204" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }
        }



        #endregion

        #endregion

        #region 信用卡卡片

        #region Report02010100 信用卡卡片-維護資料查詢

        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Report02010100(卡片/維護資料查詢)資料並產出Excel
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/10
        /// </summary>
        /// <param name="strRptID"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="strAgentName"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report02010100(string strRptID, string strRptBeforeDate, string strRptEndDate, string strAgentName, string agentId, ref string strPathFile, ref string strMsgID, ref bool isCSV)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "member1.xls";
                int totalCount = 0;
                DataTable dt = getData_member1(agentId, ref totalCount, "02010100", 0, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                if (dt.Rows.Count < 60000)
                {
                    //資料少於6萬筆，以Excel產出
                    FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                    HSSFWorkbook wb = new HSSFWorkbook(fs);
                    ISheet sheet = wb.GetSheet("工作表1");

                    //取得樣式
                    HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                    #region 表頭
                    sheet.GetRow(0).GetCell(0).SetCellValue("維護記錄查詢");
                    sheet.GetRow(2).GetCell(0).SetCellValue(string.Format("卡片:{0}", strRptID));
                    sheet.GetRow(3).GetCell(0).SetCellValue(string.Format("維護日期 : {0} ~ {1}", strRptBeforeDate, strRptEndDate));
                    sheet.GetRow(3).GetCell(6).SetCellValue(DateTime.Now.ToString("yyyyMMdd"));//製表日
                    sheet.GetRow(4).GetCell(6).SetCellValue(strAgentName);//製表人
                    sheet.GetRow(7).GetCell(0).SetCellValue("卡號");
                    #endregion

                    #region 資料
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        sheet.CreateRow(sheet.LastRowNum + 1);
                        for (int b = 0; b < 7; b++)
                        {
                            sheet.GetRow(sheet.LastRowNum).CreateCell(b);
                            sheet.GetRow(sheet.LastRowNum).GetCell(b).CellStyle = contentFormat;
                        }
                        sheet.GetRow(sheet.LastRowNum).GetCell(0).SetCellValue(dt.Rows[i]["CUST_ID"].ToString().Trim());
                        sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue(dt.Rows[i]["FLD_NAME"].ToString().Trim());
                        sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(dt.Rows[i]["BEFOR_UPD"].ToString().Trim());
                        sheet.GetRow(sheet.LastRowNum).GetCell(3).SetCellValue(dt.Rows[i]["AFTER_UPD"].ToString().Trim());
                        sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(dt.Rows[i]["MAINT_D"].ToString().Trim());
                        sheet.GetRow(sheet.LastRowNum).GetCell(5).SetCellValue(dt.Rows[i]["MAINT_T"].ToString().Trim());
                        sheet.GetRow(sheet.LastRowNum).GetCell(6).SetCellValue(dt.Rows[i]["USER_ID"].ToString().Trim());
                    }
                    #endregion

                    // 保存文件到運行目錄下
                    strPathFile = strPathFile + @"\ExcelFile_Report02010100" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                    FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                    wb.Write(fs1);
                    fs1.Close();
                    fs.Close();
                }
                else
                {
                    //資料多於6萬筆，以CSV產出
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine("卡號,欄位名稱,調整前內容,調整後內容,維護日期,維護時間,維護員");
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        sb.AppendLine(
                            string.Format("=\"{0}\"", dt.Rows[i]["CUST_ID"].ToString().Trim()) + "," +
                            string.Format("=\"{0}\"", dt.Rows[i]["FLD_NAME"].ToString().Trim()) + "," +
                            string.Format("=\"{0}\"", dt.Rows[i]["BEFOR_UPD"].ToString().Trim().Replace(",", "")) + "," +
                            string.Format("=\"{0}\"", dt.Rows[i]["AFTER_UPD"].ToString().Trim().Replace(",", "")) + "," +
                            string.Format("=\"{0}\"", dt.Rows[i]["MAINT_D"].ToString().Trim()) + "," +
                            string.Format("=\"{0}\"", dt.Rows[i]["MAINT_T"].ToString().Trim()) + "," +
                            string.Format("=\"{0}\"", dt.Rows[i]["USER_ID"].ToString().Trim())
                            );
                    }
                    strPathFile = strPathFile + @"\ExcelFile_Report02010100" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";
                    File.WriteAllText(strPathFile, sb.ToString(), Encoding.Default);
                    isCSV = true;
                }


                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }


        }
        #endregion

        #region Report02020100 信用卡卡片-維護員統計表
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Excel_Report02020100(卡片/統計表或關係表查詢/維護員統計表)
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/10
        /// </summary>
        /// <param name="strName"></param>
        /// <param name="strRptPeople"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="agentId"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report02020100(string strName, string strRptPeople, string strRptBeforeDate, string strRptEndDate, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "user.xls";
                int totalCount = 0;
                DataTable dt = getData_Comm(agentId, string.Format(sqlComm_01010201, UtilHelper.GetAppSettings("DB_CP_DBF")), ref totalCount, "02020100", 0, true, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 表頭
                sheet.GetRow(4).GetCell(0).SetCellValue("維護員：" + strRptPeople);//維護員
                sheet.GetRow(5).GetCell(0).SetCellValue("維護日期：" + strRptBeforeDate + " ~ " + strRptEndDate);//維護日期
                sheet.GetRow(4).GetCell(6).SetCellValue("製表日：" + DateTime.Now.ToString("yyyy/MM/dd"));//製表日
                sheet.GetRow(5).GetCell(6).SetCellValue("製表人：" + strName);//製表人

                #endregion


                #region 表身

                //資料去空白
                removeBlank(ref dt);

                ExportExcelForNPOI_SubTotal(dt, ref wb, 7, "工作表1");

                #region 合併相同維護員資料
                int startRow = 0;
                int endRow = 0;
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    if (row == sheet.LastRowNum)
                    {
                        if (startRow > 0)
                        {
                            endRow = row;
                        }
                        if (endRow - startRow >= 1)
                        {
                            sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                            sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                            //建立小計
                            NPOI_AddSubTotal(sheet, startRow, endRow, contentFormat);
                            startRow = 0;
                            endRow = 0;
                        }
                        break;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() == sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString())
                    {
                        if (startRow != 0)
                            continue;
                        startRow = row;
                        continue;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() != sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString() || (row == sheet.LastRowNum - 1 && startRow > 0))
                    {
                        if (startRow != 0)
                        {
                            endRow = row;

                            if (endRow - startRow >= 1)
                            {
                                sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                                sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                                // 建立小計
                                NPOI_AddSubTotal(sheet, startRow, endRow, contentFormat);
                                startRow = 0;
                                endRow = 0;
                                continue;
                            }
                        }

                        if (startRow == 0 && sheet.GetRow(row).GetCell(3).StringCellValue.ToString() != "小計")
                        {
                            // 建立小計
                            NPOI_AddSubTotal(sheet, row, row, contentFormat);
                        }
                    }
                }

                //尾列總計
                int sumValue = 0;
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    if (sheet.GetRow(row).GetCell(3).StringCellValue == "小計")
                    {
                        int result = 0;
                        bool tryParse = int.TryParse(sheet.GetRow(row).GetCell(4).NumericCellValue.ToString(), out result);
                        if (tryParse)
                        {
                            sumValue += result;
                        }
                    }
                }
                sheet.CreateRow(sheet.LastRowNum + 1);
                for (int col = 2; col < 5; col++)
                {
                    sheet.GetRow(sheet.LastRowNum).CreateCell(col);
                    sheet.GetRow(sheet.LastRowNum).GetCell(col).CellStyle = contentFormat;
                }
                sheet.AddMergedRegion(new CellRangeAddress(sheet.LastRowNum, sheet.LastRowNum, 2, 3));
                sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue("總計");
                sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(sumValue);
                #endregion

                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report02020100" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }

        }
        #endregion

        #region Report02010202 信用卡卡片-卡人與維護員關係表

        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Excel_Report02010202(卡片/統計表或關係表查詢/卡人與維護員關係表)
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/30
        /// </summary>
        /// <param name="strName"></param>
        /// <param name="strRptPeople"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="agentId"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report02010202(string strName, string strRptPeople, string strRptBeforeDate, string strRptEndDate, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "card_r.xls";
                int totalCount = 0;
                DataTable dt = getData_Comm(agentId, string.Format(sqlComm_01010202, UtilHelper.GetAppSettings("DB_CP_DBF")), ref totalCount, "02010202", 0, true, "P");
                if (dt.Rows.Count <= 0)
                    return false;
                DataTable dt2 = getData_Comm(agentId, sqlComm_01010202_Count, ref totalCount, "02010202差異總計", 0, false);
                string totalDiffCount = "";
                if (dt2.Rows.Count > 0)
                {
                    totalDiffCount = dt2.Rows[0][0].ToString();
                }
                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 表頭
                sheet.GetRow(1).GetCell(1).SetCellValue("卡片與維護員關係表");
                sheet.GetRow(4).GetCell(0).SetCellValue("維護員：" + strRptPeople);//維護員
                sheet.GetRow(5).GetCell(0).SetCellValue("維護日期：" + strRptBeforeDate + " ~ " + strRptEndDate);//維護日期
                sheet.GetRow(4).GetCell(5).SetCellValue("製表日：" + DateTime.Now.ToString("yyyy/MM/dd"));//製表日
                sheet.GetRow(5).GetCell(5).SetCellValue("製表人：" + strName);//製表人
                sheet.GetRow(6).GetCell(2).SetCellValue("卡號");
                #endregion


                #region 表身

                //資料去空白
                removeBlank(ref dt);

                //資料寫入
                ExportExcelForNPOI(dt, ref wb, 7, "工作表1", 1);

                #region 合併相同維護員資料
                int startRow = 0;
                int endRow = 0;
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    if (row == sheet.LastRowNum)
                    {
                        if (startRow > 0)
                        {
                            endRow = row;
                        }
                        if (endRow - startRow >= 1)
                        {
                            sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 1, 1));
                            sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                            sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                            startRow = 0;
                            endRow = 0;
                        }
                        break;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() == sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString())
                    {
                        if (startRow != 0)
                            continue;
                        startRow = row;
                        continue;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() != sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString() || (row == sheet.LastRowNum - 1 && startRow > 0))
                    {
                        if (startRow != 0)
                        {
                            endRow = row;

                            if (endRow - startRow >= 1)
                            {
                                sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 1, 1));
                                sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                                sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                                startRow = 0;
                                endRow = 0;
                                continue;
                            }
                        }
                    }
                }
                #endregion

                //增加尾列總計
                int sumValue = NPOI_ColumnSum(sheet, 7, sheet.LastRowNum);
                sheet.CreateRow(sheet.LastRowNum + 1);
                for (int c = 1; c < 6; c++)
                {
                    sheet.GetRow(sheet.LastRowNum).CreateCell(c);
                    sheet.GetRow(sheet.LastRowNum).GetCell(c).CellStyle = contentFormat;
                }
                sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue("總計");
                sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(totalDiffCount);
                sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(sumValue);

                //小計欄位合併
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    sheet.GetRow(row).CreateCell(5);
                    sheet.GetRow(row).GetCell(5).CellStyle = contentFormat;
                }
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    sheet.AddMergedRegion(new CellRangeAddress(row, row, 4, 5));
                    sheet.GetRow(row).GetCell(4).CellStyle = contentFormat;
                }
                //額外合併尾列維護員、小計
                sheet.AddMergedRegion(new CellRangeAddress(sheet.LastRowNum, sheet.LastRowNum, 3, 4));
                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report02010202" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }
        }
        #endregion

        #region Report02010203 信用卡卡片-維護欄位統計表
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Excel_Report02010203(卡片/統計表或關係表查詢/維護欄位統計表)
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/30
        /// </summary>
        /// <param name="strName"></param>
        /// <param name="strRptFld"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="agentId"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report02010203(string strName, string strRptFld, string strRptBeforeDate, string strRptEndDate, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "fld.xls";
                int totalCount = 0;
                DataTable dt = getData_Comm(agentId, string.Format(sqlComm_01010203, UtilHelper.GetAppSettings("DB_CP_DBF")), ref totalCount, "02010203", 0, true, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 表頭
                sheet.GetRow(4).GetCell(0).SetCellValue("維護欄位：" + strRptFld);//維護欄位
                sheet.GetRow(5).GetCell(0).SetCellValue("維護日期：" + strRptBeforeDate + " ~ " + strRptEndDate);//維護日期
                sheet.GetRow(4).GetCell(6).SetCellValue("製表日：" + DateTime.Now.ToString("yyyy/MM/dd"));//製表日
                sheet.GetRow(5).GetCell(6).SetCellValue("製表人：" + strName);//製表人

                #endregion


                #region 表身

                //資料去空白
                removeBlank(ref dt);

                ExportExcelForNPOI_SubTotal(dt, ref wb, 7, "工作表1");

                #region 合併相同維護員資料
                int startRow = 0;
                int endRow = 0;
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    if (row == sheet.LastRowNum)
                    {
                        if (startRow > 0)
                        {
                            endRow = row;
                        }
                        if (endRow - startRow >= 1)
                        {
                            sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                            sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                            //建立小計
                            NPOI_AddSubTotal(sheet, startRow, endRow, contentFormat);
                            startRow = 0;
                            endRow = 0;
                        }
                        break;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() == sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString())
                    {
                        if (startRow != 0)
                            continue;
                        startRow = row;
                        continue;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() != sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString() || (row == sheet.LastRowNum - 1 && startRow > 0))
                    {
                        if (startRow != 0)
                        {
                            endRow = row;

                            if (endRow - startRow >= 1)
                            {
                                sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                                sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                                // 建立小計
                                NPOI_AddSubTotal(sheet, startRow, endRow, contentFormat);
                                startRow = 0;
                                endRow = 0;
                                continue;
                            }
                        }

                        if (startRow == 0 && sheet.GetRow(row).GetCell(3).StringCellValue.ToString() != "小計")
                        {
                            // 建立小計
                            NPOI_AddSubTotal(sheet, row, row, contentFormat);
                        }
                    }
                }

                //尾列總計
                int sumValue = 0;
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    if (sheet.GetRow(row).GetCell(3).StringCellValue == "小計")
                    {
                        int result = 0;
                        bool tryParse = int.TryParse(sheet.GetRow(row).GetCell(4).NumericCellValue.ToString(), out result);
                        if (tryParse)
                        {
                            sumValue += result;
                        }
                    }
                }
                sheet.CreateRow(sheet.LastRowNum + 1);
                for (int col = 2; col < 5; col++)
                {
                    sheet.GetRow(sheet.LastRowNum).CreateCell(col);
                    sheet.GetRow(sheet.LastRowNum).GetCell(col).CellStyle = contentFormat;
                }
                sheet.AddMergedRegion(new CellRangeAddress(sheet.LastRowNum, sheet.LastRowNum, 2, 3));
                sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue("總計");
                sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(sumValue);
                #endregion

                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report02010203" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }

        }
        #endregion

        #region Report02010204 信用卡卡片-卡片調整統計表
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Excel_Report02010204(卡片/統計表或關係表查詢/卡片調整統計表)
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/30
        /// </summary>
        /// <param name="strName"></param>
        /// <param name="strRptFld"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="agentId"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report02010204(string strName, string strRptFld, string strRptBeforeDate, string strRptEndDate, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "card.xls";
                int totalCount = 0;
                DataTable dt = getData_Comm(agentId, string.Format(sqlComm_01010204, UtilHelper.GetAppSettings("DB_CP_DBF")), ref totalCount, "02010204", 0, true, "P");
                if (dt.Rows.Count <= 0)
                    return false;
                DataTable dt2 = getData_Comm(agentId, sqlComm_01010204_Count, ref totalCount, "02010204差異總計", 0, false);
                string totalDiffCount = "";
                if (dt2.Rows.Count > 0)
                {
                    totalDiffCount = dt2.Rows[0][0].ToString();
                }
                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 表頭
                sheet.GetRow(1).GetCell(1).SetCellValue("卡片調整統計表");
                sheet.GetRow(4).GetCell(0).SetCellValue("維護欄位：" + strRptFld);//維護欄位
                sheet.GetRow(5).GetCell(0).SetCellValue("維護日期：" + strRptBeforeDate + " ~ " + strRptEndDate);//維護日期
                sheet.GetRow(4).GetCell(5).SetCellValue("製表日：" + DateTime.Now.ToString("yyyy/MM/dd"));//製表日
                sheet.GetRow(5).GetCell(5).SetCellValue("製表人：" + strName);//製表人
                sheet.GetRow(6).GetCell(2).SetCellValue("卡號");
                #endregion


                #region 表身

                //資料去空白
                removeBlank(ref dt);

                //資料寫入
                ExportExcelForNPOI(dt, ref wb, 7, "工作表1", 1);

                #region 合併相同資料
                int startRow = 0;
                int endRow = 0;
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    if (row == sheet.LastRowNum)
                    {
                        if (startRow > 0)
                        {
                            endRow = row;
                        }
                        if (endRow - startRow >= 1)
                        {
                            sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 1, 1));
                            sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                            sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                            startRow = 0;
                            endRow = 0;
                        }
                        break;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() == sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString())
                    {
                        if (startRow != 0)
                            continue;
                        startRow = row;
                        continue;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() != sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString() || (row == sheet.LastRowNum - 1 && startRow > 0))
                    {
                        if (startRow != 0)
                        {
                            endRow = row;

                            if (endRow - startRow >= 1)
                            {
                                sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 1, 1));
                                sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                                sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                                startRow = 0;
                                endRow = 0;
                                continue;
                            }
                        }
                    }
                }
                #endregion

                //增加尾列總計
                int sumValue = NPOI_ColumnSum(sheet, 7, sheet.LastRowNum);
                sheet.CreateRow(sheet.LastRowNum + 1);
                for (int c = 1; c < 6; c++)
                {
                    sheet.GetRow(sheet.LastRowNum).CreateCell(c);
                    sheet.GetRow(sheet.LastRowNum).GetCell(c).CellStyle = contentFormat;
                }
                sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue("總計");
                sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(totalDiffCount);
                sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(sumValue);

                //小計欄位合併
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    sheet.GetRow(row).CreateCell(5);
                    sheet.GetRow(row).GetCell(5).CellStyle = contentFormat;
                }
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    sheet.AddMergedRegion(new CellRangeAddress(row, row, 4, 5));
                    sheet.GetRow(row).GetCell(4).CellStyle = contentFormat;
                }
                //額外合併尾列維護員、小計
                sheet.AddMergedRegion(new CellRangeAddress(sheet.LastRowNum, sheet.LastRowNum, 3, 4));
                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report02010204" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }
        }



        #endregion

        #region Report02010301 信用卡卡片-調整固定額度
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Report02010301(卡片/額度查詢/調整固定額度)資料並產出Excel
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/10
        /// </summary>
        /// <param name="strRptPeople"></param>
        /// <param name="strRptBeforeAmount"></param>
        /// <param name="strRptEndAmount"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="strAgentName"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report02010301(string strRptPeople, string strRptBeforeAmount, string strRptEndAmount, string strRptBeforeDate, string strRptEndDate, string strAgentName, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "member1.xls";
                int totalCount = 0;
                DataTable dt = getData_member1(agentId, ref totalCount, "02010301", 0, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");
                #region 表頭
                sheet.GetRow(0).GetCell(0).SetCellValue("調整固定額度");
                sheet.GetRow(2).GetCell(0).SetCellValue(string.Format("維護員 : {0}", strRptPeople));
                sheet.GetRow(3).GetCell(0).SetCellValue(string.Format("額度 : {0} ~ {1}", strRptBeforeAmount, strRptEndAmount));
                sheet.GetRow(3).GetCell(6).SetCellValue(DateTime.Now.ToString("yyyyMMdd"));//製表日
                sheet.GetRow(4).GetCell(0).SetCellValue(string.Format("維護日期 : {0} ~ {1}", strRptBeforeDate, strRptEndDate));
                sheet.GetRow(4).GetCell(6).SetCellValue(strAgentName);//製表人
                sheet.GetRow(7).GetCell(0).SetCellValue("卡號");
                #endregion

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 資料
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sheet.CreateRow(sheet.LastRowNum + 1);
                    for (int b = 0; b < 7; b++)
                    {
                        sheet.GetRow(sheet.LastRowNum).CreateCell(b);
                        sheet.GetRow(sheet.LastRowNum).GetCell(b).CellStyle = contentFormat;
                    }
                    sheet.GetRow(sheet.LastRowNum).GetCell(0).SetCellValue(dt.Rows[i]["CUST_ID"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue(dt.Rows[i]["FLD_NAME"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(dt.Rows[i]["BEFOR_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(3).SetCellValue(dt.Rows[i]["AFTER_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(dt.Rows[i]["MAINT_D"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(5).SetCellValue(dt.Rows[i]["MAINT_T"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(6).SetCellValue(dt.Rows[i]["USER_ID"].ToString().Trim());
                }
                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report02010301" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }
        }
        #endregion

        #region Report02010302 信用卡卡片-調整臨時額度
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Report02010302(卡片/額度查詢/調整臨時額度)資料並產出Excel
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/10
        /// </summary>
        /// <param name="strRptPeople"></param>
        /// <param name="strRptBeforeAmount"></param>
        /// <param name="strRptEndAmount"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="strAgentName"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report02010302(string strRptPeople, string strRptBeforeAmount, string strRptEndAmount, string strRptBeforeDate, string strRptEndDate, string strAgentName, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "member1.xls";
                int totalCount = 0;
                DataTable dt = getData_member1(agentId, ref totalCount, "02010302", 0, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");
                #region 表頭
                sheet.GetRow(0).GetCell(0).SetCellValue("調整臨時額度");
                sheet.GetRow(2).GetCell(0).SetCellValue(string.Format("維護員 : {0}", strRptPeople));
                sheet.GetRow(3).GetCell(0).SetCellValue(string.Format("額度 : {0} ~ {1}", strRptBeforeAmount, strRptEndAmount));
                sheet.GetRow(3).GetCell(6).SetCellValue(DateTime.Now.ToString("yyyyMMdd"));//製表日
                sheet.GetRow(4).GetCell(0).SetCellValue(string.Format("維護日期 : {0} ~ {1}", strRptBeforeDate, strRptEndDate));
                sheet.GetRow(4).GetCell(6).SetCellValue(strAgentName);//製表人
                sheet.GetRow(7).GetCell(0).SetCellValue("卡號");
                #endregion

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 資料
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sheet.CreateRow(sheet.LastRowNum + 1);
                    for (int b = 0; b < 7; b++)
                    {
                        sheet.GetRow(sheet.LastRowNum).CreateCell(b);
                        sheet.GetRow(sheet.LastRowNum).GetCell(b).CellStyle = contentFormat;
                    }
                    sheet.GetRow(sheet.LastRowNum).GetCell(0).SetCellValue(dt.Rows[i]["CUST_ID"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue(dt.Rows[i]["FLD_NAME"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(dt.Rows[i]["BEFOR_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(3).SetCellValue(dt.Rows[i]["AFTER_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(dt.Rows[i]["MAINT_D"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(5).SetCellValue(dt.Rows[i]["MAINT_T"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(6).SetCellValue(dt.Rows[i]["USER_ID"].ToString().Trim());
                }
                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report02010302" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }
        }
        #endregion

        #region Report02010303 信用卡卡片-新卡額度
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Report02010303(卡片/額度查詢/新卡額度)資料並產出Excel
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/10
        /// </summary>
        /// <param name="strRptPeople"></param>
        /// <param name="strRptBeforeAmount"></param>
        /// <param name="strRptEndAmount"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="strAgentName"></param>
        /// <param name="agentId"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report02010303(string strRptPeople, string strRptBeforeAmount, string strRptEndAmount, string strRptBeforeDate, string strRptEndDate, string strAgentName, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "member1.xls";
                int totalCount = 0;
                DataTable dt = getData_member1(agentId, ref totalCount, "02010303", 0, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");
                #region 表頭
                sheet.GetRow(0).GetCell(0).SetCellValue("新卡額度");
                sheet.GetRow(2).GetCell(0).SetCellValue(string.Format("維護員 : {0}", strRptPeople));
                sheet.GetRow(3).GetCell(0).SetCellValue(string.Format("額度 : {0} ~ {1}", strRptBeforeAmount, strRptEndAmount));
                sheet.GetRow(3).GetCell(6).SetCellValue(DateTime.Now.ToString("yyyyMMdd"));//製表日
                sheet.GetRow(4).GetCell(0).SetCellValue(string.Format("維護日期 : {0} ~ {1}", strRptBeforeDate, strRptEndDate));
                sheet.GetRow(4).GetCell(6).SetCellValue(strAgentName);//製表人
                sheet.GetRow(7).GetCell(0).SetCellValue("卡號");
                #endregion

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 資料
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sheet.CreateRow(sheet.LastRowNum + 1);
                    for (int b = 0; b < 7; b++)
                    {
                        sheet.GetRow(sheet.LastRowNum).CreateCell(b);
                        sheet.GetRow(sheet.LastRowNum).GetCell(b).CellStyle = contentFormat;
                    }
                    sheet.GetRow(sheet.LastRowNum).GetCell(0).SetCellValue(dt.Rows[i]["CUST_ID"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue(dt.Rows[i]["FLD_NAME"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(dt.Rows[i]["BEFOR_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(3).SetCellValue(dt.Rows[i]["AFTER_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(dt.Rows[i]["MAINT_D"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(5).SetCellValue(dt.Rows[i]["MAINT_T"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(6).SetCellValue(dt.Rows[i]["USER_ID"].ToString().Trim());
                }
                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report02010303" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }
        }
        #endregion

        #region Report02010401 信用卡卡片-非流通BlockCode調整為流通中
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Report02010401(卡片/特殊查詢/非流通BlockCode調整為流通中)資料並產出Excel
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/10
        /// </summary>
        /// <param name="strRptPeople"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="strAgentName"></param>
        /// <param name="agentId"></param>
        /// <param name="strRptBeforeContent"></param>
        /// <param name="strRptEndContent"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report02010401(string strRptPeople, string strRptBeforeDate, string strRptEndDate, string strAgentName, string agentId, string strRptBeforeContent, string strRptEndContent, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "member1.xls";
                int totalCount = 0;
                DataTable dt = getData_member1(agentId, ref totalCount, "02010401", 0, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");
                #region 表頭
                sheet.GetRow(0).GetCell(0).SetCellValue("非流通BlockCode調整為流通中");
                sheet.GetRow(2).GetCell(0).SetCellValue(string.Format("調整前內容 : {0}", strRptBeforeContent));
                sheet.GetRow(3).GetCell(0).SetCellValue(string.Format("調整後內容 : {0}", strRptEndContent));
                sheet.GetRow(3).GetCell(6).SetCellValue(DateTime.Now.ToString("yyyyMMdd"));//製表日
                sheet.GetRow(4).GetCell(0).SetCellValue(string.Format("維護員 : {0}", strRptPeople));
                sheet.GetRow(4).GetCell(6).SetCellValue(strAgentName);//製表人
                sheet.GetRow(5).GetCell(0).SetCellValue(string.Format("維護日期 : {0} ~ {1}", strRptBeforeDate, strRptEndDate));
                sheet.GetRow(7).GetCell(0).SetCellValue("卡號");
                #endregion

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 資料
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sheet.CreateRow(sheet.LastRowNum + 1);
                    for (int b = 0; b < 7; b++)
                    {
                        sheet.GetRow(sheet.LastRowNum).CreateCell(b);
                        sheet.GetRow(sheet.LastRowNum).GetCell(b).CellStyle = contentFormat;
                    }
                    sheet.GetRow(sheet.LastRowNum).GetCell(0).SetCellValue(dt.Rows[i]["CUST_ID"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue(dt.Rows[i]["FLD_NAME"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(dt.Rows[i]["BEFOR_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(3).SetCellValue(dt.Rows[i]["AFTER_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(dt.Rows[i]["MAINT_D"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(5).SetCellValue(dt.Rows[i]["MAINT_T"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(6).SetCellValue(dt.Rows[i]["USER_ID"].ToString().Trim());
                }
                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report02010401" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }
        }
        #endregion

        #region Report02010402 信用卡卡片-年費無優惠調整為優惠條件
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Report02010402(卡片/特殊查詢/年費無優惠調整為優惠條件)資料並產出Excel
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/10
        /// </summary>
        /// <param name="strRptPeople"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="strAgentName"></param>
        /// <param name="agentId"></param>
        /// <param name="strRptBeforeContent"></param>
        /// <param name="strRptEndContent"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report02010402(string strRptPeople, string strRptBeforeDate, string strRptEndDate, string strAgentName, string agentId, string strRptBeforeContent, string strRptEndContent, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "member1.xls";
                int totalCount = 0;
                DataTable dt = getData_member1(agentId, ref totalCount, "02010402", 0, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");
                #region 表頭
                sheet.GetRow(0).GetCell(0).SetCellValue("年費無優惠調整爲優惠條件");
                sheet.GetRow(2).GetCell(0).SetCellValue(string.Format("調整前內容 : {0}", strRptBeforeContent));
                sheet.GetRow(3).GetCell(0).SetCellValue(string.Format("調整后內容 : {0}", strRptEndContent));
                sheet.GetRow(3).GetCell(6).SetCellValue(DateTime.Now.ToString("yyyyMMdd"));//製表日
                sheet.GetRow(4).GetCell(0).SetCellValue(string.Format("維護員 : {0}", strRptPeople));
                sheet.GetRow(4).GetCell(6).SetCellValue(strAgentName);//製表人
                sheet.GetRow(5).GetCell(0).SetCellValue(string.Format("維護日期 : {0} ~ {1}", strRptBeforeDate, strRptEndDate));
                sheet.GetRow(7).GetCell(0).SetCellValue("卡號");
                #endregion

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 資料
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sheet.CreateRow(sheet.LastRowNum + 1);
                    for (int b = 0; b < 7; b++)
                    {
                        sheet.GetRow(sheet.LastRowNum).CreateCell(b);
                        sheet.GetRow(sheet.LastRowNum).GetCell(b).CellStyle = contentFormat;
                    }
                    sheet.GetRow(sheet.LastRowNum).GetCell(0).SetCellValue(dt.Rows[i]["CUST_ID"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue(dt.Rows[i]["FLD_NAME"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(dt.Rows[i]["BEFOR_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(3).SetCellValue(dt.Rows[i]["AFTER_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(dt.Rows[i]["MAINT_D"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(5).SetCellValue(dt.Rows[i]["MAINT_T"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(6).SetCellValue(dt.Rows[i]["USER_ID"].ToString().Trim());
                }
                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report02010402" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }
        }
        #endregion

        #endregion

        #region VD卡人

        #region Report03010100 VD卡人-維護記錄查詢
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Report03010100(VD卡人/維護記錄查詢)資料並產出Excel
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/10
        /// </summary>
        /// <param name="strRptID"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="strAgentName"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report03010100(string strRptID, string strRptBeforeDate, string strRptEndDate, string strAgentName, string agentId, ref string strPathFile, ref string strMsgID, ref bool isCSV)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "member1.xls";
                int totalCount = 0;
                DataTable dt = getData_member1(agentId, ref totalCount, "03010100", 0, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                if (dt.Rows.Count < 60000)
                {
                    //資料少於6萬筆，以Excel產出
                    FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                    HSSFWorkbook wb = new HSSFWorkbook(fs);
                    ISheet sheet = wb.GetSheet("工作表1");
                    #region 表頭
                    sheet.GetRow(0).GetCell(0).SetCellValue("VD維護記錄查詢");
                    sheet.GetRow(2).GetCell(0).SetCellValue(string.Format("卡人:{0}", strRptID));
                    sheet.GetRow(3).GetCell(0).SetCellValue(string.Format("維護日期 : {0} ~ {1}", strRptBeforeDate, strRptEndDate));
                    sheet.GetRow(3).GetCell(6).SetCellValue(DateTime.Now.ToString("yyyyMMdd"));//製表日
                    sheet.GetRow(4).GetCell(6).SetCellValue(strAgentName);//製表人
                    sheet.GetRow(7).GetCell(0).SetCellValue("卡人ID");
                    #endregion

                    //取得樣式
                    HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                    #region 資料
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        sheet.CreateRow(sheet.LastRowNum + 1);
                        for (int c = 0; c < 7; c++)
                        {
                            sheet.GetRow(sheet.LastRowNum).CreateCell(c);
                            sheet.GetRow(sheet.LastRowNum).GetCell(c).CellStyle = contentFormat;
                        }
                        sheet.GetRow(sheet.LastRowNum).GetCell(0).SetCellValue(dt.Rows[i]["CUST_ID"].ToString().Trim());
                        sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue(dt.Rows[i]["FLD_NAME"].ToString().Trim());
                        sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(dt.Rows[i]["BEFOR_UPD"].ToString().Trim());
                        sheet.GetRow(sheet.LastRowNum).GetCell(3).SetCellValue(dt.Rows[i]["AFTER_UPD"].ToString().Trim());
                        sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(dt.Rows[i]["MAINT_D"].ToString().Trim());
                        sheet.GetRow(sheet.LastRowNum).GetCell(5).SetCellValue(dt.Rows[i]["MAINT_T"].ToString().Trim());
                        sheet.GetRow(sheet.LastRowNum).GetCell(6).SetCellValue(dt.Rows[i]["USER_ID"].ToString().Trim());
                    }
                    #endregion

                    // 保存文件到運行目錄下
                    strPathFile = strPathFile + @"\ExcelFile_Report03010100" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                    FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                    wb.Write(fs1);
                    fs1.Close();
                    fs.Close();
                }
                else
                {
                    //資料多於6萬筆，以CSV產出
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine("卡人ID,欄位名稱,調整前內容,調整後內容,維護日期,維護時間,維護員");
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        sb.AppendLine(
                            string.Format("=\"{0}\"", dt.Rows[i]["CUST_ID"].ToString().Trim()) + "," +
                            string.Format("=\"{0}\"", dt.Rows[i]["FLD_NAME"].ToString().Trim()) + "," +
                            string.Format("=\"{0}\"", dt.Rows[i]["BEFOR_UPD"].ToString().Trim().Replace(",", "")) + "," +
                            string.Format("=\"{0}\"", dt.Rows[i]["AFTER_UPD"].ToString().Trim().Replace(",", "")) + "," +
                            string.Format("=\"{0}\"", dt.Rows[i]["MAINT_D"].ToString().Trim()) + "," +
                            string.Format("=\"{0}\"", dt.Rows[i]["MAINT_T"].ToString().Trim()) + "," +
                            string.Format("=\"{0}\"", dt.Rows[i]["USER_ID"].ToString().Trim())
                            );
                    }
                    strPathFile = strPathFile + @"\ExcelFile_Report03010100" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";
                    File.WriteAllText(strPathFile, sb.ToString(), Encoding.Default);
                    isCSV = true;
                }


                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }


        }
        #endregion

        #region Report03010201 VD卡人-維護員統計表
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Excel_Report03010201(VD卡人/統計表或關係表查詢/維護員統計表)
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/30
        /// </summary>
        /// <param name="strName"></param>
        /// <param name="strRptPeople"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="agentId"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report03010201(string strName, string strRptPeople, string strRptBeforeDate, string strRptEndDate, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "user.xls";
                int totalCount = 0;
                DataTable dt = getData_Comm(agentId, string.Format(sqlComm_01010201, UtilHelper.GetAppSettings("DB_CP_DBF")), ref totalCount, "03010201", 0, true, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 表頭
                sheet.GetRow(4).GetCell(0).SetCellValue("維護員：" + strRptPeople);//維護員
                sheet.GetRow(5).GetCell(0).SetCellValue("維護日期：" + strRptBeforeDate + " ~ " + strRptEndDate);//維護日期
                sheet.GetRow(4).GetCell(6).SetCellValue("製表日：" + DateTime.Now.ToString("yyyy/MM/dd"));//製表日
                sheet.GetRow(5).GetCell(6).SetCellValue("製表人：" + strName);//製表人
                sheet.GetRow(1).GetCell(2).SetCellValue("VD維護員統計表");

                #endregion


                #region 表身

                //資料去空白
                removeBlank(ref dt);

                ExportExcelForNPOI_SubTotal(dt, ref wb, 7, "工作表1");

                #region 合併相同維護員資料
                int startRow = 0;
                int endRow = 0;
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    if (row == sheet.LastRowNum)
                    {
                        if (startRow > 0)
                        {
                            endRow = row;
                        }
                        if (endRow - startRow >= 1)
                        {
                            sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                            sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                            //建立小計
                            NPOI_AddSubTotal(sheet, startRow, endRow, contentFormat);
                            startRow = 0;
                            endRow = 0;
                        }
                        break;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() == sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString())
                    {
                        if (startRow != 0)
                            continue;
                        startRow = row;
                        continue;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() != sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString() || (row == sheet.LastRowNum - 1 && startRow > 0))
                    {
                        if (startRow != 0)
                        {
                            endRow = row;

                            if (endRow - startRow >= 1)
                            {
                                sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                                sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                                // 建立小計
                                NPOI_AddSubTotal(sheet, startRow, endRow, contentFormat);
                                startRow = 0;
                                endRow = 0;
                                continue;
                            }
                        }

                        if (startRow == 0 && sheet.GetRow(row).GetCell(3).StringCellValue.ToString() != "小計")
                        {
                            // 建立小計
                            NPOI_AddSubTotal(sheet, row, row, contentFormat);
                        }
                    }
                }

                //尾列總計
                int sumValue = 0;
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    if (sheet.GetRow(row).GetCell(3).StringCellValue == "小計")
                    {
                        int result = 0;
                        bool tryParse = int.TryParse(sheet.GetRow(row).GetCell(4).NumericCellValue.ToString(), out result);
                        if (tryParse)
                        {
                            sumValue += result;
                        }
                    }
                }
                sheet.CreateRow(sheet.LastRowNum + 1);
                for (int col = 2; col < 5; col++)
                {
                    sheet.GetRow(sheet.LastRowNum).CreateCell(col);
                    sheet.GetRow(sheet.LastRowNum).GetCell(col).CellStyle = contentFormat;
                }
                sheet.AddMergedRegion(new CellRangeAddress(sheet.LastRowNum, sheet.LastRowNum, 2, 3));
                sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue("總計");
                sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(sumValue);
                #endregion

                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report03010201" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }

        }
        #endregion

        #region Report03010202 VD卡人-卡人與維護員關係表
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Excel_Report03010202(VD卡人/統計表或關係表查詢/卡人與維護員關係表)
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/30
        /// </summary>
        /// <param name="strName"></param>
        /// <param name="strRptPeople"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="agentId"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report03010202(string strName, string strRptPeople, string strRptBeforeDate, string strRptEndDate, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "card_r.xls";
                int totalCount = 0;
                DataTable dt = getData_Comm(agentId, string.Format(sqlComm_01010202, UtilHelper.GetAppSettings("DB_CP_DBF")), ref totalCount, "03010202", 0, true, "P");
                if (dt.Rows.Count <= 0)
                    return false;
                DataTable dt2 = getData_Comm(agentId, sqlComm_01010202_Count, ref totalCount, "03010202差異總計", 0, false);
                string totalDiffCount = "";
                if (dt2.Rows.Count > 0)
                {
                    totalDiffCount = dt2.Rows[0][0].ToString();
                }
                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 表頭
                sheet.GetRow(4).GetCell(0).SetCellValue("維護員：" + strRptPeople);//維護員
                sheet.GetRow(5).GetCell(0).SetCellValue("維護日期：" + strRptBeforeDate + " ~ " + strRptEndDate);//維護日期
                sheet.GetRow(4).GetCell(5).SetCellValue("製表日：" + DateTime.Now.ToString("yyyy/MM/dd"));//製表日
                sheet.GetRow(5).GetCell(5).SetCellValue("製表人：" + strName);//製表人
                sheet.GetRow(1).GetCell(1).SetCellValue("VD卡人與維護員關係表");//標題
                #endregion


                #region 表身

                //資料去空白
                removeBlank(ref dt);

                //資料寫入
                ExportExcelForNPOI(dt, ref wb, 7, "工作表1", 1);

                #region 合併相同維護員資料
                int startRow = 0;
                int endRow = 0;
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    if (row == sheet.LastRowNum)
                    {
                        if (startRow > 0)
                        {
                            endRow = row;
                        }
                        if (endRow - startRow >= 1)
                        {
                            sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 1, 1));
                            sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                            sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                            startRow = 0;
                            endRow = 0;
                        }
                        break;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() == sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString())
                    {
                        if (startRow != 0)
                            continue;
                        startRow = row;
                        continue;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() != sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString() || (row == sheet.LastRowNum - 1 && startRow > 0))
                    {
                        if (startRow != 0)
                        {
                            endRow = row;

                            if (endRow - startRow >= 1)
                            {
                                sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 1, 1));
                                sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                                sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                                startRow = 0;
                                endRow = 0;
                                continue;
                            }
                        }
                    }
                }
                #endregion

                //增加尾列總計
                int sumValue = NPOI_ColumnSum(sheet, 7, sheet.LastRowNum);
                sheet.CreateRow(sheet.LastRowNum + 1);
                for (int c = 1; c < 6; c++)
                {
                    sheet.GetRow(sheet.LastRowNum).CreateCell(c);
                    sheet.GetRow(sheet.LastRowNum).GetCell(c).CellStyle = contentFormat;
                }
                sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue("總計");
                sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(totalDiffCount);
                sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(sumValue);

                //小計欄位合併
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    sheet.GetRow(row).CreateCell(5);
                    sheet.GetRow(row).GetCell(5).CellStyle = contentFormat;
                }
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    sheet.AddMergedRegion(new CellRangeAddress(row, row, 4, 5));
                    sheet.GetRow(row).GetCell(4).CellStyle = contentFormat;
                }
                //額外合併尾列維護員、小計
                sheet.AddMergedRegion(new CellRangeAddress(sheet.LastRowNum, sheet.LastRowNum, 3, 4));
                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report03010202" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }
        }
        #endregion

        #region Report03010203 VD卡人-維護欄位統計表
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Excel_Report03010203(VD卡人/統計表或關係表查詢/維護欄位統計表)
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/30
        /// </summary>
        /// <param name="strName"></param>
        /// <param name="strRptFld"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="agentId"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report03010203(string strName, string strRptFld, string strRptBeforeDate, string strRptEndDate, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "fld.xls";
                int totalCount = 0;
                DataTable dt = getData_Comm(agentId, string.Format(sqlComm_01010203, UtilHelper.GetAppSettings("DB_CP_DBF")), ref totalCount, "03010203", 0, true, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 表頭
                sheet.GetRow(1).GetCell(2).SetCellValue("VD維護欄位統計表");
                sheet.GetRow(4).GetCell(0).SetCellValue("維護欄位：" + strRptFld);//維護欄位
                sheet.GetRow(5).GetCell(0).SetCellValue("維護日期：" + strRptBeforeDate + " ~ " + strRptEndDate);//維護日期
                sheet.GetRow(4).GetCell(6).SetCellValue("製表日：" + DateTime.Now.ToString("yyyy/MM/dd"));//製表日
                sheet.GetRow(5).GetCell(6).SetCellValue("製表人：" + strName);//製表人
                #endregion


                #region 表身

                //資料去空白
                removeBlank(ref dt);

                ExportExcelForNPOI_SubTotal(dt, ref wb, 7, "工作表1");

                #region 合併相同維護員資料
                int startRow = 0;
                int endRow = 0;
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    if (row == sheet.LastRowNum)
                    {
                        if (startRow > 0)
                        {
                            endRow = row;
                        }
                        if (endRow - startRow >= 1)
                        {
                            sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                            sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                            //建立小計
                            NPOI_AddSubTotal(sheet, startRow, endRow, contentFormat);
                            startRow = 0;
                            endRow = 0;
                        }
                        break;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() == sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString())
                    {
                        if (startRow != 0)
                            continue;
                        startRow = row;
                        continue;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() != sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString() || (row == sheet.LastRowNum - 1 && startRow > 0))
                    {
                        if (startRow != 0)
                        {
                            endRow = row;

                            if (endRow - startRow >= 1)
                            {
                                sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                                sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                                // 建立小計
                                NPOI_AddSubTotal(sheet, startRow, endRow, contentFormat);
                                startRow = 0;
                                endRow = 0;
                                continue;
                            }
                        }

                        if (startRow == 0 && sheet.GetRow(row).GetCell(3).StringCellValue.ToString() != "小計")
                        {
                            // 建立小計
                            NPOI_AddSubTotal(sheet, row, row, contentFormat);
                        }
                    }
                }

                //尾列總計
                int sumValue = 0;
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    if (sheet.GetRow(row).GetCell(3).StringCellValue == "小計")
                    {
                        int result = 0;
                        bool tryParse = int.TryParse(sheet.GetRow(row).GetCell(4).NumericCellValue.ToString(), out result);
                        if (tryParse)
                        {
                            sumValue += result;
                        }
                    }
                }
                sheet.CreateRow(sheet.LastRowNum + 1);
                for (int col = 2; col < 5; col++)
                {
                    sheet.GetRow(sheet.LastRowNum).CreateCell(col);
                    sheet.GetRow(sheet.LastRowNum).GetCell(col).CellStyle = contentFormat;
                }
                sheet.AddMergedRegion(new CellRangeAddress(sheet.LastRowNum, sheet.LastRowNum, 2, 3));
                sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue("總計");
                sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(sumValue);
                #endregion

                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report03010203" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }

        }
        #endregion

        #region Report03010204 VD卡人-卡人調整統計表
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Excel_Report03010204(VD卡人/統計表或關係表查詢/卡人調整統計表)
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/30
        /// </summary>
        /// <param name="strName"></param>
        /// <param name="strRptFld"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="agentId"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report03010204(string strName, string strRptFld, string strRptBeforeDate, string strRptEndDate, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "card.xls";
                int totalCount = 0;
                DataTable dt = getData_Comm(agentId, string.Format(sqlComm_01010204, UtilHelper.GetAppSettings("DB_CP_DBF")), ref totalCount, "03010204", 0, true, "P");
                if (dt.Rows.Count <= 0)
                    return false;
                DataTable dt2 = getData_Comm(agentId, sqlComm_01010204_Count, ref totalCount, "03010204差異總計", 0, false);
                string totalDiffCount = "";
                if (dt2.Rows.Count > 0)
                {
                    totalDiffCount = dt2.Rows[0][0].ToString();
                }
                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 表頭
                sheet.GetRow(1).GetCell(1).SetCellValue("VD卡人調整統計表");//標題
                sheet.GetRow(4).GetCell(0).SetCellValue("維護欄位：" + strRptFld);//維護欄位
                sheet.GetRow(5).GetCell(0).SetCellValue("維護日期：" + strRptBeforeDate + " ~ " + strRptEndDate);//維護日期
                sheet.GetRow(4).GetCell(5).SetCellValue("製表日：" + DateTime.Now.ToString("yyyy/MM/dd"));//製表日
                sheet.GetRow(5).GetCell(5).SetCellValue("製表人：" + strName);//製表人
                #endregion


                #region 表身

                //資料去空白
                removeBlank(ref dt);

                //資料寫入
                ExportExcelForNPOI(dt, ref wb, 7, "工作表1", 1);

                #region 合併相同資料
                int startRow = 0;
                int endRow = 0;
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    if (row == sheet.LastRowNum)
                    {
                        if (startRow > 0)
                        {
                            endRow = row;
                        }
                        if (endRow - startRow >= 1)
                        {
                            sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 1, 1));
                            sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                            sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                            startRow = 0;
                            endRow = 0;
                        }
                        break;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() == sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString())
                    {
                        if (startRow != 0)
                            continue;
                        startRow = row;
                        continue;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() != sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString() || (row == sheet.LastRowNum - 1 && startRow > 0))
                    {
                        if (startRow != 0)
                        {
                            endRow = row;

                            if (endRow - startRow >= 1)
                            {
                                sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 1, 1));
                                sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                                sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                                startRow = 0;
                                endRow = 0;
                                continue;
                            }
                        }
                    }
                }
                #endregion

                //增加尾列總計
                int sumValue = NPOI_ColumnSum(sheet, 7, sheet.LastRowNum);
                sheet.CreateRow(sheet.LastRowNum + 1);
                for (int c = 1; c < 6; c++)
                {
                    sheet.GetRow(sheet.LastRowNum).CreateCell(c);
                    sheet.GetRow(sheet.LastRowNum).GetCell(c).CellStyle = contentFormat;
                }
                sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue("總計");
                sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(totalDiffCount);
                sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(sumValue);

                //小計欄位合併
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    sheet.GetRow(row).CreateCell(5);
                    sheet.GetRow(row).GetCell(5).CellStyle = contentFormat;
                }
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    sheet.AddMergedRegion(new CellRangeAddress(row, row, 4, 5));
                    sheet.GetRow(row).GetCell(4).CellStyle = contentFormat;
                }
                //額外合併尾列維護員、小計
                sheet.AddMergedRegion(new CellRangeAddress(sheet.LastRowNum, sheet.LastRowNum, 3, 4));
                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report03010204" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }
        }



        #endregion

        #region Report03010301 VD卡人-調整固定額度
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Report03010301(VD卡人/額度查詢/調整固定額度)資料並產出Excel
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/10
        /// </summary>
        /// <param name="strRptPeople"></param>
        /// <param name="strRptBeforeAmount"></param>
        /// <param name="strRptEndAmount"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="strAgentName"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report03010301(string strRptPeople, string strRptBeforeAmount, string strRptEndAmount, string strRptBeforeDate, string strRptEndDate, string strAgentName, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "member1.xls";
                int totalCount = 0;
                DataTable dt = getData_member1(agentId, ref totalCount, "03010301", 0, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");
                #region 表頭
                sheet.GetRow(0).GetCell(0).SetCellValue("VD調整固定額度");
                sheet.GetRow(2).GetCell(0).SetCellValue(string.Format("維護員 : {0}", strRptPeople));
                sheet.GetRow(3).GetCell(0).SetCellValue(string.Format("額度 : {0} ~ {1}", strRptBeforeAmount, strRptEndAmount));
                sheet.GetRow(3).GetCell(6).SetCellValue(DateTime.Now.ToString("yyyyMMdd"));//製表日
                sheet.GetRow(4).GetCell(0).SetCellValue(string.Format("維護日期 : {0} ~ {1}", strRptBeforeDate, strRptEndDate));
                sheet.GetRow(4).GetCell(6).SetCellValue(strAgentName);//製表人
                sheet.GetRow(7).GetCell(0).SetCellValue("卡人ID");
                #endregion

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 資料
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sheet.CreateRow(sheet.LastRowNum + 1);
                    for (int b = 0; b < 7; b++)
                    {
                        sheet.GetRow(sheet.LastRowNum).CreateCell(b);
                        sheet.GetRow(sheet.LastRowNum).GetCell(b).CellStyle = contentFormat;
                    }
                    sheet.GetRow(sheet.LastRowNum).GetCell(0).SetCellValue(dt.Rows[i]["CUST_ID"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue(dt.Rows[i]["FLD_NAME"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(dt.Rows[i]["BEFOR_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(3).SetCellValue(dt.Rows[i]["AFTER_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(dt.Rows[i]["MAINT_D"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(5).SetCellValue(dt.Rows[i]["MAINT_T"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(6).SetCellValue(dt.Rows[i]["USER_ID"].ToString().Trim());
                }
                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report03010301" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }
        }
        #endregion

        #region Report03010302 VD卡人-調整臨時額度
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Report03010302(VD卡人/額度查詢/調整臨時額度)資料並產出Excel
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/10
        /// </summary>
        /// <param name="strRptPeople"></param>
        /// <param name="strRptBeforeAmount"></param>
        /// <param name="strRptEndAmount"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="strAgentName"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report03010302(string strRptPeople, string strRptBeforeAmount, string strRptEndAmount, string strRptBeforeDate, string strRptEndDate, string strAgentName, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "member1.xls";
                int totalCount = 0;
                DataTable dt = getData_member1(agentId, ref totalCount, "03010302", 0, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");
                #region 表頭
                sheet.GetRow(0).GetCell(0).SetCellValue("VD調整臨時額度");
                sheet.GetRow(2).GetCell(0).SetCellValue(string.Format("維護員 : {0}", strRptPeople));
                sheet.GetRow(3).GetCell(0).SetCellValue(string.Format("額度 : {0} ~ {1}", strRptBeforeAmount, strRptEndAmount));
                sheet.GetRow(3).GetCell(6).SetCellValue(DateTime.Now.ToString("yyyyMMdd"));//製表日
                sheet.GetRow(4).GetCell(0).SetCellValue(string.Format("維護日期 : {0} ~ {1}", strRptBeforeDate, strRptEndDate));
                sheet.GetRow(4).GetCell(6).SetCellValue(strAgentName);//製表人
                sheet.GetRow(7).GetCell(0).SetCellValue("卡人ID");
                #endregion

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 資料
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sheet.CreateRow(sheet.LastRowNum + 1);
                    for (int b = 0; b < 7; b++)
                    {
                        sheet.GetRow(sheet.LastRowNum).CreateCell(b);
                        sheet.GetRow(sheet.LastRowNum).GetCell(b).CellStyle = contentFormat;
                    }
                    sheet.GetRow(sheet.LastRowNum).GetCell(0).SetCellValue(dt.Rows[i]["CUST_ID"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue(dt.Rows[i]["FLD_NAME"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(dt.Rows[i]["BEFOR_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(3).SetCellValue(dt.Rows[i]["AFTER_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(dt.Rows[i]["MAINT_D"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(5).SetCellValue(dt.Rows[i]["MAINT_T"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(6).SetCellValue(dt.Rows[i]["USER_ID"].ToString().Trim());
                }
                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report03010302" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }
        }
        #endregion

        #region Report03010303 VD卡人-新卡額度
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Report03010303(VD卡人/額度查詢/新卡額度)資料並產出Excel
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/10
        /// </summary>
        /// <param name="strRptPeople"></param>
        /// <param name="strRptBeforeAmount"></param>
        /// <param name="strRptEndAmount"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="strAgentName"></param>
        /// <param name="agentId"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report03010303(string strRptPeople, string strRptBeforeAmount, string strRptEndAmount, string strRptBeforeDate, string strRptEndDate, string strAgentName, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "member1.xls";
                int totalCount = 0;
                DataTable dt = getData_member1(agentId, ref totalCount, "03010303", 0, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");
                #region 表頭
                sheet.GetRow(0).GetCell(0).SetCellValue("VD新卡額度");
                sheet.GetRow(2).GetCell(0).SetCellValue(string.Format("維護員 : {0}", strRptPeople));
                sheet.GetRow(3).GetCell(0).SetCellValue(string.Format("額度 : {0} ~ {1}", strRptBeforeAmount, strRptEndAmount));
                sheet.GetRow(3).GetCell(6).SetCellValue(DateTime.Now.ToString("yyyyMMdd"));//製表日
                sheet.GetRow(4).GetCell(0).SetCellValue(string.Format("維護日期 : {0} ~ {1}", strRptBeforeDate, strRptEndDate));
                sheet.GetRow(4).GetCell(6).SetCellValue(strAgentName);//製表人
                sheet.GetRow(7).GetCell(0).SetCellValue("卡人ID");
                #endregion

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 資料
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sheet.CreateRow(sheet.LastRowNum + 1);
                    for (int b = 0; b < 7; b++)
                    {
                        sheet.GetRow(sheet.LastRowNum).CreateCell(b);
                        sheet.GetRow(sheet.LastRowNum).GetCell(b).CellStyle = contentFormat;
                    }
                    sheet.GetRow(sheet.LastRowNum).GetCell(0).SetCellValue(dt.Rows[i]["CUST_ID"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue(dt.Rows[i]["FLD_NAME"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(dt.Rows[i]["BEFOR_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(3).SetCellValue(dt.Rows[i]["AFTER_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(dt.Rows[i]["MAINT_D"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(5).SetCellValue(dt.Rows[i]["MAINT_T"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(6).SetCellValue(dt.Rows[i]["USER_ID"].ToString().Trim());
                }
                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report03010303" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }
        }
        #endregion

        #region Report03010401 VD卡人-員工調整記錄
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Report03010401(VD卡人/特殊查詢/員工調整記錄)資料並產出Excel
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/10
        /// </summary>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="strAgentName"></param>
        /// <param name="agentId"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report03010401(string strRptBeforeDate, string strRptEndDate, string strAgentName, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "member.xls";
                int totalCount = 0;
                DataTable dt = getData_member(agentId, ref totalCount, "03010401", 0, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");

                #region 表頭
                sheet.GetRow(0).GetCell(0).SetCellValue("VD員工調整記錄");
                sheet.GetRow(3).GetCell(6).SetCellValue(DateTime.Now.ToString("yyyyMMdd"));//製表日
                sheet.GetRow(3).GetCell(0).SetCellValue(string.Format("維護日期 : {0} ~ {1}", strRptBeforeDate, strRptEndDate));
                sheet.GetRow(4).GetCell(6).SetCellValue(strAgentName);//製表人
                sheet.GetRow(7).GetCell(0).SetCellValue("卡人ID");
                #endregion

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 資料
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sheet.CreateRow(sheet.LastRowNum + 1);
                    for (int b = 0; b < 7; b++)
                    {
                        sheet.GetRow(sheet.LastRowNum).CreateCell(b);
                        sheet.GetRow(sheet.LastRowNum).GetCell(b).CellStyle = contentFormat;
                    }
                    sheet.GetRow(sheet.LastRowNum).GetCell(0).SetCellValue(dt.Rows[i]["ID"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue(dt.Rows[i]["FLD_NAME"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(dt.Rows[i]["BEFOR_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(3).SetCellValue(dt.Rows[i]["AFTER_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(dt.Rows[i]["MAINT_D"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(5).SetCellValue(dt.Rows[i]["MAINT_T"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(6).SetCellValue(dt.Rows[i]["USER_ID"].ToString().Trim());
                }
                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report03010401" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }
        }
        #endregion

        #region Report03010402 VD卡人-自扣帳戶ID與卡人ID不同者
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Report03010402(VD卡人/特殊查詢/自扣帳戶ID與卡人ID不同者)資料並產出Excel
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/10
        /// </summary>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="strAgentName"></param>
        /// <param name="agentId"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report03010402(string strRptBeforeDate, string strRptEndDate, string strAgentName, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "member1.xls";
                int totalCount = 0;
                DataTable dt = getData_member1(agentId, ref totalCount, "03010402", 0, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");

                #region 表頭
                sheet.GetRow(0).GetCell(0).SetCellValue("VD自扣ID與卡人ID不同者");
                sheet.GetRow(3).GetCell(0).SetCellValue("維護欄位 : DIRECT DEBIT ID");
                sheet.GetRow(3).GetCell(6).SetCellValue(DateTime.Now.ToString("yyyyMMdd"));//製表日
                sheet.GetRow(4).GetCell(0).SetCellValue(string.Format("維護日期 : {0} ~ {1}", strRptBeforeDate, strRptEndDate));
                sheet.GetRow(4).GetCell(6).SetCellValue(strAgentName);//製表人
                sheet.GetRow(7).GetCell(0).SetCellValue("卡人ID");
                #endregion

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 資料
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sheet.CreateRow(sheet.LastRowNum + 1);
                    for (int b = 0; b < 7; b++)
                    {
                        sheet.GetRow(sheet.LastRowNum).CreateCell(b);
                        sheet.GetRow(sheet.LastRowNum).GetCell(b).CellStyle = contentFormat;
                    }
                    sheet.GetRow(sheet.LastRowNum).GetCell(0).SetCellValue(dt.Rows[i]["CUST_ID"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue(dt.Rows[i]["FLD_NAME"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(dt.Rows[i]["BEFOR_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(3).SetCellValue(dt.Rows[i]["AFTER_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(dt.Rows[i]["MAINT_D"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(5).SetCellValue(dt.Rows[i]["MAINT_T"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(6).SetCellValue(dt.Rows[i]["USER_ID"].ToString().Trim());
                }
                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report03010402" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }
        }
        #endregion

        #endregion

        #region VD卡片

        #region Report04010100 VD卡片-維護記錄查詢

        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Report04010100(VD卡片/維護記錄查詢)資料並產出Excel
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/10
        /// </summary>
        /// <param name="strRptID"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="strAgentName"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report04010100(string strRptID, string strRptBeforeDate, string strRptEndDate, string strAgentName, string agentId, ref string strPathFile, ref string strMsgID, ref bool isCSV)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "member1.xls";
                int totalCount = 0;
                DataTable dt = getData_member1(agentId, ref totalCount, "04010100", 0, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                if (dt.Rows.Count < 60000)
                {
                    //資料少於6萬筆，以Excel產出

                    FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                    HSSFWorkbook wb = new HSSFWorkbook(fs);
                    ISheet sheet = wb.GetSheet("工作表1");
                    #region 表頭
                    sheet.GetRow(0).GetCell(0).SetCellValue("VD維護記錄查詢");
                    sheet.GetRow(2).GetCell(0).SetCellValue(string.Format("卡片:{0}", strRptID));
                    sheet.GetRow(3).GetCell(0).SetCellValue(string.Format("維護日期 : {0} ~ {1}", strRptBeforeDate, strRptEndDate));
                    sheet.GetRow(3).GetCell(6).SetCellValue(DateTime.Now.ToString("yyyyMMdd"));//製表日
                    sheet.GetRow(4).GetCell(6).SetCellValue(strAgentName);//製表人
                    sheet.GetRow(7).GetCell(0).SetCellValue("卡號");
                    #endregion

                    //取得樣式
                    HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                    #region 資料
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        sheet.CreateRow(sheet.LastRowNum + 1);
                        for (int b = 0; b < 7; b++)
                        {
                            sheet.GetRow(sheet.LastRowNum).CreateCell(b);
                            sheet.GetRow(sheet.LastRowNum).GetCell(b).CellStyle = contentFormat;
                        }
                        sheet.GetRow(sheet.LastRowNum).GetCell(0).SetCellValue(dt.Rows[i]["CUST_ID"].ToString().Trim());
                        sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue(dt.Rows[i]["FLD_NAME"].ToString().Trim());
                        sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(dt.Rows[i]["BEFOR_UPD"].ToString().Trim());
                        sheet.GetRow(sheet.LastRowNum).GetCell(3).SetCellValue(dt.Rows[i]["AFTER_UPD"].ToString().Trim());
                        sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(dt.Rows[i]["MAINT_D"].ToString().Trim());
                        sheet.GetRow(sheet.LastRowNum).GetCell(5).SetCellValue(dt.Rows[i]["MAINT_T"].ToString().Trim());
                        sheet.GetRow(sheet.LastRowNum).GetCell(6).SetCellValue(dt.Rows[i]["USER_ID"].ToString().Trim());
                    }
                    #endregion

                    // 保存文件到運行目錄下
                    strPathFile = strPathFile + @"\ExcelFile_Report04010100" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                    FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                    wb.Write(fs1);
                    fs1.Close();
                    fs.Close();
                }
                else
                {
                    //資料多於6萬筆，以CSV產出
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine("卡號,欄位名稱,調整前內容,調整後內容,維護日期,維護時間,維護員");
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        sb.AppendLine(
                            string.Format("=\"{0}\"", dt.Rows[i]["CUST_ID"].ToString().Trim()) + "," +
                            string.Format("=\"{0}\"", dt.Rows[i]["FLD_NAME"].ToString().Trim()) + "," +
                            string.Format("=\"{0}\"", dt.Rows[i]["BEFOR_UPD"].ToString().Trim().Replace(",", "")) + "," +
                            string.Format("=\"{0}\"", dt.Rows[i]["AFTER_UPD"].ToString().Trim().Replace(",", "")) + "," +
                            string.Format("=\"{0}\"", dt.Rows[i]["MAINT_D"].ToString().Trim()) + "," +
                            string.Format("=\"{0}\"", dt.Rows[i]["MAINT_T"].ToString().Trim()) + "," +
                            string.Format("=\"{0}\"", dt.Rows[i]["USER_ID"].ToString().Trim())
                            );
                    }
                    strPathFile = strPathFile + @"\ExcelFile_Report04010100" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";
                    File.WriteAllText(strPathFile, sb.ToString(), Encoding.Default);
                    isCSV = true;
                }

                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }


        }
        #endregion

        #region Report04010201 VD卡片-維護員統計表
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Excel_Report04010201(VD卡片/統計表或關係表查詢/維護員統計表)
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/30
        /// </summary>
        /// <param name="strName"></param>
        /// <param name="strRptPeople"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="agentId"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report04010201(string strName, string strRptPeople, string strRptBeforeDate, string strRptEndDate, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "user.xls";
                int totalCount = 0;
                DataTable dt = getData_Comm(agentId, string.Format(sqlComm_01010201, UtilHelper.GetAppSettings("DB_CP_DBF")), ref totalCount, "04010201", 0, true, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 表頭
                sheet.GetRow(1).GetCell(2).SetCellValue("VD維護員統計表");
                sheet.GetRow(4).GetCell(0).SetCellValue("維護員：" + strRptPeople);//維護員
                sheet.GetRow(5).GetCell(0).SetCellValue("維護日期：" + strRptBeforeDate + " ~ " + strRptEndDate);//維護日期
                sheet.GetRow(4).GetCell(6).SetCellValue("製表日：" + DateTime.Now.ToString("yyyy/MM/dd"));//製表日
                sheet.GetRow(5).GetCell(6).SetCellValue("製表人：" + strName);//製表人
                #endregion


                #region 表身

                //資料去空白
                removeBlank(ref dt);

                ExportExcelForNPOI_SubTotal(dt, ref wb, 7, "工作表1");

                #region 合併相同維護員資料
                int startRow = 0;
                int endRow = 0;
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    if (row == sheet.LastRowNum)
                    {
                        if (startRow > 0)
                        {
                            endRow = row;
                        }
                        if (endRow - startRow >= 1)
                        {
                            sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                            sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                            //建立小計
                            NPOI_AddSubTotal(sheet, startRow, endRow, contentFormat);
                            startRow = 0;
                            endRow = 0;
                        }
                        break;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() == sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString())
                    {
                        if (startRow != 0)
                            continue;
                        startRow = row;
                        continue;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() != sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString() || (row == sheet.LastRowNum - 1 && startRow > 0))
                    {
                        if (startRow != 0)
                        {
                            endRow = row;

                            if (endRow - startRow >= 1)
                            {
                                sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                                sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                                // 建立小計
                                NPOI_AddSubTotal(sheet, startRow, endRow, contentFormat);
                                startRow = 0;
                                endRow = 0;
                                continue;
                            }
                        }

                        if (startRow == 0 && sheet.GetRow(row).GetCell(3).StringCellValue.ToString() != "小計")
                        {
                            // 建立小計
                            NPOI_AddSubTotal(sheet, row, row, contentFormat);
                        }
                    }
                }

                //尾列總計
                int sumValue = 0;
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    if (sheet.GetRow(row).GetCell(3).StringCellValue == "小計")
                    {
                        int result = 0;
                        bool tryParse = int.TryParse(sheet.GetRow(row).GetCell(4).NumericCellValue.ToString(), out result);
                        if (tryParse)
                        {
                            sumValue += result;
                        }
                    }
                }
                sheet.CreateRow(sheet.LastRowNum + 1);
                for (int col = 2; col < 5; col++)
                {
                    sheet.GetRow(sheet.LastRowNum).CreateCell(col);
                    sheet.GetRow(sheet.LastRowNum).GetCell(col).CellStyle = contentFormat;
                }
                sheet.AddMergedRegion(new CellRangeAddress(sheet.LastRowNum, sheet.LastRowNum, 2, 3));
                sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue("總計");
                sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(sumValue);
                #endregion

                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report04010201" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }

        }
        #endregion

        #region Report04010202 VD卡片-卡人與維護員關係表
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Excel_Report04010202(VD卡片/統計表或關係表查詢/卡人與維護員關係表)
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/30
        /// </summary>
        /// <param name="strName"></param>
        /// <param name="strRptPeople"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="agentId"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report04010202(string strName, string strRptPeople, string strRptBeforeDate, string strRptEndDate, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "card_r.xls";
                int totalCount = 0;
                DataTable dt = getData_Comm(agentId, string.Format(sqlComm_01010202, UtilHelper.GetAppSettings("DB_CP_DBF")), ref totalCount, "04010202", 0, true, "P");
                if (dt.Rows.Count <= 0)
                    return false;
                DataTable dt2 = getData_Comm(agentId, sqlComm_01010202_Count, ref totalCount, "04010202差異總計", 0, false);
                string totalDiffCount = "";
                if (dt2.Rows.Count > 0)
                {
                    totalDiffCount = dt2.Rows[0][0].ToString();
                }
                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 表頭
                sheet.GetRow(1).GetCell(1).SetCellValue("VD卡片與維護員關係表");
                sheet.GetRow(4).GetCell(0).SetCellValue("維護員：" + strRptPeople);//維護員
                sheet.GetRow(5).GetCell(0).SetCellValue("維護日期：" + strRptBeforeDate + " ~ " + strRptEndDate);//維護日期
                sheet.GetRow(4).GetCell(5).SetCellValue("製表日：" + DateTime.Now.ToString("yyyy/MM/dd"));//製表日
                sheet.GetRow(5).GetCell(5).SetCellValue("製表人：" + strName);//製表人
                sheet.GetRow(6).GetCell(2).SetCellValue("卡號");
                #endregion


                #region 表身

                //資料去空白
                removeBlank(ref dt);

                //資料寫入
                ExportExcelForNPOI(dt, ref wb, 7, "工作表1", 1);

                #region 合併相同維護員資料
                int startRow = 0;
                int endRow = 0;
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    if (row == sheet.LastRowNum)
                    {
                        if (startRow > 0)
                        {
                            endRow = row;
                        }
                        if (endRow - startRow >= 1)
                        {
                            sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 1, 1));
                            sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                            sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                            startRow = 0;
                            endRow = 0;
                        }
                        break;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() == sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString())
                    {
                        if (startRow != 0)
                            continue;
                        startRow = row;
                        continue;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() != sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString() || (row == sheet.LastRowNum - 1 && startRow > 0))
                    {
                        if (startRow != 0)
                        {
                            endRow = row;

                            if (endRow - startRow >= 1)
                            {
                                sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 1, 1));
                                sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                                sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                                startRow = 0;
                                endRow = 0;
                                continue;
                            }
                        }
                    }
                }
                #endregion

                //增加尾列總計
                int sumValue = NPOI_ColumnSum(sheet, 7, sheet.LastRowNum);
                sheet.CreateRow(sheet.LastRowNum + 1);
                for (int c = 1; c < 6; c++)
                {
                    sheet.GetRow(sheet.LastRowNum).CreateCell(c);
                    sheet.GetRow(sheet.LastRowNum).GetCell(c).CellStyle = contentFormat;
                }
                sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue("總計");
                sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(totalDiffCount);
                sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(sumValue);

                //小計欄位合併
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    sheet.GetRow(row).CreateCell(5);
                    sheet.GetRow(row).GetCell(5).CellStyle = contentFormat;
                }
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    sheet.AddMergedRegion(new CellRangeAddress(row, row, 4, 5));
                    sheet.GetRow(row).GetCell(4).CellStyle = contentFormat;
                }
                //額外合併尾列維護員、小計
                sheet.AddMergedRegion(new CellRangeAddress(sheet.LastRowNum, sheet.LastRowNum, 3, 4));
                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report04010202" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }
        }
        #endregion

        #region Report04010203 VD卡片-維護欄位統計表
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Excel_Report04010203(VD卡片/統計表或關係表查詢/維護欄位統計表)
        /// 作    者:Ares Stanley
        /// 創建時間:2021/12/13
        /// </summary>
        /// <param name="strName"></param>
        /// <param name="strRptFld"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="agentId"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report04010203(string strName, string strRptFld, string strRptBeforeDate, string strRptEndDate, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "fld.xls";
                int totalCount = 0;
                DataTable dt = getData_Comm(agentId, string.Format(sqlComm_01010203, UtilHelper.GetAppSettings("DB_CP_DBF")), ref totalCount, "04010203", 0, true, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 表頭
                sheet.GetRow(1).GetCell(2).SetCellValue("VD維護欄位統計表");
                sheet.GetRow(4).GetCell(0).SetCellValue("維護欄位：" + strRptFld);//維護欄位
                sheet.GetRow(5).GetCell(0).SetCellValue("維護日期：" + strRptBeforeDate + " ~ " + strRptEndDate);//維護日期
                sheet.GetRow(4).GetCell(6).SetCellValue("製表日：" + DateTime.Now.ToString("yyyy/MM/dd"));//製表日
                sheet.GetRow(5).GetCell(6).SetCellValue("製表人：" + strName);//製表人
                #endregion

                #region 表身

                //資料去空白
                removeBlank(ref dt);

                ExportExcelForNPOI_SubTotal(dt, ref wb, 7, "工作表1");

                #region 合併相同維護員資料
                int startRow = 0;
                int endRow = 0;
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    if (row == sheet.LastRowNum)
                    {
                        if (startRow > 0)
                        {
                            endRow = row;
                        }
                        if (endRow - startRow >= 1)
                        {
                            sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                            sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                            //建立小計
                            NPOI_AddSubTotal(sheet, startRow, endRow, contentFormat);
                            startRow = 0;
                            endRow = 0;
                        }
                        break;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() == sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString())
                    {
                        if (startRow != 0)
                            continue;
                        startRow = row;
                        continue;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() != sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString() || (row == sheet.LastRowNum - 1 && startRow > 0))
                    {
                        if (startRow != 0)
                        {
                            endRow = row;

                            if (endRow - startRow >= 1)
                            {
                                sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                                sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                                // 建立小計
                                NPOI_AddSubTotal(sheet, startRow, endRow, contentFormat);
                                startRow = 0;
                                endRow = 0;
                                continue;
                            }
                        }

                        if (startRow == 0 && sheet.GetRow(row).GetCell(3).StringCellValue.ToString() != "小計")
                        {
                            // 建立小計
                            NPOI_AddSubTotal(sheet, row, row, contentFormat);
                        }
                    }
                }

                //尾列總計
                int sumValue = 0;
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    if (sheet.GetRow(row).GetCell(3).StringCellValue == "小計")
                    {
                        int result = 0;
                        bool tryParse = int.TryParse(sheet.GetRow(row).GetCell(4).NumericCellValue.ToString(), out result);
                        if (tryParse)
                        {
                            sumValue += result;
                        }
                    }
                }
                sheet.CreateRow(sheet.LastRowNum + 1);
                for (int col = 2; col < 5; col++)
                {
                    sheet.GetRow(sheet.LastRowNum).CreateCell(col);
                    sheet.GetRow(sheet.LastRowNum).GetCell(col).CellStyle = contentFormat;
                }
                sheet.AddMergedRegion(new CellRangeAddress(sheet.LastRowNum, sheet.LastRowNum, 2, 3));
                sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue("總計");
                sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(sumValue);
                #endregion

                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report04010203" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }

        }
        #endregion

        #region Report04010204 VD卡片-卡片調整統計表
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Excel_Report04010204(VD卡片/統計表或關係表查詢/卡片調整統計表)
        /// 作    者:Ares Stanley
        /// 創建時間:2021/12/13
        /// </summary>
        /// <param name="strName"></param>
        /// <param name="strRptFld"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="agentId"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report04010204(string strName, string strRptFld, string strRptBeforeDate, string strRptEndDate, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "card.xls";
                int totalCount = 0;
                DataTable dt = getData_Comm(agentId, string.Format(sqlComm_01010204, UtilHelper.GetAppSettings("DB_CP_DBF")), ref totalCount, "04010204", 0, true, "P");
                if (dt.Rows.Count <= 0)
                    return false;
                DataTable dt2 = getData_Comm(agentId, sqlComm_01010204_Count, ref totalCount, "04010204", 0, false);
                string totalDiffCount = "";
                if (dt2.Rows.Count > 0)
                {
                    totalDiffCount = dt2.Rows[0][0].ToString();
                }
                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 表頭
                sheet.GetRow(1).GetCell(1).SetCellValue("VD卡片調整統計表");
                sheet.GetRow(4).GetCell(0).SetCellValue("維護欄位：" + strRptFld);//維護欄位
                sheet.GetRow(5).GetCell(0).SetCellValue("維護日期：" + strRptBeforeDate + " ~ " + strRptEndDate);//維護日期
                sheet.GetRow(4).GetCell(5).SetCellValue("製表日：" + DateTime.Now.ToString("yyyy/MM/dd"));//製表日
                sheet.GetRow(5).GetCell(5).SetCellValue("製表人：" + strName);//製表人
                sheet.GetRow(6).GetCell(2).SetCellValue("卡號");
                #endregion


                #region 表身

                //資料去空白
                removeBlank(ref dt);

                //資料寫入
                ExportExcelForNPOI(dt, ref wb, 7, "工作表1", 1);

                #region 合併相同資料
                int startRow = 0;
                int endRow = 0;
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    if (row == sheet.LastRowNum)
                    {
                        if (startRow > 0)
                        {
                            endRow = row;
                        }
                        if (endRow - startRow >= 1)
                        {
                            sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 1, 1));
                            sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                            sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                            startRow = 0;
                            endRow = 0;
                        }
                        break;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() == sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString())
                    {
                        if (startRow != 0)
                            continue;
                        startRow = row;
                        continue;
                    }

                    if (sheet.GetRow(row).GetCell(2).StringCellValue.ToString() != sheet.GetRow(row + 1).GetCell(2).StringCellValue.ToString() || (row == sheet.LastRowNum - 1 && startRow > 0))
                    {
                        if (startRow != 0)
                        {
                            endRow = row;

                            if (endRow - startRow >= 1)
                            {
                                sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 1, 1));
                                sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));
                                sheet.GetRow(startRow).GetCell(2).CellStyle = contentFormat;
                                startRow = 0;
                                endRow = 0;
                                continue;
                            }
                        }
                    }
                }
                #endregion

                //增加尾列總計
                int sumValue = NPOI_ColumnSum(sheet, 7, sheet.LastRowNum);
                sheet.CreateRow(sheet.LastRowNum + 1);
                for (int c = 1; c < 6; c++)
                {
                    sheet.GetRow(sheet.LastRowNum).CreateCell(c);
                    sheet.GetRow(sheet.LastRowNum).GetCell(c).CellStyle = contentFormat;
                }
                sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue("總計");
                sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(totalDiffCount);
                sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(sumValue);

                //小計欄位合併
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    sheet.GetRow(row).CreateCell(5);
                    sheet.GetRow(row).GetCell(5).CellStyle = contentFormat;
                }
                for (int row = 7; row < sheet.LastRowNum + 1; row++)
                {
                    sheet.AddMergedRegion(new CellRangeAddress(row, row, 4, 5));
                    sheet.GetRow(row).GetCell(4).CellStyle = contentFormat;
                }
                //額外合併尾列維護員、小計
                sheet.AddMergedRegion(new CellRangeAddress(sheet.LastRowNum, sheet.LastRowNum, 3, 4));
                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report04010204" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }
        }



        #endregion

        #region Report04010301 VD卡片-調整固定額度
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Report04010301(VD卡片/額度查詢/調整固定額度)資料並產出Excel
        /// 作    者:Ares Stanley
        /// 創建時間:2021/12/13
        /// </summary>
        /// <param name="strRptPeople"></param>
        /// <param name="strRptBeforeAmount"></param>
        /// <param name="strRptEndAmount"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="strAgentName"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report04010301(string strRptPeople, string strRptBeforeAmount, string strRptEndAmount, string strRptBeforeDate, string strRptEndDate, string strAgentName, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "member1.xls";
                int totalCount = 0;
                DataTable dt = getData_member1(agentId, ref totalCount, "04010301", 0, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");
                #region 表頭
                sheet.GetRow(0).GetCell(0).SetCellValue("VD調整固定額度");
                sheet.GetRow(2).GetCell(0).SetCellValue(string.Format("維護員 : {0}", strRptPeople));
                sheet.GetRow(3).GetCell(0).SetCellValue(string.Format("額度 : {0} ~ {1}", strRptBeforeAmount, strRptEndAmount));
                sheet.GetRow(3).GetCell(6).SetCellValue(DateTime.Now.ToString("yyyyMMdd"));//製表日
                sheet.GetRow(4).GetCell(0).SetCellValue(string.Format("維護日期 : {0} ~ {1}", strRptBeforeDate, strRptEndDate));
                sheet.GetRow(4).GetCell(6).SetCellValue(strAgentName);//製表人
                sheet.GetRow(7).GetCell(0).SetCellValue("卡號");
                #endregion

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 資料
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sheet.CreateRow(sheet.LastRowNum + 1);
                    for (int b = 0; b < 7; b++)
                    {
                        sheet.GetRow(sheet.LastRowNum).CreateCell(b);
                        sheet.GetRow(sheet.LastRowNum).GetCell(b).CellStyle = contentFormat;
                    }
                    sheet.GetRow(sheet.LastRowNum).GetCell(0).SetCellValue(dt.Rows[i]["CUST_ID"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue(dt.Rows[i]["FLD_NAME"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(dt.Rows[i]["BEFOR_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(3).SetCellValue(dt.Rows[i]["AFTER_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(dt.Rows[i]["MAINT_D"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(5).SetCellValue(dt.Rows[i]["MAINT_T"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(6).SetCellValue(dt.Rows[i]["USER_ID"].ToString().Trim());
                }
                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report04010301" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }
        }
        #endregion

        #region Report04010302 VD卡片-調整臨時額度
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Report04010302(VD卡片/額度查詢/調整臨時額度)資料並產出Excel
        /// 作    者:Ares Stanley
        /// 創建時間:2021/12/13
        /// </summary>
        /// <param name="strRptPeople"></param>
        /// <param name="strRptBeforeAmount"></param>
        /// <param name="strRptEndAmount"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="strAgentName"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report04010302(string strRptPeople, string strRptBeforeAmount, string strRptEndAmount, string strRptBeforeDate, string strRptEndDate, string strAgentName, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "member1.xls";
                int totalCount = 0;
                DataTable dt = getData_member1(agentId, ref totalCount, "04010302", 0, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");
                #region 表頭
                sheet.GetRow(0).GetCell(0).SetCellValue("VD調整臨時額度");
                sheet.GetRow(2).GetCell(0).SetCellValue(string.Format("維護員 : {0}", strRptPeople));
                sheet.GetRow(3).GetCell(0).SetCellValue(string.Format("額度 : {0} ~ {1}", strRptBeforeAmount, strRptEndAmount));
                sheet.GetRow(3).GetCell(6).SetCellValue(DateTime.Now.ToString("yyyyMMdd"));//製表日
                sheet.GetRow(4).GetCell(0).SetCellValue(string.Format("維護日期 : {0} ~ {1}", strRptBeforeDate, strRptEndDate));
                sheet.GetRow(4).GetCell(6).SetCellValue(strAgentName);//製表人
                sheet.GetRow(7).GetCell(0).SetCellValue("卡號");
                #endregion

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 資料
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sheet.CreateRow(sheet.LastRowNum + 1);
                    for (int b = 0; b < 7; b++)
                    {
                        sheet.GetRow(sheet.LastRowNum).CreateCell(b);
                        sheet.GetRow(sheet.LastRowNum).GetCell(b).CellStyle = contentFormat;
                    }
                    sheet.GetRow(sheet.LastRowNum).GetCell(0).SetCellValue(dt.Rows[i]["CUST_ID"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue(dt.Rows[i]["FLD_NAME"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(dt.Rows[i]["BEFOR_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(3).SetCellValue(dt.Rows[i]["AFTER_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(dt.Rows[i]["MAINT_D"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(5).SetCellValue(dt.Rows[i]["MAINT_T"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(6).SetCellValue(dt.Rows[i]["USER_ID"].ToString().Trim());
                }
                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report04010302" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }
        }
        #endregion

        #region Report04010303 VD卡片-新卡額度
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Report04010303(VD卡片/額度查詢/新卡額度)資料並產出Excel
        /// 作    者:Ares Stanley
        /// 創建時間:2021/12/13
        /// </summary>
        /// <param name="strRptPeople"></param>
        /// <param name="strRptBeforeAmount"></param>
        /// <param name="strRptEndAmount"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="strAgentName"></param>
        /// <param name="agentId"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report04010303(string strRptPeople, string strRptBeforeAmount, string strRptEndAmount, string strRptBeforeDate, string strRptEndDate, string strAgentName, string agentId, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "member1.xls";
                int totalCount = 0;
                DataTable dt = getData_member1(agentId, ref totalCount, "04010303", 0, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");
                #region 表頭
                sheet.GetRow(0).GetCell(0).SetCellValue("VD新卡額度");
                sheet.GetRow(2).GetCell(0).SetCellValue(string.Format("維護員 : {0}", strRptPeople));
                sheet.GetRow(3).GetCell(0).SetCellValue(string.Format("額度 : {0} ~ {1}", strRptBeforeAmount, strRptEndAmount));
                sheet.GetRow(3).GetCell(6).SetCellValue(DateTime.Now.ToString("yyyyMMdd"));//製表日
                sheet.GetRow(4).GetCell(0).SetCellValue(string.Format("維護日期 : {0} ~ {1}", strRptBeforeDate, strRptEndDate));
                sheet.GetRow(4).GetCell(6).SetCellValue(strAgentName);//製表人
                sheet.GetRow(7).GetCell(0).SetCellValue("卡號");
                #endregion

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 資料
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sheet.CreateRow(sheet.LastRowNum + 1);
                    for (int b = 0; b < 7; b++)
                    {
                        sheet.GetRow(sheet.LastRowNum).CreateCell(b);
                        sheet.GetRow(sheet.LastRowNum).GetCell(b).CellStyle = contentFormat;
                    }
                    sheet.GetRow(sheet.LastRowNum).GetCell(0).SetCellValue(dt.Rows[i]["CUST_ID"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue(dt.Rows[i]["FLD_NAME"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(dt.Rows[i]["BEFOR_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(3).SetCellValue(dt.Rows[i]["AFTER_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(dt.Rows[i]["MAINT_D"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(5).SetCellValue(dt.Rows[i]["MAINT_T"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(6).SetCellValue(dt.Rows[i]["USER_ID"].ToString().Trim());
                }
                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report04010303" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }
        }
        #endregion

        #region Report04010401 VD卡片-非流通BlockCode調整為流通中
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Report04010401(VD卡片/特殊查詢/非流通BlockCode調整為流通中)資料並產出Excel
        /// 作    者:Ares Stanley
        /// 創建時間:2021/12/13
        /// </summary>
        /// <param name="strRptPeople"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="strAgentName"></param>
        /// <param name="agentId"></param>
        /// <param name="strRptBeforeContent"></param>
        /// <param name="strRptEndContent"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report04010401(string strRptPeople, string strRptBeforeDate, string strRptEndDate, string strAgentName, string agentId, string strRptBeforeContent, string strRptEndContent, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "member1.xls";
                int totalCount = 0;
                DataTable dt = getData_member1(agentId, ref totalCount, "04010401", 0, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");
                #region 表頭
                sheet.GetRow(0).GetCell(0).SetCellValue("VD非流通BlockCode調整為流通中");
                sheet.GetRow(2).GetCell(0).SetCellValue(string.Format("調整前內容 : {0}", strRptBeforeContent));
                sheet.GetRow(3).GetCell(0).SetCellValue(string.Format("調整後內容 : {0}", strRptEndContent));
                sheet.GetRow(3).GetCell(6).SetCellValue(DateTime.Now.ToString("yyyyMMdd"));//製表日
                sheet.GetRow(4).GetCell(0).SetCellValue(string.Format("維護員 : {0}", strRptPeople));
                sheet.GetRow(4).GetCell(6).SetCellValue(strAgentName);//製表人
                sheet.GetRow(5).GetCell(0).SetCellValue(string.Format("維護日期 : {0} ~ {1}", strRptBeforeDate, strRptEndDate));
                sheet.GetRow(7).GetCell(0).SetCellValue("卡號");
                #endregion

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 資料
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sheet.CreateRow(sheet.LastRowNum + 1);
                    for (int b = 0; b < 7; b++)
                    {
                        sheet.GetRow(sheet.LastRowNum).CreateCell(b);
                        sheet.GetRow(sheet.LastRowNum).GetCell(b).CellStyle = contentFormat;
                    }
                    sheet.GetRow(sheet.LastRowNum).GetCell(0).SetCellValue(dt.Rows[i]["CUST_ID"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue(dt.Rows[i]["FLD_NAME"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(dt.Rows[i]["BEFOR_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(3).SetCellValue(dt.Rows[i]["AFTER_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(dt.Rows[i]["MAINT_D"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(5).SetCellValue(dt.Rows[i]["MAINT_T"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(6).SetCellValue(dt.Rows[i]["USER_ID"].ToString().Trim());
                }
                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report04010401" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }
        }
        #endregion

        #region Report04010402 VD卡片-年費無優惠調整為優惠條件
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Report04010402(VD卡片/特殊查詢/年費無優惠調整為優惠條件)資料並產出Excel
        /// 作    者:Ares Stanley
        /// 創建時間:2021/12/13
        /// </summary>
        /// <param name="strRptPeople"></param>
        /// <param name="strRptBeforeDate"></param>
        /// <param name="strRptEndDate"></param>
        /// <param name="strAgentName"></param>
        /// <param name="agentId"></param>
        /// <param name="strRptBeforeContent"></param>
        /// <param name="strRptEndContent"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_Report04010402(string strRptPeople, string strRptBeforeDate, string strRptEndDate, string strAgentName, string agentId, string strRptBeforeContent, string strRptEndContent, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "member1.xls";
                int totalCount = 0;
                DataTable dt = getData_member1(agentId, ref totalCount, "04010402", 0, "P");
                if (dt.Rows.Count <= 0)
                    return false;

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");
                #region 表頭
                sheet.GetRow(0).GetCell(0).SetCellValue("VD年費無優惠調整爲優惠條件");
                sheet.GetRow(2).GetCell(0).SetCellValue(string.Format("調整前內容 : {0}", strRptBeforeContent));
                sheet.GetRow(3).GetCell(0).SetCellValue(string.Format("調整后內容 : {0}", strRptEndContent));
                sheet.GetRow(3).GetCell(6).SetCellValue(DateTime.Now.ToString("yyyyMMdd"));//製表日
                sheet.GetRow(4).GetCell(0).SetCellValue(string.Format("維護員 : {0}", strRptPeople));
                sheet.GetRow(4).GetCell(6).SetCellValue(strAgentName);//製表人
                sheet.GetRow(5).GetCell(0).SetCellValue(string.Format("維護日期 : {0} ~ {1}", strRptBeforeDate, strRptEndDate));
                sheet.GetRow(7).GetCell(0).SetCellValue("卡號");
                #endregion

                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 資料
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sheet.CreateRow(sheet.LastRowNum + 1);
                    for (int b = 0; b < 7; b++)
                    {
                        sheet.GetRow(sheet.LastRowNum).CreateCell(b);
                        sheet.GetRow(sheet.LastRowNum).GetCell(b).CellStyle = contentFormat;
                    }
                    sheet.GetRow(sheet.LastRowNum).GetCell(0).SetCellValue(dt.Rows[i]["CUST_ID"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue(dt.Rows[i]["FLD_NAME"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(dt.Rows[i]["BEFOR_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(3).SetCellValue(dt.Rows[i]["AFTER_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(dt.Rows[i]["MAINT_D"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(5).SetCellValue(dt.Rows[i]["MAINT_T"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(6).SetCellValue(dt.Rows[i]["USER_ID"].ToString().Trim());
                }
                #endregion

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report04010402" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return false;
            }
        }
        #endregion

        #endregion

        #region 匯入紀錄
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Report05010000(匯入紀錄查詢)資料並產出Excel
        /// 作    者:Ares Stanley
        /// 創建時間:2021/12/13
        /// </summary>
        /// <param name="strBeforeData"></param>
        /// <param name="strEndData"></param>
        /// <param name="sqlCondition"></param>
        /// <param name="strAgentName"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_05010000(string strBeforeData, string strEndData, string sqlCondition, string strAgentName, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string connection = UtilHelper.GetConnectionStrings("Connection_System");
                SqlConnection sql_conn = new SqlConnection(connection);
                DataTable dt = new DataTable();
                SqlCommand sqlComm = new SqlCommand();
                Dictionary<string, string> commandParameters = new Dictionary<string, string>();
                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "importLog.xls";
                if (!string.IsNullOrEmpty(strBeforeData) && !string.IsNullOrEmpty(strEndData))
                {
                    //有輸入日期
                    sqlComm.CommandText = sqlComm_05010000_withCondition;
                    sqlComm.Parameters.Add(new SqlParameter("@dateStart", strBeforeData));
                    sqlComm.Parameters.Add(new SqlParameter("@dateEnd", strEndData));
                    commandParameters.Add("@dateStart", strBeforeData);
                    commandParameters.Add("@dateEnd", strEndData);
                }
                else
                {
                    //無輸入日期
                    sqlComm.CommandText = sqlComm_05010000_withoutCondition;
                }



                DataHelper dh = new DataHelper();
                Stopwatch sw = new Stopwatch();
                sw.Start();
                DataSet ds = dh.ExecuteDataSet(sqlComm);
                sw.Stop();

                if (ds.Tables.Count > 0)
                {
                    dt = ds.Tables[0];
                }

                //紀錄SQL
                PrintSQL(sqlComm.CommandText, sw, commandParameters, "05010000", dt, "P");

                if (dt.Rows.Count <= 0)
                {
                    strMsgID = "00_00000000_037";//無資料
                    return false;
                }

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");
                #region 表頭
                sheet.GetRow(4).GetCell(0).SetCellValue(string.Format("匯入日期 : {0}~{1}", strBeforeData, strEndData));//匯入日期
                sheet.GetRow(4).GetCell(4).SetCellValue(string.Format("製表日：{0}", DateTime.Now.ToString("yyyyMMdd")));//製表日
                sheet.GetRow(5).GetCell(4).SetCellValue(string.Format("製表人：{0}", strAgentName));//製表人
                #endregion

                ExportExcelForNPOI(dt, ref wb, 7, "工作表1");

                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report05010000" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                strMsgID = "00_00000000_038";//下載失敗
                return false;
            }
        }

        #endregion

        #region 匯入紀錄明細
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:產出Report05010000Detail(匯入紀錄查詢明細)資料並產出Excel
        /// 作    者:Ares Stanley
        /// 創建時間:2021/12/13
        /// </summary>
        /// <param name="strBeforeData"></param>
        /// <param name="strEndData"></param>
        /// <param name="sqlCondition"></param>
        /// <param name="strAgentName"></param>
        /// <param name="strPathFile"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool CreateExcelFile_05010000Detail(string fileName, string tableName, string importDate, string strAgentName, ref string strPathFile, ref string strMsgID)
        {
            try
            {
                // 檢查目錄，并刪除以前的文檔資料
                CheckDirectory(ref strPathFile);

                string connection = UtilHelper.GetConnectionStrings("Connection_System");
                SqlConnection sql_conn = new SqlConnection(connection);
                DataTable dt = new DataTable();
                SqlCommand sqlComm = new SqlCommand();

                string strExcelPathFile = AppDomain.CurrentDomain.BaseDirectory + UtilHelper.GetAppSettings("ReportTemplate") + "importLogDetail.xls";

                sqlComm.CommandText = string.Format(sqlComm_05010000Detail, tableName);
                sqlComm.Parameters.Add(new SqlParameter("@EXE_Name", fileName));

                //紀錄SQL
                Dictionary<string, string> commandParameters = new Dictionary<string, string>();
                commandParameters.Add("@EXE_Name", fileName);

                DataHelper dh = new DataHelper();
                Stopwatch sw = new Stopwatch();
                sw.Start();
                DataSet ds = dh.ExecuteDataSet(sqlComm);
                sw.Stop();
                if (ds.Tables.Count > 0)
                {
                    dt = ds.Tables[0];
                }

                PrintSQL(sqlComm.CommandText, sw, commandParameters, "05010000_Detail", dt, "P");

                if (dt.Rows.Count <= 0)
                {
                    strMsgID = "00_00000000_037";//無資料
                    return false;
                }

                FileStream fs = new FileStream(strExcelPathFile, FileMode.Open);
                HSSFWorkbook wb = new HSSFWorkbook(fs);
                ISheet sheet = wb.GetSheet("工作表1");
                #region 表頭
                sheet.GetRow(4).GetCell(0).SetCellValue(string.Format("匯入日期 : {0}", importDate));//匯入日期
                sheet.GetRow(4).GetCell(6).SetCellValue(string.Format("製表日：{0}", DateTime.Now.ToString("yyyyMMdd")));//製表日
                sheet.GetRow(5).GetCell(0).SetCellValue(string.Format("檔名：{0}", fileName));//檔名
                sheet.GetRow(5).GetCell(6).SetCellValue(string.Format("製表人：{0}", strAgentName));//製表人
                #endregion

                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                #region 資料
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sheet.CreateRow(sheet.LastRowNum + 1);
                    for (int b = 0; b < 7; b++)
                    {
                        sheet.GetRow(sheet.LastRowNum).CreateCell(b);
                        sheet.GetRow(sheet.LastRowNum).GetCell(b).CellStyle = contentFormat;
                    }
                    sheet.GetRow(sheet.LastRowNum).GetCell(0).SetCellValue(dt.Rows[i]["CUST_ID"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(1).SetCellValue(dt.Rows[i]["FLD_NAME"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(2).SetCellValue(dt.Rows[i]["BEFOR_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(3).SetCellValue(dt.Rows[i]["AFTER_UPD"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(4).SetCellValue(dt.Rows[i]["MAINT_D"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(5).SetCellValue(dt.Rows[i]["MAINT_T"].ToString().Trim());
                    sheet.GetRow(sheet.LastRowNum).GetCell(6).SetCellValue(dt.Rows[i]["USER_ID"].ToString().Trim());
                }
                #endregion
                // 保存文件到運行目錄下
                strPathFile = strPathFile + @"\ExcelFile_Report05010000Detail" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                FileStream fs1 = new FileStream(strPathFile, FileMode.Create);
                wb.Write(fs1);
                fs1.Close();
                fs.Close();
                return true;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                strMsgID = "00_00000000_038";//下載失敗
                return false;
            }
        }

        #endregion

        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:檢查路徑是否存在，存在刪除該路徑下所有的文檔資料
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/10
        /// </summary>
        /// <param name="strPath"></param>
        public static void CheckDirectory(ref string strPath)
        {
            try
            {
                string strOldPath = strPath;
                //* 判斷路徑是否存在
                strPath = strPath + "\\" + DateTime.Now.ToString("yyyyMMdd");
                if (!Directory.Exists(strPath))
                {
                    //* 如果不存在，創建路徑
                    Directory.CreateDirectory(strPath);
                }

                //* 取該路徑下所有路徑
                string[] strDirectories = Directory.GetDirectories(strOldPath);
                for (int intLoop = 0; intLoop < strDirectories.Length; intLoop++)
                {
                    if (strDirectories[intLoop].ToString() != strPath)
                    {
                        if (Directory.Exists(strDirectories[intLoop]))
                        {
                            // * 刪除目錄下的所有文檔
                            DirectoryInfo di = new DirectoryInfo(strDirectories[intLoop]);
                            FileSystemInfo[] fsi = di.GetFileSystemInfos();
                            for (int intIndex = 0; intIndex < fsi.Length; intIndex++)
                            {
                                FileInfo fi = fsi[intIndex] as FileInfo;
                                if (fi != null)
                                {
                                    fi.Delete();
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception exp)
            {
                Logging.Log(exp, LogLayer.BusinessRule);
                throw exp;
            }
        }

        #region 共用NPOI
        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:按資料排序並產出Excel
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/10
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="wb"></param>
        /// <param name="start"></param>
        /// <param name="sheetName"></param>
        private static void ExportExcelForNPOI(DataTable dt, ref HSSFWorkbook wb, Int32 startRow, String sheetName, Int32 startCell = 0)
        {
            try
            {
                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                if (dt != null && dt.Rows.Count != 0)
                {
                    int count = startRow;
                    ISheet sheet = wb.GetSheet(sheetName);
                    int cols = dt.Columns.Count;
                    foreach (DataRow dr in dt.Rows)
                    {
                        int cell = startCell;
                        IRow row = (IRow)sheet.CreateRow(count);
                        //row.CreateCell(0).SetCellValue(count.ToString());
                        for (int i = 0; i < cols; i++)
                        {
                            row.CreateCell(cell).SetCellValue(dr[i].ToString());
                            row.GetCell(cell).CellStyle = contentFormat;
                            cell++;
                        }
                        count++;
                    }
                }
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                throw;
            }
        }

        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:按資料排序並產出Excel(含小計列)
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/10
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="wb"></param>
        /// <param name="start"></param>
        /// <param name="sheetName"></param>
        private static void ExportExcelForNPOI_SubTotal(DataTable dt, ref HSSFWorkbook wb, Int32 start, String sheetName)
        {
            try
            {
                //取得樣式
                HSSFCellStyle contentFormat = getDefaultContentFormat(wb);

                if (dt != null && dt.Rows.Count != 0)
                {
                    int count = start;
                    ISheet sheet = wb.GetSheet(sheetName);
                    int cols = dt.Columns.Count;
                    foreach (DataRow dr in dt.Rows)
                    {
                        int cell = 2;
                        #region 加入小計列
                        if (sheet.GetRow(count - 1).GetCell(2).StringCellValue.ToString() != dr[0].ToString() && count != start)
                        {
                            IRow rowSum = (IRow)sheet.CreateRow(count);

                            for (int i = 0; i < cols; i++)
                            {
                                rowSum.CreateCell(cell).SetCellValue("");
                                rowSum.GetCell(cell).CellStyle = contentFormat;
                                cell++;
                            }
                            cell = 2;
                            count++;
                        }
                        #endregion

                        IRow row = (IRow)sheet.CreateRow(count);
                        //row.CreateCell(0).SetCellValue(count.ToString());
                        for (int i = 0; i < cols; i++)
                        {
                            row.CreateCell(cell).SetCellValue(dr[i].ToString());
                            row.GetCell(cell).CellStyle = contentFormat;
                            cell++;
                        }
                        count++;
                    }
                    //新增尾列
                    sheet.CreateRow(sheet.LastRowNum + 1);
                    for (int i = 2; i < 5; i++)
                    {
                        sheet.GetRow(sheet.LastRowNum).CreateCell(i);
                        sheet.GetRow(sheet.LastRowNum).GetCell(i).CellStyle = contentFormat;
                    }
                }
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                throw;
            }
        }

        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:按資料排序並產出Excel(可移除末N欄不需要的資料)
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/10
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="wb"></param>
        /// <param name="start"></param>
        /// <param name="delColumn">移除欄位數量</param>
        /// <param name="sheetName"></param>
        private static void ExportExcelForNPOI_filter(DataTable dt, ref HSSFWorkbook wb, Int32 start, Int32 delColumn, String sheetName)
        {
            try
            {
                HSSFCellStyle cs = (HSSFCellStyle)wb.CreateCellStyle();
                cs.BorderBottom = BorderStyle.Thin;
                cs.BorderLeft = BorderStyle.Thin;
                cs.BorderTop = BorderStyle.Thin;
                cs.BorderRight = BorderStyle.Thin;

                //啟動多行文字
                cs.WrapText = true;
                //文字置中
                cs.VerticalAlignment = VerticalAlignment.Center;
                cs.Alignment = HorizontalAlignment.Center;

                HSSFFont font1 = (HSSFFont)wb.CreateFont();
                //字體尺寸
                font1.FontHeightInPoints = 12;
                font1.FontName = "新細明體";
                cs.SetFont(font1);

                if (dt != null && dt.Rows.Count != 0)
                {
                    int count = start;
                    ISheet sheet = wb.GetSheet(sheetName);
                    int cols = dt.Columns.Count - delColumn;
                    foreach (DataRow dr in dt.Rows)
                    {
                        int cell = 0;
                        IRow row = (IRow)sheet.CreateRow(count);
                        row.CreateCell(0).SetCellValue(count.ToString());
                        for (int i = 0; i < cols; i++)
                        {
                            row.CreateCell(cell).SetCellValue(dr[i].ToString());
                            row.GetCell(cell).CellStyle = cs;
                            cell++;
                        }
                        count++;
                    }
                }
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                throw;
            }
        }
        #endregion


        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:添加維護員統計小計資料
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/29
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startRow"></param>
        /// <param name="endRow"></param>
        /// <param name="cs"></param>
        private static void NPOI_AddSubTotal(ISheet sheet, int startRow, int endRow, HSSFCellStyle cs)
        {
            try
            {
                sheet.CreateRow(endRow + 1);
                for (int i = 2; i < 5; i++)
                {
                    sheet.GetRow(endRow + 1).CreateCell(i);
                    sheet.GetRow(endRow + 1).GetCell(i).CellStyle = cs;
                }
                sheet.GetRow(endRow + 1).GetCell(2).SetCellValue(sheet.GetRow(startRow).GetCell(2).StringCellValue.ToString());
                sheet.GetRow(endRow + 1).GetCell(3).SetCellValue("小計");

                int sumValue = 0;
                for (int i = startRow; i < (endRow + 1); i++)
                {
                    int result = 0;
                    bool tryParse = int.TryParse(sheet.GetRow(i).GetCell(4).StringCellValue.ToString(), out result);
                    if (tryParse)
                    {
                        sumValue += result;
                    }
                }
                sheet.GetRow(endRow + 1).GetCell(4).SetCellValue(sumValue);

            }
            catch (Exception ex)
            {
                Logging.Log(ex);
            }
        }

        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:尾欄總計
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/29
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startRow"></param>
        /// <param name="endRow"></param>
        /// <returns></returns>
        private static int NPOI_ColumnSum(ISheet sheet, int startRow, int endRow)
        {
            try
            {
                int sumValue = 0;
                for (int row = startRow; row < (endRow + 1); row++)
                {
                    int result = 0;
                    bool tryParse = int.TryParse(sheet.GetRow(row).GetCell(4).StringCellValue.ToString(), out result);
                    if (tryParse)
                    {
                        sumValue += result;
                    }
                }
                return sumValue;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return 0;
            }
        }

        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:移除資料空白
        /// 作    者:Ares Stanley
        /// 創建時間:2021/11/08
        /// </summary>
        /// <param name="dt"></param>
        public static void removeBlank(ref DataTable dt)
        {
            foreach (DataRow dr in dt.Rows)
            {
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    dr[i] = dr[i].ToString().Trim();
                }
            }
        }

        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:紀錄查詢SQL
        /// 作    者:Ares Stanley
        /// 創建時間:2021/12/06
        /// </summary>
        /// <param name="SQLCommand"></param>
        /// <param name="commandParameters"></param>
        public static void PrintSQL(string SQLCommand, Stopwatch sw, Dictionary<string, string> commandParameters = null, string reportName = "", DataTable dt = null, string queryTypeName = "")
        {
            int dataCount = 0;
            string userID = string.Empty;
            if (commandParameters != null)
            {
                foreach (KeyValuePair<string, string> param in commandParameters)
                {
                    SQLCommand = SQLCommand.Replace(param.Key, $"'{param.Value}'");
                    userID = param.Value;
                }
            }

            if (dt != null)
            {
                dataCount = dt.Rows.Count;
            }

            TimeSpan ts = sw.Elapsed;
            Logging.Log($"直接撈取[Rpt_CPMAST]該{userID}的所有資料顯示於頁面 \n" + SQLCommand);
            Logging.Log($"執行結果共 {dataCount} 筆，共花費 {ts.TotalMilliseconds} ms");
            Logging.Log($"====================報表 {reportName}  [{queryTypeName}]  紀錄結束==================== \n", LogState.Info, LogLayer.Util);
        }

        /// <summary>
        /// 專案代號:20210058-CSIP作業服務平台現代化II
        /// 功能說明:取得共用報表樣式
        /// 作    者:Ares Stanley
        /// 創建時間:2021/12/10
        /// </summary>
        /// <param name="wb"></param>
        /// <returns></returns>
        public static HSSFCellStyle getDefaultContentFormat(HSSFWorkbook wb)
        {
            HSSFCellStyle contentFormat = (HSSFCellStyle)wb.CreateCellStyle(); //建立文字格式
            try
            {
                contentFormat.VerticalAlignment = VerticalAlignment.Top; //垂直置中
                //contentFormat.Alignment = HorizontalAlignment.Center; //水平置中
                contentFormat.DataFormat = HSSFDataFormat.GetBuiltinFormat("@"); //將儲存格內容設定為文字
                contentFormat.BorderBottom = BorderStyle.Thin; // 儲存格框線
                contentFormat.BorderLeft = BorderStyle.Thin;
                contentFormat.BorderTop = BorderStyle.Thin;
                contentFormat.BorderRight = BorderStyle.Thin;

                HSSFFont contentFont = (HSSFFont)wb.CreateFont(); //建立文字樣式
                contentFont.FontHeightInPoints = 12; //字體大小
                contentFont.FontName = "新細明體"; //字型
                contentFormat.SetFont(contentFont); //設定儲存格的文字樣式
                return contentFormat;
            }
            catch (Exception ex)
            {
                Logging.Log(ex);
                return contentFormat;
            }

        }

    }

}
