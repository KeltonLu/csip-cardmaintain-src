//******************************************************************
//*  作    者：Ares JACK
//*  創建日期：2021/08/03
//*  功能說明：搭配參數維護,刪除一年前的紀錄
//                  1.清除使用者紀錄, 
//                  2.清除JOB錯誤資訊,
//                  3.操作紀錄, 
//                  4.使用者連線ID, 
//                  5.功能歷程記錄, 
//                  6.動作歷程記錄,
//                  7.屬性歷程記錄, 
//                  8.角色權限歷程記錄, 
//                  9.角色歷程記錄, 
//                  10.工作日歷程記錄,
//                  11.AP執行紀錄, 
//                  12.排程執行紀錄
//* 修改紀錄：2021/08/12_Ares JACK 刪除一年改為可手動輸入天數
//*<author>            <time>            <TaskID>                <desc>
//*******************************************************************

using System;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Diagnostics;
using Framework.Data;
using Framework.Common.Utility;
using Framework.Common.Logging;
using CSIPCommonModel.BusinessRules;
using Quartz;
using System.Linq;

/// <summary>
/// DB HouseKeeping 的摘要描述
/// </summary>
public class JobDEL_HistoryData_3 : IJob
{
    public JobDEL_HistoryData_3()
    {
        //
        // TODO: 在此加入建構函式的程式碼
        //
    }

    protected string StrJobId;
    protected string StrFunctionKey = UtilHelper.GetAppSettings("FunctionKey");
    protected JobHelper JobHelper = new JobHelper();
    protected DateTime StartTime;
    protected DateTime EndTime;
    protected string StrMailMsg = string.Empty;

    private readonly int BatchNum = 1000;

    private readonly string _strFrom = UtilHelper.GetAppSettings("MailSender");
    private string strSucc = "S";// 運行成功失敗

    #region IJob 程式入口

    /// <summary>
    /// Job 調用入口
    /// 修改紀錄:2021/08/03_Ares_Jack-調整查詢條件的SQL語法
    /// </summary>
    /// <param name="context"></param>
    public void Execute(JobExecutionContext context)
    {
        var mailTo = context.JobDetail.JobDataMap.GetString("mail");
        var strSubject = context.JobDetail.JobDataMap.GetString("title");


        try
        {
            #region 獲取JOBID

            StrJobId = context.JobDetail.JobDataMap["jobid"].ToString().Trim();
            JobHelper.strJobID = StrJobId;

            #endregion

            #region 記錄job啟動時間

            StartTime = DateTime.Now;
            JobHelper.SaveLog(StrJobId + ", HOUSE_KEEPING 啟動！", LogState.Info);

            #endregion

            #region 判斷job工作狀態

            if (JobHelper.SerchJobStatus(StrJobId) == "" || JobHelper.SerchJobStatus(StrJobId) == "0")
            {
                JobHelper.SaveLog("JOB 工作狀態為：停止！", LogState.Info);
                return;
            }

            #endregion

            #region 檢測JOB是否在執行中

            if (BusinessRules.BRM_LBatchLog.JobStatusChk(StrFunctionKey, StrJobId, DateTime.Now))
            {
                JobHelper.SaveLog("JOB 工作狀態為：正在執行！", LogState.Info);
                // 返回不在執行           
                return;
            }
            else
            {
                BusinessRules.BRM_LBatchLog.Insert(StrFunctionKey, StrJobId, StartTime, "R", "JOB 工作狀態為：開始執行！");
            }

            #endregion

            // 預計刪除總筆數,批次次數
            int totalNum = 0;
            decimal runNum = 0;
            // 實際刪除筆數計算
            int totalDeleteCount = 0;
            // 資料庫
            //string dbName = "";
            SqlCommand sql = new SqlCommand();

            string tempMsg = "";

            // 計時器
            Stopwatch sw;
            
            ConsumeTimeStart(sw = new Stopwatch());
            #region 紀錄刪除------------------------------------------------------------------------------------------

            DataTable dt = new DataTable();
            if (!CSIPCommonModel.BusinessRules.BRM_PROPERTY_KEY.GetEnableProperty("03", "HouseKeepingData", ref dt))
            {
                //報錯訊息
                Logging.Log("資料錯誤", StrJobId, LogState.Error);
                strSucc = "F";
                StrMailMsg += "資料錯誤 ";
                return ;
            }

            foreach(DataRow dr in dt.Rows)
            {
                //判斷參數數量是否正確
                if (dr["PROPERTY_NAME"].ToString().Count(f => (f == ',')) != 3)
                {
                    Logging.Log("------------------------------------------------------------", StrJobId, LogState.Info);
                    Logging.Log("格式錯誤,需用逗號分隔內容!", StrJobId, LogState.Error);
                    strSucc = "F";
                    StrMailMsg += "格式錯誤,需用逗號分隔內容! ";
                    continue;
                }

                //拆分變數陣列,逗點分隔
                string[] splitName = dr["PROPERTY_NAME"].ToString().Split(',');
                string propertyDBName = splitName[0];
                string propertyTableName = splitName[1];
                string propertyColumName = splitName[2];
                string propertyDeleteDays = splitName[3];
                //判斷參數是否為空
                if (propertyDBName == "" || propertyTableName == "" || propertyColumName == "" || propertyDeleteDays == "")
                {
                    Logging.Log("------------------------------------------------------------", StrJobId, LogState.Info);
                    Logging.Log("格式錯誤,內容不得為空!", StrJobId, LogState.Error);
                    strSucc = "F";
                    StrMailMsg += "格式錯誤,內容不得為空! ";
                    continue;
                }
                //判斷輸入天數是否為數字
                if (!propertyDeleteDays.All(char.IsDigit))
                {
                    Logging.Log("------------------------------------------------------------", StrJobId, LogState.Info);
                    Logging.Log("格式錯誤,刪除天數必須為數字!", StrJobId, LogState.Error);
                    strSucc = "F";
                    StrMailMsg += "格式錯誤,刪除天數必須為數字! ";
                    continue;
                }

                Logging.Log("------------------------------------------------------------", StrJobId, LogState.Info);
                #region 紀錄預計刪除日期Log
                sql = new SqlCommand
                {
                    CommandType = CommandType.Text,
                    CommandText = "SELECT CONVERT(VARCHAR, DATEADD(day,-" + propertyDeleteDays + ",GETDATE()), 111)"
                };
                var dh = new DataHelper("Connection_CSIP");
                DataSet ds = dh.ExecuteDataSet(sql);
                if (ds != null && ds.Tables[0] != null && ds.Tables[0].Rows.Count > 0)
                {
                    Logging.Log("【預計刪除 " + ds.Tables[0].Rows[0][0].ToString() + " 前資料】", StrJobId, LogState.Info);
                }
                #endregion 

                Logging.Log("計算預計刪除 " + propertyTableName + " Table筆數", StrJobId, LogState.Info);
                
                //* 計算預計刪除筆數
                sql = new SqlCommand
                {
                    CommandType = CommandType.Text,
                    CommandText = "SELECT COUNT(*) FROM " + propertyTableName + " WHERE CONVERT(VARCHAR, " + propertyColumName + ", 111) <= DATEADD( day,- " + propertyDeleteDays + ", GETDATE( ) )"
                };

                //計算筆數
                TotalNum(sql, propertyDBName, ref totalNum, ref runNum);
                Logging.Log("預計刪除 " + propertyTableName + " Table " + totalNum + " 筆", StrJobId, LogState.Info);

                Logging.Log("刪除 " + propertyTableName + " Table 開始", StrJobId, LogState.Info);
                //實際刪除筆數計算
                totalDeleteCount = 0;
                int i = 0;
                do
                {
                    //* 刪除語法
                    sql = new SqlCommand
                    {
                        CommandType = CommandType.Text,
                        CommandText = "WITH DEL AS(SELECT TOP " + BatchNum +
                                  " * FROM " + propertyTableName + " WHERE CONVERT(VARCHAR, " + propertyColumName + " ,111)<= DATEADD( day,- " + propertyDeleteDays + ", GETDATE( ) ))" +
                                  "DELETE DEL WHERE 1=1"
                    };
                    int count = 0;
                    Delete(sql, propertyDBName, ref count);
                    totalDeleteCount += count;
                    i++;
                } while (i < runNum-1);
                Logging.Log("刪除 " + propertyTableName + " Table 結束", StrJobId, LogState.Info);

                Logging.Log("使用連線字串為 " + propertyDBName, StrJobId, LogState.Info);
                tempMsg =  propertyTableName + " 總共刪除：" + totalDeleteCount + " 筆";
                Logging.Log(tempMsg, StrJobId, LogState.Info);
                StrMailMsg += tempMsg + "<br>";
            }
            #endregion
            ConsumeTimeEnd(sw);

            #region 記錄結束時間

            EndTime = DateTime.Now;

            #endregion

            #region 執行結果寫入日誌

            //*批次完成記錄LOG信息
            //*插入L_BATCH_LOG資料庫
            BRL_BATCH_LOG.Delete(StrFunctionKey, StrJobId, "R");
            BRL_BATCH_LOG.Insert(StrFunctionKey, StrJobId, StartTime, strSucc, StrMailMsg);

            #endregion
        }
        catch (Exception exp)
        {
            //*批次完成記錄LOG信息
            //*插入L_BATCH_LOG資料庫
            BRL_BATCH_LOG.Delete(StrFunctionKey, StrJobId, "R");
            BRL_BATCH_LOG.Insert(StrFunctionKey, StrJobId, StartTime, "F", exp.Message.ToString());
            Logging.Log(exp.Message, StrJobId, LogState.Error);


            if (!string.IsNullOrWhiteSpace(mailTo))
            {
                var strTo = new[] { mailTo.Trim() };
                var strBody = StrMailMsg + ",Exception:" + exp;
                JobHelper.SendMail(_strFrom, strTo, strSubject + "執行失敗！", strBody);
            }
            else
            {
                JobHelper.SaveLog("寄信失敗，無 Email通知人資料。", LogState.Error);
            }
        }
        finally
        {
            //20220714_Ares_Jack_寄信訊息調整
            SqlCommand sqlcmd = new SqlCommand();
            sqlcmd.CommandType = CommandType.Text;
            sqlcmd.CommandText = "SELECT * FROM M_FUNCTION_KEY WHERE FUNCTION_KEY = '03'";
            DataSet ds = BRM_FUNCTION_KEY.SearchOnDataSet(sqlcmd, "Connection_CSIP");

            string FunctionKeyListName = ds == null ? "卡人卡片維護系統" : ds.Tables[0].Rows[0]["FUNCTION_KEY_NAME"].ToString();

            if (!string.IsNullOrWhiteSpace(mailTo))
            {
                var strTo = mailTo.Trim().Split(';');
                string strBody = FunctionKeyListName + "<br>" + StrMailMsg;

                if (JobHelper.SendMail(_strFrom, strTo, string.Format("【{0}】 {1} {2}", FunctionKeyListName, DateTime.Now.ToString("yyyy/MM/dd"), strSubject + (strSucc == "S" ? "執行成功！" : "執行失敗！")), strBody))
                {
                    Logging.Log("寄信成功。", StrJobId, LogState.Info);
                }
                else
                {
                    Logging.Log("寄信失敗。", StrJobId, LogState.Error);
                }
            }
            else
            {
                Logging.Log("寄信失敗，無 Email通知人資料。。", StrJobId, LogState.Error);
            }

            Logging.Log("HOUSE_KEEPING 結束！", StrJobId, LogState.Info);
        }
    }

    #endregion


    /// <summary>
    /// 專案代號:20200031-CSIP EOS
    /// 功能說明:Delete by SQL
    /// 作    者:Ares Jack
    /// 創建時間:2021/08/03
    /// </summary>
    /// <param name="command">DbCommand</param>
    /// <param name="strConnectionName">連接字串名</param>
    /// <param name="count">筆數</param>
    private bool Delete(DbCommand command, string strConnectionName, ref int count)
    {
        var dh = new DataHelper(strConnectionName);
        try
        {
            Logging.Log("執行SQL:" + command.CommandText, StrJobId, LogState.Info);
            command.CommandTimeout = int.Parse(UtilHelper.GetAppSettings("SqlCmdTimeoutMax"));
            count = dh.ExecuteNonQuery(command);
            return true;
        }
        catch (Exception exp)
        {
            Logging.Log(exp, StrJobId);
            return false;
        }
    }


    /// <summary>
    /// 專案代號:20200031-CSIP EOS
    /// 功能說明:統計總筆數與批次執行次數
    /// 作    者:Ares Jack
    /// 創建時間:2021/08/03
    /// 修改紀錄:Ares_Jack-新增TimeOut時間
    /// </summary>
    /// <param name="command">DbCommand</param>
    /// <param name="strConnectionName">連接字串名</param>
    /// <param name="totalNum">總筆數</param>
    /// <param name="runNum">執行次數</param>
    private void TotalNum(DbCommand command, string strConnectionName, ref int totalNum, ref decimal runNum)
    {
        var dh = new DataHelper(strConnectionName);
        try
        {
            Logging.Log("執行SQL:" + command.CommandText, StrJobId, LogState.Info);
            command.CommandTimeout = int.Parse(UtilHelper.GetAppSettings("SqlCmdTimeoutMax"));
            totalNum = 0;
            runNum = 0;
            totalNum = (Int32) dh.ExecuteScalar(command);
            runNum = (decimal) Math.Ceiling((double) totalNum / BatchNum);
        }
        catch (Exception exp)
        {
            Logging.Log(exp, StrJobId);
        }
    }


    /// <summary>
    /// 專案代號:20200031-CSIP EOS
    /// 功能說明:開始計時器並記錄LOG
    /// 作    者:Ares Jack
    /// 創建時間:2021/08/03
    /// 修改紀錄:
    /// </summary>
    /// <param name="sw"></param>
    private void ConsumeTimeStart(Stopwatch sw)
    {
        Logging.Log("開始時間:" + DateTime.Now.ToString(), StrJobId, LogState.Info);
        sw.Start();
    }

    /// <summary>
    /// 專案代號:20200031-CSIP EOS
    /// 功能說明:結束計時器並記錄LOG
    /// 作    者:Ares Jack
    /// 創建時間:2021/08/03
    /// 修改紀錄:
    /// </summary>
    /// <param name="sw"></param>
    private void ConsumeTimeEnd(Stopwatch sw)
    {
        Logging.Log("------------------------------------------------------------", StrJobId, LogState.Info);
        Logging.Log("結束時間:" + DateTime.Now.ToString(), StrJobId, LogState.Info);
        Logging.Log("總花費時間:" + sw.ElapsedMilliseconds + "毫秒", StrJobId, LogState.Info);
        Logging.Log("============================================================", StrJobId, LogState.Info);
        sw.Stop();
    }
}