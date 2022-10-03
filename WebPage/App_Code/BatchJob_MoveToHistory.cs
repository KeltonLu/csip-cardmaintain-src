//******************************************************************
//*  作    者：Ares Stanley
//*  創建日期：2022/01/10
//*  功能說明：Stored Procedure轉批次，將一年半之前的資料搬到 P4 卡人卡片歷史資料檔中

//*<author>            <time>            <TaskID>                <desc>
//*******************************************************************

using System;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using Quartz;
using CSIPCardMaintain.BusinessRules;
using CSIPCommonModel.BusinessRules;
using CSIPCommonModel.EntityLayer;
using Framework.Common.Logging;
using Framework.Common.Utility;
using Framework.Data;

/// <summary>
/// BatchJob_MoveToHistory 的摘要描述
/// </summary>
public class BatchJob_MoveToHistory : Quartz.IJob
{
    protected string FunctionKey = UtilHelper.GetAppSettings("FunctionKey").ToString();
    protected DateTime StartTime = DateTime.Now; //Job啟動時間
    protected DateTime EndTime;
    protected JobHelper JobHelper = new JobHelper();
    public void Execute(JobExecutionContext context)
    {
        string jobID = context.JobDetail.JobDataMap["jobid"].ToString();
        JobHelper.strJobID = jobID;
        string msgID = string.Empty;

        JobDataMap jobDataMap = context.JobDetail.JobDataMap;
        int dataCount = 0;

        try
        {
            JobHelper.SaveLog("*********** " + jobID + " 將一年半之前的資料搬到 P4 卡人卡片歷史資料檔 批次 START ************** ", LogState.Info);


            bool isContinue = CheckJobIsContinue(jobID, ref msgID);

            if (isContinue)
            {
                #region 判斷是否手動設置參數啟動排程
                JobHelper.SaveLog("判斷是否手動輸入參數 啟動排程：開始！", LogState.Info);

                if (context.JobDetail.JobDataMap["param"] != null)
                {
                    JobHelper.SaveLog("手動輸入參數啟動排程：是！", LogState.Info);
                    JobHelper.SaveLog("檢核輸入參數：開始！", LogState.Info);

                    string strParam = context.JobDetail.JobDataMap["param"].ToString();

                    if (strParam.Length == 10)
                    {
                        DateTime tempDt;
                        if (DateTime.TryParse(strParam, out tempDt))
                        {
                            JobHelper.SaveLog("檢核參數：成功！ 參數：" + strParam, LogState.Info);

                            if (BRM_FileInfo.UpdateParam(jobID, tempDt.ToString("yyyyMMdd")))
                            {
                                JobHelper.SaveLog("更新參數至FileInfo：成功！ 參數：" + tempDt.ToString("yyyyMMdd"), LogState.Info);
                            }
                            else
                            {
                                JobHelper.SaveLog("更新參數至FileInfo：失敗！ 參數：" + tempDt.ToString("yyyyMMdd"), LogState.Error);
                                return;
                            }
                        }
                        else
                        {
                            JobHelper.SaveLog("檢核參數：異常！ 參數：" + strParam, LogState.Error);
                            return;
                        }
                    }
                    else
                    {
                        JobHelper.SaveLog("檢核參數：異常！ 參數：" + strParam, LogState.Error);
                        return;
                    }

                    JobHelper.SaveLog("檢核輸入參數：結束！", LogState.Info);
                }
                else
                {
                    JobHelper.SaveLog("手動輸入參數啟動排程：否！", LogState.Info);
                }

                JobHelper.SaveLog("判斷是否手動輸入參數 啟動排程：結束！", LogState.Info);

                #endregion

            }
            else
            {
                return;
            }

            #region 取得資料筆數
            string now = DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-18).ToString("yyyyMMdd");

            JobHelper.SaveLog("查詢一年半之前資料筆數 開始", LogState.Info);

            const string sqlText = @"
SELECT
	TYPE,
	CUST_ID,
	CARD_TYPE,
	FLD_NAME,
	BEFOR_UPD,
	AFTER_UPD,
	LST_LIMIT,
	CUR_LIMIT,
	MAINT_D,
	MAINT_T,
	USER_ID,
	TER_ID,
	EXE_Name 
FROM
	CPMAST 
WHERE
	MAINT_D < @now ";

            JobHelper.SaveLog(string.Format("紀錄查詢SQL：{0}\r\n", sqlText.Replace("@now", string.Format("'{0}'", now)), LogState.Info));

            SqlCommand sqlcmd = new SqlCommand();
            sqlcmd.CommandText = sqlText;
            sqlcmd.Parameters.Add(new SqlParameter("@now", now));
            DataSet ds = BRCPMAST.SearchOnDataSet(sqlcmd);

            if (ds == null || ds.Tables.Count <= 0)
            {
                JobHelper.SaveLog("查詢時發生錯誤，請確認 JobLog 或 DefaultLog", LogState.Info);
                InsertBatchLog(jobID, "F", "查詢時發生錯誤，請確認 JobLog 或 DefaultLog");
                return;
            }

            DataTable dt = ds.Tables[0];

            if (dt == null || dt.Rows.Count <= 0)
            {
                JobHelper.SaveLog("無一年半前資料", LogState.Info);
                InsertBatchLog(jobID, "S", "無一年半前資料");
                return;
            }
            else
            {
                dataCount = dt.Rows.Count;
            }

            JobHelper.SaveLog(string.Format("一年半之前資料筆數共 {0} 筆", dataCount), LogState.Info);

            JobHelper.SaveLog("查詢一年半之前資料筆數 結束", LogState.Info);
            #endregion

            #region 執行SP
            JobHelper.SaveLog("執行 sp_MoveToHistory 開始", LogState.Info);

            DataHelper dh = new DataHelper();
            SqlCommand sqlcmd2 = new SqlCommand();
            bool executeSpStatus = false;
            sqlcmd2.CommandText = "sp_MoveToHistory";
            sqlcmd2.CommandTimeout = int.Parse(UtilHelper.GetAppSettings("SqlCmdTimeoutMax"));//20220530_Ares_Jack_新增TimeOut時間 70分鐘
            int affectRows = dh.ExecuteNonQuery(sqlcmd2);
            executeSpStatus = affectRows > 0;

            if (!executeSpStatus)
            {
                JobHelper.SaveLog("sp_MoveToHistory 執行失敗，請確認", LogState.Info);
                InsertBatchLog(jobID, "F", "sp_MoveToHistory 執行失敗，請確認");
                return;
            }
            else
            {
                JobHelper.SaveLog("sp_MoveToHistory 執行成功", LogState.Info);
            }
            JobHelper.SaveLog("執行 sp_MoveToHistory 結束", LogState.Info);

            InsertBatchLog(jobID, "S", string.Format("異動筆數共 {0} 筆", dataCount));
            #endregion
        }
        catch (Exception ex)
        {
            Logging.Log(ex.Message);
            JobHelper.SaveLog("BatchJob_MoveToHistory 發生例外錯誤：" + ex.Message);
            InsertBatchLog(jobID, "F", "排程 BatchJob_MoveToHistory 發生例外錯誤");
            return;
        }
        finally
        {
            JobHelper.SaveLog("*********** " + jobID + " 將一年半之前的資料搬到 P4 卡人卡片歷史資料檔 批次 END ************** ", LogState.Info);
        }
    }

    #region Method

    /// <summary>
    /// 寫入Batch_Log
    /// </summary>
    /// <param name="jobID">jobID</param>
    /// <param name="status">執行狀態</param>
    /// <param name="message">執行訊息</param>
    private void InsertBatchLog(string jobID, string status, string message)
    {
        try
        {
            StringBuilder sbMessage = new StringBuilder();

            if (message.Trim() != "" && status != "S")
            {
                sbMessage.Append("失敗訊息：" + message);//*失敗訊息
            }

            if (message.Trim() != "" && status == "S")
            {
                sbMessage.Append(message);//*成功訊息
            }

            BRL_BATCH_LOG.Delete("03", jobID, "R");
            BRL_BATCH_LOG.Insert("03", jobID, this.StartTime, status, sbMessage.ToString());
        }
        catch (Exception ex)
        {
            Logging.Log(ex.Message);
            JobHelper.SaveLog("寫入 Batch_Log 發生例外錯誤：" + ex.Message);
        }
    }

    /// <summary>
    /// 判斷Job工作狀態(0:停止 1:運行)
    /// </summary>
    /// <param name="jobID"></param>
    /// <param name="msgID"></param>
    /// <returns></returns>
    private bool CheckJobIsContinue(string jobID, ref string msgID)
    {
        bool result = true;
        if (JobHelper.SerchJobStatus(jobID).Equals("") || JobHelper.SerchJobStatus(jobID).Equals("0"))
        {
            // Job停止
            JobHelper.SaveLog("[FAIL] Job工作狀態為：停止！");
            result = false;
        }

        // 檢測Job是否在執行中
        try
        {
            DataTable dtInfo = BRL_BATCH_LOG.GetRunningDate(FunctionKey, jobID, "R", ref msgID);
            if (dtInfo == null || dtInfo.Rows.Count > 0)
            {
                JobHelper.SaveLog("JOB 工作狀態為：正在執行！", LogState.Info);
                // 返回不執行
                result = false;
            }
            else
            {
                // 記錄Job執行資訊
                BRL_BATCH_LOG.InsertRunning(FunctionKey, jobID, StartTime, "R", "");
            }
        }
        catch (Exception ex)
        {
            result = false;
            Logging.Log(ex.Message);
            JobHelper.SaveLog("判斷Job工作狀態發生例外錯誤：" + ex.Message);
        }

        return result;
    }
    #endregion
}