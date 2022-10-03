//******************************************************************
//* 作    者：
//* 功能說明：
//* 創建日期：
//* 修改紀錄：
//* <author>            <time>            <TaskID>                <desc>
//* Ares Luke          2020/11/19         20200031-CSIP EOS       調整取web.config加解密參數
//* Ares JaJa          2021/02/04         20200031-CSIP EOS       補充LOG
//*******************************************************************

using System;
using System.Data;
using Quartz;
using CSIPCommonModel.BusinessRules;
using Framework.Common.Logging;
using Framework.Common.Cryptography;
using System.Web.Compilation;
using Quartz.Impl;
using Framework.Common.Utility;

/// <summary>
/// jobRescheduleJob 的摘要描述
/// </summary>
public class jobRescheduleJob : IJob
{
    protected JobHelper JobHelper = new JobHelper();
    public jobRescheduleJob()
    {
        //
        // TODO: 在此加入建構函式的程式碼
        //
    }


    #region IJob 成員

    public void Execute(JobExecutionContext context)
    {

        string strFuncKey = UtilHelper.GetAppSettings("FunctionKey").ToString();
        string strJobID = "";
        string strJobTitle;
        string strLdapID;
        string strLdapPWD;
        string strRacfID;
        string strRacfPwd;
        string strMailList;
        DateTime dtStart = DateTime.Now;

        try
        {
            strJobID = context.JobDetail.JobDataMap.GetString("jobid").Trim();
            strJobTitle = context.JobDetail.JobDataMap.GetString("title").Trim();
            strLdapID = context.JobDetail.JobDataMap.GetString("userId").Trim();
            strLdapPWD = context.JobDetail.JobDataMap.GetString("passWord").Trim();
            strRacfID = context.JobDetail.JobDataMap.GetString("racfId").Trim();
            strRacfPwd = context.JobDetail.JobDataMap.GetString("racfPassWord").Trim();
            strMailList = context.JobDetail.JobDataMap.GetString("mail").Trim();
           JobHelper.strJobID = strJobID;
            
            JobHelper.Write(strJobID, "****************************** " + strJobID + " START ****************************** ", LogState.Info);

            string strMsgID = "";
            int iSuccess = 0;
            int iFail = 0;
            JobHelper.Write(strJobID, "****************************** JobRefresh 查詢資料檔L_BATCH_LOG，查看是否上次作業還未停止 ******************************", LogState.Info);
            
            //*查詢資料檔L_BATCH_LOG，查看是否上次作業還未停止
            DataTable dtInfo = BRL_BATCH_LOG.GetRunningDate(strFuncKey, strJobID, "R", ref strMsgID);
            if (dtInfo == null)
            {
                JobHelper.Write(strJobID, "*********** JobRefresh dtInfo == null ***********", LogState.Info);
                return;
            }
            if (dtInfo.Rows.Count > 0)
            {
                JobHelper.Write(strJobID, "*********** JobRefresh dtInfo.Rows.Count > 0 ***********", LogState.Info);
                return;
            }
            //*開始批次作業
            if (!BRL_BATCH_LOG.InsertRunning(strFuncKey, strJobID, dtStart, "R", ""))
            {
                JobHelper.Write(strJobID, "*********** JobRefresh InsertRunning R ***********", LogState.Info);
                return;
            }
            
            //* 取得排程信息
            ISchedulerFactory sf = new StdSchedulerFactory();
            IScheduler sched = sf.GetScheduler();

            //* 取得資料庫中設定值
            DataTable dtblAutoJob = BRM_AUTOJOB.GetJobData(strFuncKey, ref strMsgID);
            if (dtblAutoJob == null) { return; }

            if (dtblAutoJob.Rows.Count > 0)
            {
                
                for (int i = 0; i < dtblAutoJob.Rows.Count; i++)
                {

                    string strJobName = "";         //* JobName
                    string strJobGroup = "";        //* JobGroup
                    string strTriggerName = "" ;    //* TriggerName
                    string strTriggerGroup = "";    //* TriggerGroup

                    try
                    {
                        strJobName = "job_" + dtblAutoJob.Rows[i]["JOB_ID"].ToString();                   //* JobName
                        strJobGroup = "CSIPGroup";                                               //* JobGroup
                        strTriggerName = "trigger_" + dtblAutoJob.Rows[i]["JOB_ID"].ToString();  //* TriggerName
                        strTriggerGroup = "CSIPGroup";                                           //* TriggerGroup

                        //* 目前排程里是否有這個Job
                        bool bAllreadyHadJob = false;
                        if (sched.GetJobDetail(strJobName, strJobGroup) == null) { bAllreadyHadJob = false; } else { bAllreadyHadJob = true; }

                        
                        if (dtblAutoJob.Rows[i]["STATUS"].ToString() == "1")
                        {
                            //* 如果資料庫里該Job爲啟動狀態

                            //* 加解密操作類
                            DESEncrypt des = new DESEncrypt();
                            des.EncryptKey = UtilHelper.GetAppSettings("EncryptKey").ToString();

                            //* 新增一個Job
                            Type type = BuildManager.GetType(dtblAutoJob.Rows[i]["EXEC_PROG"].ToString(), false);
                            JobDetail job = new JobDetail(strJobName, strJobGroup, type);

                            //* 新增Job參數
                            JobDataMap jobDataMap = new JobDataMap();
                            jobDataMap.Put("userId", dtblAutoJob.Rows[i]["RUN_USER_LDAPID"].ToString());
                            jobDataMap.Put("passWord", des.DecryptString(dtblAutoJob.Rows[i]["RUN_USER_LDAPPWD"].ToString()));
                            jobDataMap.Put("racfId", dtblAutoJob.Rows[i]["RUN_USER_RACFID"].ToString());
                            jobDataMap.Put("racfPassWord", des.DecryptString(dtblAutoJob.Rows[i]["RUN_USER_RACFPWD"].ToString()));
                            jobDataMap.Put("mail", dtblAutoJob.Rows[i]["MAIL_TO"].ToString());
                            jobDataMap.Put("title", dtblAutoJob.Rows[i]["DESCRIPTION"].ToString());
                            jobDataMap.Put("jobid", dtblAutoJob.Rows[i]["JOB_ID"].ToString());
                            job.JobDataMap = jobDataMap;
                            //* 時間表達式
                            string cronExpr = dtblAutoJob.Rows[i]["RUN_SECONDS"].ToString().Trim() + " " +
                                            dtblAutoJob.Rows[i]["RUN_MINUTES"].ToString().Trim() + " " +
                                            dtblAutoJob.Rows[i]["RUN_HOURS"].ToString().Trim() + " " +
                                            dtblAutoJob.Rows[i]["RUN_DAY_OF_MONTH"].ToString().Trim() + " " +
                                            dtblAutoJob.Rows[i]["RUN_MONTH"].ToString().Trim() + " " +
                                            dtblAutoJob.Rows[i]["RUN_DAY_OF_WEEK"].ToString().Trim();

                            //* 新增一個Trigger觸發器
                            CronTrigger trigger = new CronTrigger(strTriggerName, strTriggerGroup, strJobName, strJobGroup, cronExpr);
                            trigger.StartTimeUtc = DateTime.UtcNow.AddSeconds(1);
                            if (bAllreadyHadJob)
                            {
                                
                                //* 已經有這個Job了                          
                                //* 爲了更新排程只能覆蓋新增了
                                sched.AddJob(job, true);
                                //* 更新Trigger
                                sched.RescheduleJob(strTriggerName, strTriggerGroup, trigger);
                            }
                            else
                            {
                                //* 沒有這個Job新增Job
                                //* 將Job和觸發器都加入到當前系統排程中去
                                sched.AddJob(job, true);
                                sched.ScheduleJob(trigger);
                            }
                        }
                        else
                        {
                            //* QuartZ排程里有,DB里已停止
                            if (bAllreadyHadJob)
                            {
                                //* 如果目前排程存在該Job則停用.
                                sched.UnscheduleJob(strTriggerName, strTriggerGroup);
                                sched.DeleteJob(strJobName, strJobGroup);
                            }
                        }

                        iSuccess = iSuccess + 1;

                    }
                    catch (Exception ex)
                    {
                        //* 任意一個Job新增失敗.不影響後面Job添加
                        Logging.Log(strJobName + " ERROR", strJobID);
                        Logging.Log(ex, strJobID);
                        iFail = iFail + 1;
                    }
                }
                JobHelper.Write(strJobID, "****************************** " + dtblAutoJob.Rows.Count + " Jobs Updated ******************************", LogState.Info);
            }


            BRL_BATCH_LOG.Delete(strFuncKey, strJobID, "R");
            if (iFail == 0)
            {
                BRL_BATCH_LOG.Insert(strFuncKey, strJobID, dtStart, "S", "成功:" + iSuccess.ToString() + ";失敗:" + iFail.ToString());
            }
            else if (iSuccess == 0)
            {
                BRL_BATCH_LOG.Insert(strFuncKey, strJobID, dtStart, "F", "成功:" + iSuccess.ToString() + ";失敗:" + iFail.ToString());
            }
            else
            {
                BRL_BATCH_LOG.Insert(strFuncKey, strJobID, dtStart, "P", "成功:" + iSuccess.ToString() + ";失敗:" + iFail.ToString());
            }
        }
        catch (Exception ex)
        {
            BRL_BATCH_LOG.Delete(strFuncKey, strJobID, "R");
            BRL_BATCH_LOG.Insert(strFuncKey, strJobID, dtStart, "F", ex.Message);
            JobHelper.SaveLog(ex);
        }
    }

    #endregion
}
