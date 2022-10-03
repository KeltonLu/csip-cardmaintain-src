//******************************************************************
//*  作    者：chaoma(Wilson)
//*  功能說明：立即運行JOB
//*  創建日期：2009/08/03
//*  修改記錄：
//*<author>            <time>            <TaskID>                <desc>
//* chaoma           2010/07/29     20100009          共用模組增加手動執行 JOB功能
//* Ares_Rick        2021/12/24     20210058          CSIP作業服務平台現代化II-新增jobRun機制
//*******************************************************************
using System;
using System.Data;
using Framework.Common.Utility;
using Framework.Common.Message;
using Quartz;
using System.Threading;
using Framework.Common.Logging;

public partial class Page_JobRun : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {


        string strMsgID = string.Empty;

        //* jobID
        string strjobID = RedirectHelper.GetDecryptString(this.Page, "JobID");
        string strExecprog = RedirectHelper.GetDecryptString(this.Page, "Execprog");
        string strParam = RedirectHelper.GetDecryptString(this.Page, "Param");

        if (string.IsNullOrEmpty(strjobID))
        {
            MessageHelper.ShowMessage(this.Page, "00_00000000_037");
        }

        System.Data.DataTable dtblAutoJob = CSIPCommonModel.BusinessRules.BRM_AUTOJOB.GetJobData(Framework.Common.Utility.UtilHelper.GetAppSettings("FunctionKey"), ref strMsgID);
        if (dtblAutoJob == null)
        {
            //* 取得Job資料失敗
            //Framework.Common.Logging.Logging.Log(strMsgID, Framework.Common.Logging.LogState.Error, Framework.Common.Logging.LogLayer.None);
            return;
        }

        DataRow[] info = dtblAutoJob.Select("JOB_ID='" + strjobID + "'");

        try
        {
            if (info.Length == 1)
            {

                //* 取得排程信息
                Quartz.ISchedulerFactory sfr = new Quartz.Impl.StdSchedulerFactory();
                Quartz.IScheduler schedr = sfr.GetScheduler();

                //* 加解密操作類
                Framework.Common.Cryptography.DESEncrypt des = new Framework.Common.Cryptography.DESEncrypt();
                des.EncryptKey = Framework.Common.Utility.UtilHelper.GetAppSettings("EncryptKey");

                string strJobName = "reRunJob_" + strjobID;
                string strJobGroup = "CSIPGroup";
                string strTriggerName = "reRunTrigger_" + strjobID;
                string strTriggerGroup = "CSIPGroup";

                //* 新增一個Job
                //info["EXEC_PROG"].ToString()
                Type type = System.Web.Compilation.BuildManager.GetType(info[0]["EXEC_PROG"].ToString(), false);
                Quartz.JobDetail job = new Quartz.JobDetail(strJobName, strJobGroup, type);

                //* 新增Job參數
                Quartz.JobDataMap jobDataMap = new Quartz.JobDataMap();
                jobDataMap.Put("userId", info[0]["RUN_USER_LDAPID"].ToString());
                jobDataMap.Put("passWord", des.DecryptString(info[0]["RUN_USER_LDAPPWD"].ToString()));
                jobDataMap.Put("racfId", info[0]["RUN_USER_RACFID"].ToString());
                jobDataMap.Put("racfPassWord", des.DecryptString(info[0]["RUN_USER_RACFPWD"].ToString()));
                jobDataMap.Put("mail", info[0]["MAIL_TO"].ToString());
                jobDataMap.Put("title", info[0]["DESCRIPTION"].ToString());
                jobDataMap.Put("jobid", info[0]["JOB_ID"].ToString());
                //* 參數設定
                jobDataMap.Put("param", strParam);

                job.JobDataMap = jobDataMap;
                //* 將Job和觸發器都加入到當前系統排程中去
                schedr.AddJob(job, true);

                DateTime runTimer = TriggerUtils.GetEvenMinuteDate(DateTime.UtcNow);
                runTimer = DateTime.Parse("2010-04-22 18:06:00").ToUniversalTime();
                SimpleTrigger triggerr = new SimpleTrigger(strTriggerName, strTriggerGroup, runTimer);
                triggerr.JobName = strJobName;
                triggerr.JobGroup = strJobGroup;

                schedr.ScheduleJob(triggerr);
                schedr.Start();

                // 延遲刪除暫存的排程
                Thread.Sleep(1500);
                schedr.DeleteJob(strJobName, strJobGroup);
            }
        }
        catch (Exception exp)
        {
            Logging.Log(exp);
        }

        //* jobID
        //string strjobID = RedirectHelper.GetDecryptString(this.Page, "JobID");

        //if (string.IsNullOrEmpty(strjobID))
        //{
            //MessageHelper.ShowMessage(this.Page, "00_00000000_037");
        //}
            //* 新建排程
            //ISchedulerFactory sfr = new StdSchedulerFactory();
            //IScheduler schedr = sfr.GetScheduler();


            //DateTime runTimer = TriggerUtils.GetEvenMinuteDate(DateTime.UtcNow);
            //runTimer = DateTime.Parse("2010-04-22 18:06:00").ToUniversalTime();
            //SimpleTrigger triggerr = new SimpleTrigger("trigger1", "group1", runTimer);
            //triggerr.JobName =  "job_" + strjobID;
            //triggerr.JobGroup = "CSIPGroup";
            //schedr.ScheduleJob(triggerr);
            //schedr.Start();
        
    }
    
}
