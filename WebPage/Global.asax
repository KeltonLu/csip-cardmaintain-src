<%@ Application Language="C#" %>
<%@ Import Namespace="Framework.Common.Utility" %>
<script runat="server">
    void Application_Start(object sender, EventArgs e) 
    {
        #region 系統Log啟動
        log4net.Config.XmlConfigurator.Configure(new System.IO.FileInfo(AppDomain.CurrentDomain.BaseDirectory + @"\Log.config"));
        #endregion

        #region 系統框架啟動
        // 應用程式啟動時執行的程式碼
        if (!Framework.Common.PopedomManager.MainPopedomManager.Inited)
        {
            Framework.Common.PopedomManager.MainPopedomManager.Init();
        }
        #endregion
        #region Job啟動      
        string strMsgID = "";
        try
        {
            if (ConfigurationManager.AppSettings["BatchStatus"].ToString() == "Y")
            {

                Framework.Common.Logging.Logging.Log("Application and Jobs start");

                System.Data.DataTable dtblAutoJob = CSIPCommonModel.BusinessRules.BRM_AUTOJOB.GetJobData(UtilHelper.GetAppSettings("FunctionKey").ToString(), ref strMsgID);
                if (dtblAutoJob == null)
                {
                    //* 取得Job資料失敗
                    Framework.Common.Logging.Logging.Log(strMsgID, Framework.Common.Logging.LogState.Error, Framework.Common.Logging.LogLayer.None);
                    return;
                }

                if (dtblAutoJob.Rows.Count > 0)
                {
                    //* 取得排程信息
                    Quartz.ISchedulerFactory sf = new Quartz.Impl.StdSchedulerFactory();
                    Quartz.IScheduler sched = sf.GetScheduler();

                    //* 加解密操作類
                    Framework.Common.Cryptography.DESEncrypt des = new Framework.Common.Cryptography.DESEncrypt();
                    des.EncryptKey = UtilHelper.GetAppSettings("EncryptKey").ToString();


                    for (int i = 0; i < dtblAutoJob.Rows.Count; i++)
                    {
                        try
                        {
                            if (dtblAutoJob.Rows[i]["STATUS"].ToString() == "1")
                            {
                                string strJobName = "job_" + dtblAutoJob.Rows[i]["JOB_ID"].ToString();
                                string strJobGroup = "CSIPGroup";
                                string strTriggerName = "trigger_" + dtblAutoJob.Rows[i]["JOB_ID"].ToString();
                                string strTriggerGroup = "CSIPGroup";

                                //* 新增一個Job
                                Type type = System.Web.Compilation.BuildManager.GetType(dtblAutoJob.Rows[i]["EXEC_PROG"].ToString(), false);
                                Quartz.JobDetail job = new Quartz.JobDetail(strJobName, strJobGroup, type);

                                //* 新增Job參數
                                Quartz.JobDataMap jobDataMap = new Quartz.JobDataMap();
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
                                Quartz.CronTrigger trigger = new Quartz.CronTrigger(strTriggerName, strTriggerGroup, strJobName, strJobGroup, cronExpr);
                                //* 將Job和觸發器都加入到當前系統排程中去
                                sched.AddJob(job, true);
                                DateTime ft = sched.ScheduleJob(trigger);

                                Framework.Common.Logging.Logging.Log(strJobName + " Start..");
                            }
                        }
                        catch (Exception ex)
                        {
                            //* 任意一個Job新增失敗.不影響後面Job添加
                            Framework.Common.Logging.Logging.Log(ex);
                        }
                    }
                    sched.Start();
                }
            }
        }
        catch (Exception ex)
        {
            Framework.Common.Logging.Logging.Log(ex.ToString());
        }
        #endregion
    }
    
    void Application_End(object sender, EventArgs e) 
    {
        //  應用程式關閉時執行的程式碼
        Quartz.ISchedulerFactory sf = new Quartz.Impl.StdSchedulerFactory();
        Quartz.IScheduler sched = sf.GetScheduler();
        if (sched != null)
        {
            sched.Shutdown(true);
			Framework.Common.Logging.Logging.Log("Application and Jobs End");
            
            //* 訪問系統頁面，以重新啟動JOB排程
            System.Net.HttpWebRequest request = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(UtilHelper.GetAppSettings("LOGOUT").ToString());
            System.Net.HttpWebResponse myResponse = (System.Net.HttpWebResponse)request.GetResponse();
        }
    }
        
    void Application_Error(object sender, EventArgs e) 
    { 
        // 發生未處理錯誤時執行的程式碼

    }

    void Session_Start(object sender, EventArgs e) 
    {
        // 啟動新工作階段時執行的程式碼

    }

    void Session_End(object sender, EventArgs e) 
    {
        // 工作階段結束時執行的程式碼。 
        // 注意: 只有在 Web.config 檔將 sessionstate 模式設定為 InProc 時，
        // 才會引發 Session_End 事件。如果將工作階段模式設定為 StateServer 
        // 或 SQLServer，就不會引發這個事件。

    }
       
</script>
