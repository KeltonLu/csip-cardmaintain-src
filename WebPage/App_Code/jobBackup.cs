//******************************************************************
//*  作    者：Mars
//*  功能說明：清檔及備份 LOG,LOGXML,上傳檔案
//*  創建日期：2012/12/14
//*  修改記錄：
//*<author>            <time>            <TaskID>                <desc>
//* Ares Stanley    2022/02/14    20210058-CSIP作業服務平台現代化II    調整webconfig取參數方式
//*******************************************************************

using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using Framework.Data.OM.Collections;
using Framework.Data.OM;
using Framework.WebControls;
using Framework.Common.Utility;
using Framework.Common.Message;
using Framework.Common.JavaScript;
using Framework.Common.Logging;
using CSIPCommonModel.EntityLayer;
using Framework.Data.OM.Transaction;
using CSIPCommonModel.BusinessRules;
using Quartz;
using Quartz.Impl;
using Framework.Common.IO;
using System.Collections;
using System.IO;
using System.Collections.Generic;


/// <summary>
/// jobBackup 的摘要描述
/// </summary>
public class jobBackup : IJob
{
    public jobBackup()
    {
        //
        // TODO: 在此加入建構函式的程式碼
        //
    }

    string strFuncKey = UtilHelper.GetAppSettings("FunctionKey").ToString();
    string strSucc = "";
    DateTime dTimeStart;
    private string strJobID;
    private string strJobMsg;
    //private string strMail;
    //private string strJobTitle;

    //*需處理的來源資料夾
    private readonly List<string> SourceFolder = new List<string>(UtilHelper.GetAppSettings("SourceFolder").Split(','));
    //*需排除附檔名列表
    private readonly List<string> SkipExtension = new List<string>(UtilHelper.GetAppSettings("SkipExtension").Split(','));

    #region IJob 成員
    /// <summary>
    /// Job 調用入口
    /// </summary>
    /// <param name="context"></param>
    public void Execute(JobExecutionContext context)
    { 
        //string strLdapID;
        //string strLdapPWD;
        //string strRacfID;
        //string strRacfPwd;

        // 獲取JOB開始時間
        dTimeStart = DateTime.Now;

        strJobMsg = "";

        try
        {
            Log("*************** 備份及清檔作業開始 ***************");
            strJobID = context.JobDetail.JobDataMap.GetString("jobid").Trim();
            string strMsgID = "";
            //*查詢資料檔L_BATCH_LOG，查看是否上次作業還未停止
            DataTable dtInfo = BRL_BATCH_LOG.GetRunningDate(strFuncKey, strJobID, "R", ref strMsgID);
            if (dtInfo == null)
            {
                return;
            }
            if (dtInfo.Rows.Count > 0)
            {
                return;
            }
            //*開始批次作業
            if (!InsertNewBatch())
            {
                return;
            }

            //*來源檔案基底位置
            string BaseDir = AppDomain.CurrentDomain.BaseDirectory;

            //*來源資料夾位置
            string SourcePath = "";

            //*備份資料夾位置
            string BackupPath = "";

            //*原始檔案保存起始日(SourceKeepDay)
            int SourceKeepDay = Convert.ToInt32(UtilHelper.GetAppSettings("SourceKeepDay").ToString());

            //*備份檔案保存起始日(BackupKeepDay)
            int BackupKeepDay = Convert.ToInt32(UtilHelper.GetAppSettings("BackupKeepDay").ToString());

            //*備份路徑(以執行當天的日期命名)
            string BackupDir = UtilHelper.GetAppSettings("BackupPath").ToString() + dTimeStart.ToString("yyyyMMdd");

            //*執行備份動作，如果BackupInitial=true則把所有檔案都備份一次
            Log("***開始備份***");
            bool BackupALL = Convert.ToBoolean(UtilHelper.GetAppSettings("BackupALL"));
            foreach (string SF in SourceFolder)
            {
                SourcePath = BaseDir + SF;
                BackupPath = BackupDir + "\\" + SF;
                if (Directory.Exists(SourcePath))
                {
                    Log(SourcePath + "”開始掃描”");
                    FileBackup(SourcePath, BackupPath, SourceKeepDay, BackupALL);
                }
                else
                    Log(SourcePath + " 此路徑不存在！");
            }
            if (BackupALL)
                strJobMsg += "備份所有資料OK；";
            else
                strJobMsg += "備份" + SourceKeepDay.ToString() + "天前資料OK；";
            Log("***備份完成***");

            //*清除來源資料夾過期檔案
            Log("***開始清除來源資料夾過期檔案***");
            foreach (string SF in SourceFolder)
            {
                SourcePath = BaseDir + SF;
                if (Directory.Exists(SourcePath))
                {
                    Log(SourcePath + "”開始掃描”");
                    ClearFile(SourcePath, SourceKeepDay);
                }
                else
                    Log(SourcePath + " 此路徑不存在！");
            }
            strJobMsg += "清除" + SourceKeepDay.ToString() + "天前來源資料OK；";
            Log("***來源資料夾過期檔案清除完成***");

            //*清除備份資料夾過期檔案
            Log("***開始清除過期備份資料夾***");
            Log(BackupDir.Substring(0, BackupDir.LastIndexOf("\\")) + "”開始掃描”");
            ClearFolder(BackupDir.Substring(0, BackupDir.LastIndexOf("\\")), BackupKeepDay);
            strJobMsg += "清除" + BackupKeepDay.ToString() + "天前備份資料OK；";
            Log("***過期備份資料夾清除完成***");

            //*批次完成記錄LOG信息
            string strMsg = strJobID + "執行於:" + DateTime.Parse(context.FireTimeUtc.ToString()).AddHours(8).ToString();
            if (context.NextFireTimeUtc.HasValue)
                strMsg += "  ;下次執行於:" + DateTime.Parse(context.NextFireTimeUtc.ToString()).AddHours(8).ToString();
            Logging.Log(strMsg, LogState.Info, LogLayer.DB);

            strSucc = "S";
            InsertBatchLog(strJobMsg, dTimeStart);

            Log("*************** 備份及清檔作業完成 ***************\r\n===========================================================================================================");
        }
        catch (Exception exp)
        {
            //*批次完成記錄LOG信息
            strSucc = "F";
            InsertBatchLog(strJobMsg + exp.Message.ToString(), dTimeStart);
            Logging.Log(exp);
        }
    }
    #endregion

    /// <summary>
    /// 備份檔案
    /// </summary>
    /// <param name="SourcePath">來源路徑</param>
    /// <param name="BackupPath">備份路徑</param>
    /// <param name="BackupDay">需備份資料的日期(幾天前)</param>
    /// <param name="BackupALL">是否備份所有資料</param>
    private void FileBackup(string SourcePath, string BackupPath, int BackupDay, bool BackupALL)
    {
        DirectoryInfo dirinfo;
        dirinfo = new DirectoryInfo(SourcePath);
        FileInfo[] FileList = dirinfo.GetFiles();

        //*檢查資料夾中所有檔案符合備份日期的執行備份
        foreach (FileInfo F in FileList)
        {
            //*BackupALL=true 備份來源資料夾所有資料，否則備份BackupDay天數之前的所有資料(EX. BackupDay=30,備份30天前的所有資料)
            if ((DateTime.Now.Date - F.LastWriteTime.Date).Days > BackupDay || BackupALL)
            {
                if (!SkipExtension.Contains(F.Extension))
                {
                    //*確認備份資料夾
                    if (!Directory.Exists(BackupPath))
                        Directory.CreateDirectory(BackupPath);

                    F.CopyTo(BackupPath + "\\" + F.Name, true);
                    Log("複製 " + F.Name + " 從 " + F.DirectoryName + " 到 " + BackupPath);
                }
            }
        }
        //*處理來源資料夾中的子資料夾
        DirectoryInfo[] Childdirinfo = dirinfo.GetDirectories();
        if (Childdirinfo.Length > 0)
        {
            foreach (DirectoryInfo D in Childdirinfo)
            {
                FileBackup(D.FullName, BackupPath + "\\" + D.Name, BackupDay, BackupALL);
            }
        }
    }

    /// <summary>
    /// 清除過期檔案
    /// </summary>
    /// <param name="ClearPath">需清除的路徑</param>
    /// <param name="ClearDate">保留天數</param>
    private void ClearFile(string ClearPath, int KeepDay)
    {
        DirectoryInfo dirinfo;
        dirinfo = new DirectoryInfo(ClearPath);
        FileInfo[] FileList = dirinfo.GetFiles();

        //*檢查資料夾中所有檔案修改日期小於清除日期執行刪除動作
        foreach (FileInfo F in FileList)
        {
            if ((DateTime.Now.Date - F.LastWriteTime.Date).Days > KeepDay)
            {
                if (!SkipExtension.Contains(F.Extension))
                {
                    F.Delete();
                    Log("刪除 " + F.FullName);
                }
            }
        }
        //*處理子資料夾
        DirectoryInfo[] Childdirinfo = dirinfo.GetDirectories();
        if (Childdirinfo.Length > 0)
        {
            foreach (DirectoryInfo D in Childdirinfo)
            {
                ClearFile(D.FullName, KeepDay);
            }
        }
    }

    /// <summary>
    /// 清除過期資料夾
    /// </summary>
    /// <param name="ClearPath">需清除的路徑</param>
    /// <param name="ClearDate">保留天數</param>
    private void ClearFolder(string ClearPath, int KeepDay)
    {
        DirectoryInfo dirinfo = new DirectoryInfo(ClearPath);
        DirectoryInfo[] DirList = dirinfo.GetDirectories();

        string ClearDate = DateTime.Now.AddDays(-KeepDay).ToString("yyyyMMdd");

        //*檢查資料夾日期小於清除日期執行刪除動作
        foreach (DirectoryInfo D in DirList)
        {
            if (D.Name.CompareTo(ClearDate) <= 0)
            {
                D.Delete(true);
                Log("刪除資料夾 " + D.FullName);
            }
        }
    }

    /// <summary>
    /// 備份及清檔LOG
    /// </summary>
    /// <param name="strError">JOB失敗信息</param>
    /// <param name="dateStart">JOB開始時間</param>
    private void Log(string LogStr)
    {
        string BackupPath = UtilHelper.GetAppSettings("BackupPath").ToString() + dTimeStart.ToString("yyyyMMdd");
        StreamWriter sw = null;
        LogStr = DateTime.Now.ToLocalTime().ToString() + " ： " + LogStr;
        try
        {
            //*確認備份資料夾
            if (!Directory.Exists(BackupPath))
                Directory.CreateDirectory(BackupPath);
            sw = new StreamWriter(BackupPath + "\\BackupJobLog.txt", true);
            sw.WriteLine(LogStr);
        }
        catch (Exception exp)
        {
            Logging.Log(exp);
        }
        finally
        {
            if (sw != null)
            {
                sw.Close();
            }
        }
    }

    /// <summary>
    /// 插入L_BATCH_LOG資料庫
    /// </summary>
    /// <param name="strError">JOB失敗信息</param>
    /// <param name="dateStart">JOB開始時間</param>
    private void InsertBatchLog(string strError, DateTime dateStart)
    {
        //*插入L_BATCH_LOG資料庫
        BRL_BATCH_LOG.Delete(strFuncKey, strJobID, "R");
        BRL_BATCH_LOG.Insert(strFuncKey, strJobID, dateStart, strSucc, strError);
    }

    /// <summary>
    /// 開始此次作業向Job_Status中插入一筆新的資料
    /// </summary>
    /// <returns>true成功，false失敗</returns>
    private bool InsertNewBatch()
    {
        return BRL_BATCH_LOG.InsertRunning(strFuncKey, strJobID, dTimeStart, "R", "");
    }

    /// <summary>
    /// JOB執行狀態
    /// </summary>
    /// <param name="strStauts">狀態英文名稱</param>
    /// <returns>JOB執行狀態</returns>
    private string GetStatusName(string strStauts)
    {
        switch (strStauts)
        {
            case "F":
                return "失敗";
            case "S":
                return "成功";
            default:
                return "";
        }
    }
}
