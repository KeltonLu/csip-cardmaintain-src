//******************************************************************
//*  作    者：Ares Luke
//*  功能說明：HouseKeeping排程
//*  創建日期：2021/04/21
//*  修改記錄：
//*<author>            <time>            <TaskID>                <desc>
//* Ares Stanley    2022/02/14    20210058-CSIP作業服務平台現代化II    調整webconfig取參數方式
//*******************************************************************

using System;
using System.Configuration;
using System.Data;
using System.IO;
using System.Text;
using CSIPCommonModel.BusinessRules;
using Framework.Common.IO;
using Framework.Common.Logging;
using Framework.Common.Utility;
using Quartz;

/// <summary>
/// HouseKeepingFile 的摘要描述
/// </summary>
public class HouseKeepingFile : Quartz.IJob
{
    private readonly string _strFunctionKey = UtilHelper.GetAppSettings("FunctionKey").ToString();

    private static readonly JobHelper JobHelper = new JobHelper();

    private static string _strJobId = string.Empty;

    //逾期天數設定
    private static readonly int SetDay = Convert.ToInt32(UtilHelper.GetAppSettings("FileDays"));

    //ZIP備份檔目錄設定
    private static readonly int SetZipDay = Convert.ToInt32(UtilHelper.GetAppSettings("ZipDays"));

    //暫存目標路徑
    private static readonly string BackupTempFolder = UtilHelper.GetAppSettings("BackupTempFolder");

    //備份ZIP目錄
    static readonly string BackupZipFolder = UtilHelper.GetAppSettings("BackupZipFolder");

    //時間
    private static DateTime _dtNow = DateTime.Now;

    string _strMsgId = string.Empty;

    public void Execute(JobExecutionContext context)
    {
        try
        {
            JobDataMap jobDataMap = context.JobDetail.JobDataMap;
            _strJobId = context.JobDetail.JobDataMap["jobid"].ToString();
            JobHelper.strJobID = _strJobId;

            JobHelper.SaveLog(_strJobId + "JOB啟動", LogState.Info);


            #region 檢測JOB是否在執行中

            // 判斷Job工作狀態(0:停止 1:運行)
            var isContinue = CheckJobIsContinue(_strJobId, _strFunctionKey, _dtNow, ref _strMsgId);

            #endregion

            if (isContinue)
            {
                JobHelper.SaveLog("#檔案逾時天數:" + SetDay, LogState.Info);
                JobHelper.SaveLog("#ZIP檔逾時天數:" + SetZipDay, LogState.Info);
                JobHelper.SaveLog("#暫存目標路徑:" + BackupTempFolder, LogState.Info);
                JobHelper.SaveLog("#ZIP備份目錄:" + BackupZipFolder, LogState.Info);

                #region 刪除暫存目錄舊檔

                JobHelper.SaveLog("開始,刪除暫存目錄舊檔(" + BackupTempFolder + ")", LogState.Info);

                if (Directory.Exists(BackupTempFolder))
                {
                    JobHelper.SaveLog("目錄已存在，刪除舊檔目錄。", LogState.Info);
                    Directory.Delete(BackupTempFolder, true);
                }

                JobHelper.SaveLog("創建暫存資料夾", LogState.Info);
                Directory.CreateDirectory(BackupTempFolder);

                JobHelper.SaveLog("結束,刪除暫存目錄舊檔。", LogState.Info);

                #endregion

                #region 遍歷來源目錄將逾時檔搬移至暫存資料夾

                JobHelper.SaveLog("開始,遍歷來源目錄將逾時檔搬移至暫存資料夾", LogState.Info);
                string sourceFolderArr = UtilHelper.GetAppSettings("sourceFolderArr");
                string[] sourceFolderList = sourceFolderArr.Split(',');

                foreach (string sfl in sourceFolderList)
                {
                    string root = AppDomain.CurrentDomain.BaseDirectory + "\\" + sfl;
                    JobHelper.SaveLog("遍歷來源目錄:" + root, LogState.Info);
                    ListFiles(new DirectoryInfo(root));
                }

                JobHelper.SaveLog("結束,遍歷來源目錄將逾時檔搬移至暫存資料夾", LogState.Info);

                #endregion

                JobHelper.SaveLog("已完成搬移，將開始執行壓縮檔。", LogState.Info);

                #region 檢查ZIP備份路徑

                JobHelper.SaveLog("開始,檢查ZIP備份目錄是否存在(" + BackupZipFolder + ")", LogState.Info);
                if (!Directory.Exists(BackupZipFolder))
                {
                    JobHelper.SaveLog("目錄不存在並建立目錄。", LogState.Info);
                    Directory.CreateDirectory(BackupZipFolder);
                }

                JobHelper.SaveLog("結束,檢查ZIP備份目錄是否存在。", LogState.Info);

                #endregion


                #region 執行加入壓縮檔

                JobHelper.SaveLog("開始,執行加入壓縮檔。", LogState.Info);
                string zipFullName = BackupZipFolder + _dtNow.ToString("yyyyMMddHHmmss") + ".zip";
                CompressToZip.ZipFile(BackupTempFolder, zipFullName);
                JobHelper.SaveLog("壓縮檔路徑:" + zipFullName, LogState.Info);

                if (File.Exists(zipFullName))
                {
                    JobHelper.SaveLog("加入壓縮檔,成功。", LogState.Info);

                    #region 刪除暫存目錄

                    if (Directory.Exists(BackupTempFolder))
                    {
                        JobHelper.SaveLog("刪除暫存目錄", LogState.Info);
                        Directory.Delete(BackupTempFolder, true);
                    }

                    #endregion
                }
                else
                {
                    JobHelper.SaveLog("加入壓縮檔,失敗。", LogState.Error);
                }

                JobHelper.SaveLog("結束,執行加入壓縮檔。", LogState.Info);

                #endregion

                #region 刪除逾時的ZIP備份目錄檔案

                DeleteOverTimeZip(new DirectoryInfo(BackupZipFolder));

                #endregion

                InsertBatchLog(_strJobId, _strFunctionKey, _dtNow, "S", "備份路徑:" + zipFullName);
            }


            JobHelper.SaveLog(_strJobId + "結束", LogState.Info);
        }
        catch (Exception ex)
        {
            InsertBatchLog(_strJobId, _strFunctionKey, _dtNow, "F", "發生錯誤：" + ex.Message);
            JobHelper.SaveLog("發錯錯誤_" + ex.ToString(), LogState.Error);
        }
    }

    /// <summary>
    /// 遍歷目錄搬移逾時檔至暫存目錄
    /// </summary>
    /// <param name="info"></param>
    private static void ListFiles(FileSystemInfo info)
    {
        if (!info.Exists) return;

        DirectoryInfo dir = info as DirectoryInfo;
        // 不是目錄 
        if (dir == null) return;

        FileSystemInfo[] files = dir.GetFileSystemInfos();

        // 搬移成功數
        int moveSuccess = 0;
        int moveFail = 0;

        foreach (var f in files)
        {
            var strFullName = f.FullName;

            //判斷是檔案或資料目錄
            Boolean isFile = File.Exists(strFullName);

            if (isFile)
            {
                //建立時間
                var t = File.GetCreationTime(strFullName);
                var elapsedTicks = _dtNow.Ticks - t.Ticks;
                var elapsedSpan = new TimeSpan(elapsedTicks);

                if (elapsedSpan.TotalDays > SetDay)
                {
                    string newFileName =
                        strFullName.Replace(AppDomain.CurrentDomain.BaseDirectory, BackupTempFolder + "\\");


                    //檢核子目錄是否存在
                    string checkUpDir = newFileName.Replace(f.Name, "");
                    if (!Directory.Exists(checkUpDir))
                    {
                        Directory.CreateDirectory(checkUpDir);
                    }

                    if (Directory.Exists(checkUpDir))
                    {
                        string msg = "[檔案移轉] 來源:" + strFullName + ";目標:" + newFileName;

                        File.Move(strFullName, newFileName);

                        if (File.Exists(newFileName))
                        {
                            JobHelper.SaveLog("[成功]" + msg, LogState.Info);
                            moveSuccess++;
                        }
                        else
                        {
                            JobHelper.SaveLog("[失敗]" + msg, LogState.Error);
                            moveFail++;
                        }
                    }
                }
            }
            else
            {
                ListFiles(f);
            }
        }

        JobHelper.SaveLog(info.Name + ":逾期檔案移轉成功共【" + moveSuccess + "】筆，失敗共【" + moveFail + "】筆。", LogState.Info);
    }


    /// <summary>
    /// 刪除逾期的備份目錄(ZIP檔)
    /// </summary>
    /// <param name="info"></param>
    private static void DeleteOverTimeZip(FileSystemInfo info)
    {
        JobHelper.SaveLog("開始,檢核逾時的ZIP備份檔", LogState.Info);

        if (!info.Exists) return;

        DirectoryInfo dir = info as DirectoryInfo;
        // 不是目錄 
        if (dir == null) return;

        FileSystemInfo[] files = dir.GetFileSystemInfos();

        // 刪除總數
        int delSuccess = 0;
        int delFail = 0;


        foreach (var f in files)
        {
            var strFullName = f.FullName;

            //判斷是檔案或資料目錄
            Boolean isFile = File.Exists(strFullName);

            if (isFile)
            {
                //建立時間
                var t = File.GetCreationTime(strFullName);
                var elapsedTicks = _dtNow.Ticks - t.Ticks;
                var elapsedSpan = new TimeSpan(elapsedTicks);

                if (elapsedSpan.TotalDays > SetZipDay)
                {
                    File.Delete(strFullName);
                    if (!File.Exists(strFullName))
                    {
                        JobHelper.SaveLog("刪除成功(" + strFullName + ")", LogState.Info);
                        delSuccess++;
                    }
                    else
                    {
                        JobHelper.SaveLog("刪除失敗(" + strFullName + ")", LogState.Error);
                        delFail++;
                    }
                }
            }
            else
            {
                ListFiles(f);
            }
        }

        JobHelper.SaveLog(info.Name + ":逾期ZIP檔案刪除成功共【" + delSuccess + "】筆，失敗共【" + delFail + "】筆。", LogState.Info);
        JobHelper.SaveLog("結束,檢核逾時的ZIP備份檔", LogState.Info);
    }


    /// <summary>
    /// 判斷Job工作狀態(0:停止 1:運行)
    /// </summary>
    /// <param name="JobID"></param>
    /// <param name="strFunctionKey"></param>
    /// <param name="dateStart"></param>
    /// <param name="msgID"></param>
    /// <returns></returns>
    private static bool CheckJobIsContinue(string JobID, string strFunctionKey, DateTime dateStart, ref string msgID)
    {
        bool result = true;
        if (JobHelper.SerchJobStatus(JobID).Equals("") || JobHelper.SerchJobStatus(JobID).Equals("0"))
        {
            // Job停止
            JobHelper.Write(JobID, "【FAIL】 Job工作狀態為：停止！");
            result = false;
        }

        // 檢測Job是否在執行中
        try
        {
            DataTable dtInfo = BRL_BATCH_LOG.GetRunningDate(strFunctionKey, JobID, "R", ref msgID);
            if (dtInfo == null || dtInfo.Rows.Count > 0) //20210531_Ares_Stanley-修正Job執行檢核條件
            {
                JobHelper.SaveLog("JOB 工作狀態為：正在執行！", LogState.Info);
                // 返回不執行
                result = false;
            }
            else
            {
                // 記錄Job執行資訊
                BRL_BATCH_LOG.InsertRunning(strFunctionKey, JobID, dateStart, "R", "");
            }
        }
        catch (Exception ex)
        {
            result = false;
            JobHelper.Write(JobID, "【FAIL】" + ex.ToString());
        }

        return result;
    }

    /// <summary>
    /// 新增 BRL_BATCH_LOG 訊息
    /// </summary>
    /// <param name="jobID"></param>
    /// <param name="strFunctionKey"></param>
    /// <param name="dateStart"></param>
    /// <param name="status"></param>
    /// <param name="message"></param>
    private static void InsertBatchLog(string jobID, string strFunctionKey, DateTime dateStart, string status,
        string message)
    {
        StringBuilder sbMessage = new StringBuilder();
        
        if (message.Trim() != "")
        {
            string msgTitle = status == "S" ? "成功" : "失敗";
            sbMessage.Append(msgTitle + "訊息：" + message); //*失敗訊息
        }

        BRL_BATCH_LOG.Delete(strFunctionKey, jobID, "R");
        BRL_BATCH_LOG.Insert(strFunctionKey, jobID, dateStart, status, sbMessage.ToString());
    }
}