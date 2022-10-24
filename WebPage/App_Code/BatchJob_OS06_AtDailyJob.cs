//******************************************************************
//*  作    者：Ares Stanley
//*  創建日期：2022/01/06
//*  功能說明：Stored Procedure轉批次，取得檔案並將資料匯入table[CPMAST]
//*  修改紀錄：
//*<author>            <time>            <TaskID>                <desc>
//* Ares Stanley 2022/02/10  20210058-CSIP作業服務平台現代化II    調整JobName，移除JOBLOG、JOBSTEPLOG
//* Ares Stanley 2022/03/15  20210058-CSIP作業服務平台現代化II    調整檢核失敗資料處理方式
//* Ares Stanley 2022/03/25  20210058-CSIP作業服務平台現代化II    調整Import_log紀錄
//* Ares Stanley 2022/04/18  20210058-CSIP作業服務平台現代化II    更新L_Batch_Log、寄信移動到最後
//* Ares Stanley 2022/05/05  20210058-CSIP作業服務平台現代化II    部分Log調整為Error層級
//*******************************************************************

using System;
using System.Threading;
using System.Data;
using System.Text;
using Quartz;
using CSIPCardMaintain.BusinessRules;
using CSIPCommonModel.BusinessRules;
using CSIPCommonModel.EntityLayer;
using Framework.Common.Logging;
using Framework.Common.Utility;
using Framework.Common.Message;
using System.Collections.Generic;
using System.Data.SqlClient;

/// <summary>
/// BatchJob_OS06_AtDailyJOB 的摘要描述
/// </summary>
public class BatchJob_OS06_AtDailyJob : Quartz.IJob
{
    protected string FunctionKey = UtilHelper.GetAppSettings("FunctionKey").ToString();
    protected DateTime StartTime = DateTime.Now; //Job啟動時間
    protected DateTime EndTime;
    protected JobHelper JobHelper = new JobHelper();
    protected string _MailTitle = "JobOS06_AtDailyJob批次執行結果：";
    protected string strMail;
    protected int datDataCount; //dat檔資料筆數
    protected bool makeErrorFile; //是否產生檢核失敗TXT檔
    protected string errorFilePath; //檢核失敗TXT檔路徑
    protected int datErrorDataCount; //dat檢核失敗資料數
    public void Execute(JobExecutionContext context)
    {
        string jobID = context.JobDetail.JobDataMap["jobid"].ToString();
        JobHelper.strJobID = jobID;
        string exeFileName = string.Empty;
        string datFileName = string.Empty;
        string msgID = string.Empty;
        string errorMsg = string.Empty;
        string date = DateTime.Now.AddDays(-1).ToString("yyyyMMdd");
        string localPath = string.Empty;
        string unZipPwd = string.Empty;
        string truncateTable = string.Empty;
        bool isUnzip = false;
        bool truncateStatus = false;
        int correctDataCount = 0;
        int errorDataCount = 0;

        // 20221004 調整將資料拆成多個檔案 By Kelton start
        List<datInfo> datInfos = new List<datInfo>();
        bool OS06FileSplitFlag = true;
        int settingfileCount = 0; // 資料庫設定可收的最大檔案數
        bool doCheck = true; //是否進行相關檢核
        string doingFileName = string.Empty; //目前正在處理的檔案
        // 20221004 調整將資料拆成多個檔案 By Kelton end

        JobDataMap jobDataMap = context.JobDetail.JobDataMap;
        strMail = jobDataMap.GetString("mail").Trim();
        EntityAGENT_INFO eAgentInfo = GetAGENT_INFO(jobDataMap);
        OS06_AtDailyJob atDailyJob = new OS06_AtDailyJob(jobID, eAgentInfo);

        try
        {
            JobHelper.SaveLog("*********** " + jobID + " 取得CPMAST資料 批次 START ************** ", LogState.Info);


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
                                date = tempDt.ToString("yyyyMMdd");
                            }
                            else
                            {
                                errorMsg = "更新參數至FileInfo：失敗！ 參數：" + tempDt.ToString("yyyyMMdd");
                                JobHelper.SaveLog("更新參數至FileInfo：失敗！ 參數：" + tempDt.ToString("yyyyMMdd"), LogState.Error);
                                return;
                            }
                        }
                        else
                        {
                            errorMsg = "檢核參數：異常！ 參數：" + strParam;
                            JobHelper.SaveLog("檢核參數：異常！ 參數：" + strParam, LogState.Error);
                            return;
                        }
                    }
                    else
                    {
                        errorMsg = "檢核參數：異常！ 參數：" + strParam;
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
                InsertBatchLog(jobID, "F", "JOB 工作狀態為：正在執行！");
                return;
            }

            // 20221004 調整將資料拆成多個檔案 By Kelton start
            // 取得資料拆檔功能開關設定資料
            SqlCommand sqlcmd = new SqlCommand();
            sqlcmd.CommandType = CommandType.Text;
            sqlcmd.CommandText = string.Format("SELECT PROPERTY_CODE, SEQUENCE FROM M_PROPERTY_CODE WHERE FUNCTION_KEY = '{0}' AND PROPERTY_KEY = '{1}' AND SEQUENCE = '3'", FunctionKey, jobID);
            DataSet ds = BRM_PROPERTY_CODE.SearchOnDataSet(sqlcmd, "Connection_CSIP");

            if (ds != null)
            {
                DataTable dt = ds.Tables[0];
                if (dt != null && dt.Rows.Count > 0)
                {
                    if (dt.Rows[0][0].ToString().ToUpper() == "N")
                    {
                        OS06FileSplitFlag = false;
                    }
                }
                else
                {
                    JobHelper.SaveLog("[FAIL] M_PROPERTY_CODE 查無資料拆檔功能開關設定資料", LogState.Error);
                }
            }
            else
            {
                JobHelper.SaveLog("[FAIL] 取得 M_PROPERTY_CODE 資料拆檔功能開關設定資料失敗", LogState.Error);
            }

            if (!OS06FileSplitFlag)
            {
                #region 下載檔案
                JobHelper.SaveLog(string.Format("檔案 OS06{0}.EXE 下載開始", date), LogState.Info);

                exeFileName = DownloadFile(jobID, atDailyJob, date, ref localPath, ref unZipPwd, ref errorMsg);
                if (!string.IsNullOrEmpty(errorMsg) || string.IsNullOrEmpty(exeFileName))
                {
                    JobHelper.SaveLog(string.Format("檔案 OS06{0}.EXE 下載 [失敗]", date), LogState.Error);
                    return;
                }
                else
                {
                    JobHelper.SaveLog(string.Format("檔案 OS06{0}.EXE 下載 [成功]", date), LogState.Info);
                }
                #endregion

                #region 解壓縮檔案
                Thread.Sleep(5000);
                JobHelper.SaveLog(string.Format("檔案 {0} 解壓縮開始", exeFileName), LogState.Info);

                isUnzip = atDailyJob.ZipExeFile(localPath, exeFileName, unZipPwd, ref errorMsg);
                if (!string.IsNullOrEmpty(errorMsg) || !isUnzip)
                {
                    JobHelper.SaveLog(string.Format("檔案 {0} 解壓縮 [失敗]", exeFileName), LogState.Error);
                    return;
                }
                else
                {
                    JobHelper.SaveLog(string.Format("檔案 {0} 解壓縮 [成功]", exeFileName), LogState.Info);
                }
                #endregion

                #region 檢查import_log是否有匯入紀錄, 若無則新增資料
                JobHelper.SaveLog("檢查Import_Log是否有匯入紀錄, 若無則新增資料 開始", LogState.Info);

                bool insertImportLogStatus = atDailyJob.CheckImportLog(date, exeFileName, ref errorMsg);
                if (!string.IsNullOrEmpty(errorMsg) || !insertImportLogStatus)
                {
                    JobHelper.SaveLog("檢查Import_Log是否有匯入紀錄, 若無則新增資料 [失敗]", LogState.Error);
                    return;
                }
                else
                {
                    JobHelper.SaveLog("檢查Import_Log是否有匯入紀錄, 若無則新增資料 [成功]", LogState.Info);
                }
                #endregion

                #region truncate table[cpmast_tmp]
                truncateTable = "CPMAST_TMP";
                JobHelper.SaveLog(string.Format("Truncate Table：{0} 開始！", truncateTable), LogState.Info);
                truncateStatus = atDailyJob.TruncateTable(truncateTable, ref errorMsg);
                if (!string.IsNullOrEmpty(errorMsg) || !truncateStatus)
                {
                    JobHelper.SaveLog(string.Format("Truncate Table：{0} [失敗]！", truncateTable), LogState.Error);
                    return;
                }
                else
                {
                    JobHelper.SaveLog(string.Format("Truncate Table：{0} [成功]！", truncateTable), LogState.Info);
                }
                #endregion

                #region 將資料匯入table[cpmast_tmp]
                JobHelper.SaveLog(string.Format("讀取 OS06{0}.dat 開始", date), LogState.Info);
                //讀檔
                DataTable datTable = atDailyJob.GetMaintainData(localPath, ref datFileName, ref errorMsg, ref this.makeErrorFile, ref this.errorFilePath, ref this.datErrorDataCount, ref this.datDataCount);
                if (datTable == null && string.IsNullOrEmpty(errorMsg))
                {
                    errorMsg = string.Format("讀取 OS06{0}.dat 時發生錯誤，請確認", date);
                }

                if (!string.IsNullOrEmpty(errorMsg))
                {
                    JobHelper.SaveLog(string.Format("讀取 {0} [失敗]", datFileName), LogState.Error);
                    return;
                }

                if (datTable.Rows.Count <= 0)
                {
                    errorMsg = string.Format("{0} 檔沒有資料", datFileName);
                    JobHelper.SaveLog(string.Format("檔案 {0} 沒有資料", datFileName), LogState.Info);
                    return;
                }
                else
                {
                    JobHelper.SaveLog(string.Format("讀取 {0} [成功]", datFileName), LogState.Info);
                }

                //匯入資料至table[cpmast_tmp]
                JobHelper.SaveLog("匯入資料至 CPMAST_TMP 開始", LogState.Info);

                bool insertStatus = atDailyJob.InsertCpmastTmp("CPMAST_TMP", datTable, ref errorMsg);
                if (!insertStatus && string.IsNullOrEmpty(errorMsg))
                {
                    errorMsg = "匯入資料至 CPMAST_TMP 時發生錯誤，請確認";
                }

                if (!string.IsNullOrEmpty(errorMsg))
                {
                    JobHelper.SaveLog("匯入資料至 CPMAST_TMP [失敗]！", LogState.Error);
                    return;
                }
                else
                {
                    JobHelper.SaveLog("匯入資料至 CPMAST_TMP [成功]！", LogState.Info);
                }
                #endregion

                #region 檢核table[cpmast_tmp]資料日期格式是否正確，若有錯誤則將錯誤資料另存至cpmast_Err，並從cpmast_tmp刪除錯誤資料

                JobHelper.SaveLog("檢核table[cpmast_tmp]資料日期格式是否正確，若有錯誤則將錯誤資料另存至cpmast_Err，並從cpmast_tmp刪除錯誤資料 開始", LogState.Info);
                bool checkCpmastTmpResult = atDailyJob.CheckCpmastTmp(exeFileName, ref errorDataCount, ref errorMsg);
                if (!checkCpmastTmpResult && string.IsNullOrEmpty(errorMsg))
                {
                    errorMsg = "檢核table[cpmast_tmp]資料日期格式是否正確，若有錯誤則將錯誤資料另存至cpmast_Err，並從cpmast_tmp刪除錯誤資料時發生錯誤，請確認";
                }
                if (!string.IsNullOrEmpty(errorMsg))
                {
                    JobHelper.SaveLog("檢核table[cpmast_tmp]資料日期格式是否正確，若有錯誤則將錯誤資料另存至cpmast_Err，並從cpmast_tmp刪除錯誤資料 [失敗]！", LogState.Error);
                    return;
                }
                else
                {
                    if (errorDataCount <= 0)
                    {
                        JobHelper.SaveLog("table[cpmast_tmp] 無日期格式錯誤資料", LogState.Info);
                    }
                    else
                    {
                        //若有日期格式錯誤的資料，Log層級改為Error by Ares Stanley 20220505
                        JobHelper.SaveLog(string.Format("table[cpmast_tmp] 有日期格式錯誤資料，共 {0} 筆", errorDataCount), LogState.Error);
                    }
                    JobHelper.SaveLog("檢核table[cpmast_tmp]資料日期格式是否正確，若有錯誤則將錯誤資料另存至cpmast_Err，並從cpmast_tmp刪除錯誤資料 [成功]！", LogState.Info);
                }
                #endregion

                #region 確認暫存檔資料筆數不為0，轉換table[cpmast_tmp]日期格式
                JobHelper.SaveLog("轉換table[cpmast_tmp]資料日期格式 開始！", LogState.Info);
                bool convertStatus = atDailyJob.ConvertCpmastTmpData(ref correctDataCount, ref errorMsg);
                if (!string.IsNullOrEmpty(errorMsg))
                {
                    JobHelper.SaveLog("轉換table[cpmast_tmp]資料日期格式 [失敗]！", LogState.Error);
                    return;
                }
                else
                {
                    JobHelper.SaveLog("轉換table[cpmast_tmp]資料日期格式 [成功]！", LogState.Info);
                }
                #endregion

                #region 將table[cpmast_tmp]資料匯入table[cpmast]
                JobHelper.SaveLog("將正確資料寫入table[cpmast] 開始！", LogState.Info);
                bool insertCpmastStatus = atDailyJob.InsertCorrectDataToCpmast(exeFileName, ref errorMsg);
                if (!string.IsNullOrEmpty(errorMsg))
                {
                    JobHelper.SaveLog("將正確資料寫入table[cpmast] [失敗]！", LogState.Error);
                    return;
                }
                else
                {
                    JobHelper.SaveLog("將正確資料寫入table[cpmast] [成功]！", LogState.Info);
                }
                #endregion

                //若有匯入失敗或檢核失敗資料則 Log 層級為 Error By Ares Stanley 20220505
                if (errorDataCount > 0 || this.datErrorDataCount > 0)
                {
                    //有匯入失敗或檢核失敗資料
                    JobHelper.SaveLog(string.Format("檔案：{0} 匯入結束，資料共 {1} 筆，匯入成功 {2} 筆，匯入失敗 {3} 筆，檢核失敗 {4} 筆", exeFileName, this.datDataCount, correctDataCount, errorDataCount, this.datErrorDataCount), LogState.Error);
                }
                else
                {
                    //無匯入失敗或檢核失敗資料
                    JobHelper.SaveLog(string.Format("檔案：{0} 匯入結束，資料共 {1} 筆，匯入成功 {2} 筆，匯入失敗 {3} 筆，檢核失敗 {4} 筆", exeFileName, this.datDataCount, correctDataCount, errorDataCount, this.datErrorDataCount), LogState.Info);
                }
            }
            else
            {
                #region 下載檔案
                int fileOKDataCount = 0;
                List<string> exeFileNames = DownloadFiles(jobID, atDailyJob, date, ref localPath, ref unZipPwd, ref errorMsg, ref fileOKDataCount, ref settingfileCount, ref doCheck);
                if (!string.IsNullOrEmpty(errorMsg) || exeFileNames == null)
                {
                    return;
                }
                #endregion
                int totalDatDataCount = 0;
                foreach (var fileName in exeFileNames)
                {
                    #region 解壓縮檔案
                    Thread.Sleep(5000);

                    JobHelper.SaveLog(string.Format("檔案 {0} 解壓縮開始", fileName), LogState.Info);

                    isUnzip = atDailyJob.ZipExeFile(localPath, fileName, unZipPwd, ref errorMsg);
                    if (!string.IsNullOrEmpty(errorMsg) || !isUnzip)
                    {
                        JobHelper.SaveLog(string.Format("檔案 {0} 解壓縮 [失敗]", fileName), LogState.Error);
                        return;
                    }
                    else
                    {
                        JobHelper.SaveLog(string.Format("檔案 {0} 解壓縮 [成功]", fileName), LogState.Info);
                    }
                    #endregion

                    if (doCheck)
                    {
                        #region 檢查檔案內容筆數是否和 FILEOK 的筆數相同
                        datFileName = fileName.Substring(0, fileName.Length - 4) + ".dat";
                        JobHelper.SaveLog(string.Format("讀取 {0} 開始", datFileName), LogState.Info);
                        //讀檔
                        DataTable datTable = atDailyJob.GetMaintainDataByFileName(localPath, fileName, ref errorMsg, ref this.makeErrorFile, ref this.errorFilePath, ref this.datErrorDataCount, ref this.datDataCount);
                        if (datTable == null && string.IsNullOrEmpty(errorMsg))
                        {
                            errorMsg = string.Format("讀取 {0} 時發生錯誤，請確認", datFileName);
                        }

                        if (!string.IsNullOrEmpty(errorMsg))
                        {
                            JobHelper.SaveLog(string.Format("讀取 {0} [失敗]", datFileName), LogState.Error);
                            return;
                        }

                        if (datTable.Rows.Count <= 0)
                        {
                            errorMsg = string.Format("{0} 檔沒有資料", datFileName);
                            JobHelper.SaveLog(string.Format("檔案 {0} 沒有資料", datFileName), LogState.Info);
                            return;
                        }
                        else
                        {
                            JobHelper.SaveLog(string.Format("讀取 {0} [成功]", datFileName), LogState.Info);
                        }

                        totalDatDataCount += datTable.Rows.Count;
                        #endregion
                    }
                }
                if (doCheck)
                {
                    if (fileOKDataCount != totalDatDataCount)
                    {
                        errorMsg = ".dat 檔案內容資料總筆數與 FILEOK 檔案內容筆數不同";
                        JobHelper.SaveLog(".dat 檔案內容資料總筆數與 FILEOK 檔案內容筆數不同", LogState.Error);
                        return;
                    }
                }

                bool imported = false;
                foreach (var fileName in exeFileNames)
                {
                    doingFileName = fileName;
                    #region 檢查import_log是否有匯入紀錄, 若無則新增資料
                    JobHelper.SaveLog("檢查Import_Log是否有匯入紀錄, 若無則新增資料 開始", LogState.Info);

                    imported = false;
                    bool insertImportLogStatus = atDailyJob.CheckImportLog_New(date, fileName, ref errorMsg, ref imported);
                    if (!string.IsNullOrEmpty(errorMsg) || !insertImportLogStatus)
                    {
                        JobHelper.SaveLog("檢查Import_Log是否有匯入紀錄, 若無則新增資料 [失敗]", LogState.Error);
                        return;
                    }
                    else
                    {
                        JobHelper.SaveLog("檢查Import_Log是否有匯入紀錄, 若無則新增資料 [成功]", LogState.Info);
                        // 判斷此檔案已匯入過則不做後續動作，避免重複匯入
                        if (imported)
                            continue;
                    }
                    #endregion

                    #region truncate table[cpmast_tmp]
                    truncateTable = "CPMAST_TMP";
                    JobHelper.SaveLog(string.Format("Truncate Table：{0} 開始！", truncateTable), LogState.Info);
                    truncateStatus = atDailyJob.TruncateTable(truncateTable, ref errorMsg);
                    if (!string.IsNullOrEmpty(errorMsg) || !truncateStatus)
                    {
                        JobHelper.SaveLog(string.Format("Truncate Table：{0} [失敗]！", truncateTable), LogState.Error);
                        return;
                    }
                    else
                    {
                        JobHelper.SaveLog(string.Format("Truncate Table：{0} [成功]！", truncateTable), LogState.Info);
                    }
                    #endregion

                    #region 將資料匯入table[cpmast_tmp]
                    datFileName = fileName.Substring(0, fileName.Length - 4) + ".dat";
                    JobHelper.SaveLog(string.Format("讀取 {0} 開始", datFileName), LogState.Info);
                    //讀檔
                    DataTable datTable = atDailyJob.GetMaintainDataByFileName(localPath, fileName, ref errorMsg, ref this.makeErrorFile, ref this.errorFilePath, ref this.datErrorDataCount, ref this.datDataCount);
                    if (datTable == null && string.IsNullOrEmpty(errorMsg))
                    {
                        errorMsg = string.Format("讀取 {0} 時發生錯誤，請確認", datFileName);
                    }

                    if (!string.IsNullOrEmpty(errorMsg))
                    {
                        JobHelper.SaveLog(string.Format("讀取 {0} [失敗]", datFileName), LogState.Error);
                        return;
                    }

                    if (datTable.Rows.Count <= 0)
                    {
                        errorMsg = string.Format("{0} 檔沒有資料", datFileName);
                        JobHelper.SaveLog(string.Format("檔案 {0} 沒有資料", datFileName), LogState.Info);
                        return;
                    }
                    else
                    {
                        JobHelper.SaveLog(string.Format("讀取 {0} [成功]", datFileName), LogState.Info);
                    }

                    //匯入資料至table[cpmast_tmp]
                    JobHelper.SaveLog("匯入資料至 CPMAST_TMP 開始", LogState.Info);

                    bool insertStatus = atDailyJob.InsertCpmastTmp("CPMAST_TMP", datTable, ref errorMsg);
                    if (!insertStatus && string.IsNullOrEmpty(errorMsg))
                    {
                        errorMsg = "匯入資料至 CPMAST_TMP 時發生錯誤，請確認";
                    }

                    if (!string.IsNullOrEmpty(errorMsg))
                    {
                        JobHelper.SaveLog("匯入資料至 CPMAST_TMP [失敗]！", LogState.Error);
                        return;
                    }
                    else
                    {
                        JobHelper.SaveLog("匯入資料至 CPMAST_TMP [成功]！", LogState.Info);
                    }
                    #endregion

                    #region 檢核table[cpmast_tmp]資料日期格式是否正確，若有錯誤則將錯誤資料另存至cpmast_Err，並從cpmast_tmp刪除錯誤資料

                    JobHelper.SaveLog("檢核table[cpmast_tmp]資料日期格式是否正確，若有錯誤則將錯誤資料另存至cpmast_Err，並從cpmast_tmp刪除錯誤資料 開始", LogState.Info);
                    bool checkCpmastTmpResult = atDailyJob.CheckCpmastTmp(fileName, ref errorDataCount, ref errorMsg);
                    if (!checkCpmastTmpResult && string.IsNullOrEmpty(errorMsg))
                    {
                        errorMsg = "檢核table[cpmast_tmp]資料日期格式是否正確，若有錯誤則將錯誤資料另存至cpmast_Err，並從cpmast_tmp刪除錯誤資料時發生錯誤，請確認";
                    }
                    if (!string.IsNullOrEmpty(errorMsg))
                    {
                        JobHelper.SaveLog("檢核table[cpmast_tmp]資料日期格式是否正確，若有錯誤則將錯誤資料另存至cpmast_Err，並從cpmast_tmp刪除錯誤資料 [失敗]！", LogState.Error);
                        return;
                    }
                    else
                    {
                        if (errorDataCount <= 0)
                        {
                            JobHelper.SaveLog("table[cpmast_tmp] 無日期格式錯誤資料", LogState.Info);
                        }
                        else
                        {
                            //若有日期格式錯誤的資料，Log層級改為Error
                            JobHelper.SaveLog(string.Format("table[cpmast_tmp] 有日期格式錯誤資料，共 {0} 筆", errorDataCount), LogState.Error);
                        }
                        JobHelper.SaveLog("檢核table[cpmast_tmp]資料日期格式是否正確，若有錯誤則將錯誤資料另存至cpmast_Err，並從cpmast_tmp刪除錯誤資料 [成功]！", LogState.Info);
                    }
                    #endregion

                    #region 確認暫存檔資料筆數不為0，轉換table[cpmast_tmp]日期格式
                    JobHelper.SaveLog("轉換table[cpmast_tmp]資料日期格式 開始！", LogState.Info);
                    bool convertStatus = atDailyJob.ConvertCpmastTmpData(ref correctDataCount, ref errorMsg);
                    if (!string.IsNullOrEmpty(errorMsg))
                    {
                        JobHelper.SaveLog("轉換table[cpmast_tmp]資料日期格式 [失敗]！", LogState.Error);
                        return;
                    }
                    else
                    {
                        JobHelper.SaveLog("轉換table[cpmast_tmp]資料日期格式 [成功]！", LogState.Info);
                    }
                    #endregion

                    #region 將table[cpmast_tmp]資料匯入table[cpmast]
                    JobHelper.SaveLog("將正確資料寫入table[cpmast] 開始！", LogState.Info);
                    bool insertCpmastStatus = atDailyJob.InsertCorrectDataToCpmast(fileName, ref errorMsg);
                    if (!string.IsNullOrEmpty(errorMsg))
                    {
                        JobHelper.SaveLog("將正確資料寫入table[cpmast] [失敗]！", LogState.Error);
                        return;
                    }
                    else
                    {
                        JobHelper.SaveLog("將正確資料寫入table[cpmast] [成功]！", LogState.Info);
                    }
                    #endregion

                    // 紀錄 dat 匯入情況
                    datInfos.Add(new datInfo
                    {
                        FILE_NAME = fileName,
                        DatDataCount = this.datDataCount,
                        CorrectDataCount = correctDataCount,
                        ErrorDataCount = errorDataCount,
                        DatErrorDataCount = this.datErrorDataCount
                    });

                    //若有匯入失敗或檢核失敗資料則 Log 層級為 Error
                    if (errorDataCount > 0 || this.datErrorDataCount > 0)
                    {
                        //有匯入失敗或檢核失敗資料
                        JobHelper.SaveLog(string.Format("檔案：{0} 匯入結束，資料共 {1} 筆，匯入成功 {2} 筆，匯入失敗 {3} 筆，檢核失敗 {4} 筆", fileName, this.datDataCount, correctDataCount, errorDataCount, this.datErrorDataCount), LogState.Error);
                    }
                    else
                    {
                        //無匯入失敗或檢核失敗資料
                        JobHelper.SaveLog(string.Format("檔案：{0} 匯入結束，資料共 {1} 筆，匯入成功 {2} 筆，匯入失敗 {3} 筆，檢核失敗 {4} 筆", fileName, this.datDataCount, correctDataCount, errorDataCount, this.datErrorDataCount), LogState.Info);
                    }

                    #region 檢查成功、失敗筆數, 更新table[import_log]
                    JobHelper.SaveLog("更新table[Import_Log] 開始！", LogState.Info);
                    bool updateStatus = atDailyJob.UpdateImportLogByFileName(correctDataCount, errorDataCount, this.datErrorDataCount, date, fileName, ref errorMsg);
                    if (!updateStatus)
                    {
                        JobHelper.SaveLog(string.Format("更新 table[Import_Log] [失敗] 檔案：{0}！", fileName), LogState.Error);
                    }
                    else
                    {
                        JobHelper.SaveLog(string.Format("更新 table[Import_Log] [成功] 檔案：{0}！", fileName), LogState.Info);
                    }
                    #endregion
                }
            }
            // 20221004 調整將資料拆成多個檔案 By Kelton end
        }
        catch (Exception ex)
        {
            Logging.Log(ex.Message);
            JobHelper.SaveLog("BatchJob_OS06_AtDailyJob 發生例外錯誤：" + ex.Message);
            errorMsg += "　排程 BatchJob_OS06_AtDailyJob 發生例外錯誤，請確認 JobLog(Log\\JobOS06_AtDailyJob\\) 或 DefaultLog(Log\\Default\\)";
            return;
        }
        finally
        {
            // 20221004 調整將資料拆成多個檔案 By Kelton start
            if (!OS06FileSplitFlag)
            {
                #region 檢查成功、失敗筆數, 更新table[import_log]
                JobHelper.SaveLog("更新table[Import_Log] 開始！", LogState.Info);
                bool updateStatus = atDailyJob.UpdateImportLog(correctDataCount, errorDataCount, this.datErrorDataCount, date, ref errorMsg);
                if (!updateStatus)
                {
                    JobHelper.SaveLog("更新table[Import_Log] [失敗]！", LogState.Error);
                }
                else
                {
                    JobHelper.SaveLog("更新table[Import_Log] [成功]！", LogState.Info);
                }
                #endregion
            }
            // 20221004 調整將資料拆成多個檔案 By Kelton End


            //清空手動輸入參數
            if (BRM_FileInfo.UpdateParam(jobID, ""))
            {
                JobHelper.SaveLog("清空手動參數 [成功]", LogState.Info);
            }
            else
            {
                errorMsg += "　清空手動參數 [失敗]";
                JobHelper.SaveLog("清空手動參數 [失敗]", LogState.Error);

            }

            // 20221004 調整將資料拆成多個檔案 By Kelton start
            if (!OS06FileSplitFlag)
            {
                if (string.IsNullOrEmpty(errorMsg))
                {
                    //更新L_BATCH_LOG
                    InsertBatchLog(jobID, "S", string.Format("檔案：{0} 匯入結束，資料共 {1} 筆，匯入成功 {2} 筆，匯入失敗 {3} 筆，檢核失敗 {4} 筆", exeFileName, this.datDataCount, correctDataCount, errorDataCount, this.datErrorDataCount));

                    //若有匯入失敗或檢核失敗的資料，額外新增一筆L_Batch_Log by Ares Stanley 20220505
                    if (errorDataCount > 0 || this.datErrorDataCount > 0)
                    {
                        InsertBatchLog(jobID, "F", string.Format("檔案：{0} 匯入結束，匯入失敗資料共 {1} 筆，檢核失敗資料共 {2} 筆", exeFileName, errorDataCount, this.datErrorDataCount));
                    }

                    string resultMsg = string.Format(" JobOS06_AtDailyJob 批次執行成功：資料共 {1} 筆，匯入成功 {2} 筆，匯入失敗 {3} 筆，檢核失敗 {4} 筆，檔案：{0} 匯入結束", exeFileName, this.datDataCount, correctDataCount, errorDataCount, this.datErrorDataCount);
                    //寄成功信
                    SendMail(_MailTitle + "成功！" + resultMsg, resultMsg, "成功", this.StartTime);
                }
                else
                {
                    //更新L_BATCH_LOG
                    InsertBatchLog(JobHelper.strJobID, "F", errorMsg);

                    //寄失敗信
                    SendMail(_MailTitle + "失敗！" + errorMsg, string.Format(" JobOS06_AtDailyJob 批次 發生錯誤：{0}", errorMsg), "失敗", this.StartTime);
                }
            }
            else
            {
                if (string.IsNullOrEmpty(errorMsg))
                {
                    int totalDataCount = 0;
                    int totalCorrectDataCount = 0;
                    int totalErrorDataCount = 0;
                    int totalDatErrorDataCount = 0;
                    string fileNames = string.Empty;
                    foreach (var datInfo in datInfos)
                    {
                        if (!string.IsNullOrEmpty(fileNames))
                            fileNames = fileNames + "，";
                        //更新L_BATCH_LOG
                        InsertBatchLog(jobID, "S", string.Format("檔案：{0} 匯入結束，資料共 {1} 筆，匯入成功 {2} 筆，匯入失敗 {3} 筆，檢核失敗 {4} 筆", datInfo.FILE_NAME, datInfo.DatDataCount, datInfo.CorrectDataCount, datInfo.ErrorDataCount, datInfo.DatErrorDataCount));

                        //若有匯入失敗或檢核失敗的資料，額外新增一筆L_Batch_Log
                        if (datInfo.ErrorDataCount > 0 || datInfo.DatErrorDataCount > 0)
                        {
                            InsertBatchLog(jobID, "F", string.Format("檔案：{0} 匯入結束，匯入失敗資料共 {1} 筆，檢核失敗資料共 {2} 筆", datInfo.FILE_NAME, datInfo.ErrorDataCount, datInfo.DatErrorDataCount));
                        }
                        totalDataCount += datInfo.DatDataCount;
                        totalCorrectDataCount += datInfo.CorrectDataCount;
                        totalErrorDataCount += datInfo.ErrorDataCount;
                        totalDatErrorDataCount += datInfo.DatErrorDataCount;
                        fileNames = fileNames + datInfo.FILE_NAME;
                    }

                    if (!string.IsNullOrEmpty(fileNames))
                    {
                        fileNames = string.Format("檔案：{0} 匯入結束。", fileNames);
                    }
                    else
                    {
                        //更新L_BATCH_LOG
                        InsertBatchLog(jobID, "S", "沒有檔案需要匯入");
                    }
                       

                    string resultMsg = string.Format(" JobOS06_AtDailyJob 批次執行成功：資料共 {1} 筆，匯入成功 {2} 筆，匯入失敗 {3} 筆，檢核失敗 {4} 筆，{0}", fileNames, totalDataCount, totalCorrectDataCount, totalErrorDataCount, totalDatErrorDataCount);
                    //寄成功信
                    SendMail(_MailTitle + "成功！" + resultMsg, resultMsg, "成功", this.StartTime);
                }
                else
                {
                    //更新L_BATCH_LOG
                    InsertBatchLog(JobHelper.strJobID, "F", errorMsg);

                    if (!string.IsNullOrEmpty(doingFileName))
                        doingFileName = string.Format("，失敗的檔案：{0}，請確認相關資料表是否已匯入順序在此檔案之前的檔案資料。", doingFileName);

                    //寄失敗信
                    SendMail(_MailTitle + "失敗！" + errorMsg, string.Format(" JobOS06_AtDailyJob 批次 發生錯誤：{0}{1}", errorMsg, doingFileName), "失敗", this.StartTime);
                }
            }
            // 20221004 調整將資料拆成多個檔案 By Kelton End


            JobHelper.SaveLog("*********** " + JobHelper.strJobID + " 取得CPMAST資料 批次 END ************** ", LogState.Info);
        }
    }

    #region Method
    /// <summary>
    /// 下載檔案
    /// </summary>
    /// <param name="jobID">jobID</param>
    /// <param name="atDailyJob"></param>
    /// <param name="date">執行日期</param>
    /// <param name="localPath">本地路徑</param>
    /// <param name="unZipPwd">解壓縮密碼</param>
    /// <param name="errorMsg">錯誤訊息</param>
    /// <returns></returns>
    private string DownloadFile(string jobID, OS06_AtDailyJob atDailyJob, string date, ref string localPath, ref string unZipPwd, ref string errorMsg)
    {
        string folderName = string.Empty;
        string ErrorChi = string.Empty;
        bool isDownloadOK = false;

        try
        {
            JobHelper.CreateFolderName(jobID, ref folderName);

            localPath = AppDomain.CurrentDomain.BaseDirectory + "FileDownload\\" + jobID + "\\" + folderName + "\\";

            string fileName = atDailyJob.DownloadFromFTP(date, localPath, "EXE", ref isDownloadOK, ref unZipPwd, ref errorMsg);

            if (!isDownloadOK && string.IsNullOrEmpty(errorMsg))
            {
                errorMsg = "檔案下載失敗！";
                return string.Empty;
            }

            return fileName;
        }
        catch (Exception ex)
        {
            Logging.Log(ex);
            JobHelper.SaveLog("檔案下載發生例外錯誤：" + ex.Message);
            errorMsg = "檔案下載發生例外錯誤";
            return string.Empty;
        }

    }

    /// <summary>
    /// 功能說明:下載多個檔案
    /// 作    者:Kelton
    /// 創建時間:2022/10/04
    /// 修改時間:
    /// </summary>
    /// <param name="jobID">jobID</param>
    /// <param name="atDailyJob"></param>
    /// <param name="date">執行日期</param>
    /// <param name="localPath">本地路徑</param>
    /// <param name="unZipPwd">解壓縮密碼</param>
    /// <param name="errorMsg">錯誤訊息</param>
    /// <returns></returns>
    private List<string> DownloadFiles(string jobID, OS06_AtDailyJob atDailyJob, string date, ref string localPath, ref string unZipPwd, ref string errorMsg, ref int fileOKDataCount,
                                       ref int settingfileCount, ref bool doCheck)
    {
        string folderName = string.Empty;
        string ErrorChi = string.Empty;
        bool isDownloadOK = false;

        try
        {
            JobHelper.CreateFolderName(jobID, ref folderName);

            localPath = AppDomain.CurrentDomain.BaseDirectory + "FileDownload\\" + jobID + "\\" + folderName + "\\";

            List<string> fileNames = atDailyJob.DownloadFromFTP_Multiple(date, localPath, "EXE", ref isDownloadOK, ref unZipPwd, ref errorMsg, ref fileOKDataCount, ref settingfileCount, ref doCheck);

            if (!isDownloadOK && string.IsNullOrEmpty(errorMsg))
            {
                errorMsg = "檔案下載失敗！";
                return null;
            }

            return fileNames;
        }
        catch (Exception ex)
        {
            Logging.Log(ex);
            JobHelper.SaveLog("檔案下載發生例外錯誤：" + ex.Message);
            errorMsg = "檔案下載發生例外錯誤";
            return null;
        }

    }

    /// <summary>
    /// 取得使用者資訊
    /// </summary>
    /// <param name="jobDataMap">批次參數</param>
    /// <returns></returns>
    private EntityAGENT_INFO GetAGENT_INFO(JobDataMap jobDataMap)
    {
        EntityAGENT_INFO eAgentInfo = new EntityAGENT_INFO();

        try
        {
            if (jobDataMap != null && jobDataMap.Count > 0)
            {
                eAgentInfo.agent_id = jobDataMap.GetString("userId");
                eAgentInfo.agent_pwd = jobDataMap.GetString("passWord");
                eAgentInfo.agent_id_racf = jobDataMap.GetString("racfId");
                eAgentInfo.agent_id_racf_pwd = jobDataMap.GetString("racfPassWord");
            }

            eAgentInfo.functionkey = FunctionKey;

            return eAgentInfo;
        }
        catch (Exception ex)
        {
            Logging.Log(ex);
            JobHelper.SaveLog("取得使用者資訊發生例外錯誤：" + ex.Message);
            return eAgentInfo;
        }

    }

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
            if (dtInfo == null || dtInfo.Rows.Count > 0) //20210531_Ares_Stanley-修正Job執行檢核條件
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

    /// <summary>
    /// 發送信件
    /// </summary>
    /// <param name="mailTitle"></param>
    /// <param name="mailBody"></param>
    /// <param name="status"></param>
    /// <param name="startTime"></param>
    public void SendMail(string mailTitle, string mailBody, string status, DateTime startTime)
    {
        try
        {
            string[] mailTos = strMail.Split(';');

            System.Collections.Specialized.NameValueCollection nvc = new System.Collections.Specialized.NameValueCollection();

            nvc["Name"] = strMail.Replace(';', ',');

            nvc["Title"] = mailTitle;

            nvc["StartTime"] = startTime.ToString();

            nvc["EndTime"] = DateTime.Now.ToString();

            if (this.makeErrorFile)
            {
                mailBody = mailBody + "\r\n 本次執行有錯誤資料，錯誤資料檔案路徑為：" + this.errorFilePath;
            }

            nvc["Message"] = mailBody.ToString().Trim();

            nvc["Status"] = status;

            MailService.MailSender(mailTos, 1, nvc, "");
        }
        catch (Exception ex)
        {
            Logging.Log(ex);
            JobHelper.SaveLog("發送信件發生例外錯誤：" + ex.Message);
        }
    }
    #endregion

    /// <summary>
    ///  紀錄 dat 匯入情況 by Kelton
    /// </summary>
    public class datInfo
    {
        private string _FILE_NAME;
        private int _DatDataCount;
        private int _CorrectDataCount;
        private int _ErrorDataCount;
        private int _DatErrorDataCount;
        public string FILE_NAME
        {
            get
            {
                return this._FILE_NAME;
            }
            set
            {
                this._FILE_NAME = value;
            }
        }

        public int DatDataCount
        {
            get
            {
                return this._DatDataCount;
            }
            set
            {
                this._DatDataCount = value;
            }
        }
        public int CorrectDataCount
        {
            get
            {
                return this._CorrectDataCount;
            }
            set
            {
                this._CorrectDataCount = value;
            }
        }
        public int ErrorDataCount
        {
            get
            {
                return this._ErrorDataCount;
            }
            set
            {
                this._ErrorDataCount = value;
            }
        }
        public int DatErrorDataCount
        {
            get
            {
                return this._DatErrorDataCount;
            }
            set
            {
                this._DatErrorDataCount = value;
            }
        }
    }
}