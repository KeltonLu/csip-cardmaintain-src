//******************************************************************
//*  作    者：Ares Stanley
//*  創建日期：2022/01/06
//*  功能說明：Stored Procedure轉批次，取得檔案並將資料匯入table[CPMAST]
//*  修改紀錄：
//* <author>            <time>            <TaskID>                <desc>
//* Ares Stanley 2022/02/10  20210058-CSIP作業服務平台現代化II    調整JobName，移除JOBLOG、JOBSTEPLOG
//* Ares Stanley 2022/03/15  20210058-CSIP作業服務平台現代化II    調整檢核失敗資料處理方式
//* Ares Stanley 2022/04/18  20210058-CSIP作業服務平台現代化II    調整LOG、解壓縮
//*******************************************************************
using System;
using System.Data;
using System.IO;
using System.Text;
using System.Threading;
using System.Data.SqlClient;
using Framework.Common.Logging;
using Framework.Common.Utility;
using Framework.Data;
using CSIPCommonModel.EntityLayer;
using System.Collections.Generic;
using CSIPCommonModel.BusinessRules;
using Framework.Data.OM.Collections;
using Framework.Data.OM;

/// <summary>
/// OS06_AtDailyJob 的摘要描述
/// </summary>
public class OS06_AtDailyJob
{
    private string jobID = string.Empty;
    private EntityAGENT_INFO eAgentInfo = new EntityAGENT_INFO();
    protected DateTime StartTime = DateTime.Now;
    protected JobHelper JobHelper = new JobHelper();

    public OS06_AtDailyJob(string jobID)
    {
        this.jobID = jobID;
        JobHelper.strJobID = jobID;
    }

    public OS06_AtDailyJob(string jobID, EntityAGENT_INFO eAgentInfo)
    {
        this.jobID = jobID;
        this.eAgentInfo = eAgentInfo;
        JobHelper.strJobID = jobID;
    }

    /// <summary>
    /// 從FTP下載檔案
    /// </summary>
    /// <param name="date">檔案日期</param>
    /// <param name="localPath">本地路徑</param>
    /// <param name="extension">副檔名</param>
    /// <param name="isDownload">是否下載成功</param>
    /// <returns></returns>
    public string DownloadFromFTP(string date, string localPath, string extension, ref bool isDownload, ref string unZipPwd, ref string errorMsg)
    {
        string fileName = string.Empty;
        try
        {
            DataTable tblFileInfo = new DataTable();

            if (!CSIPCardMaintain.BusinessRules.BRM_FileInfo.GetFTPFileInfo(this.jobID, ref tblFileInfo))
            {
                isDownload = false;
                errorMsg = "取得 JOB tblFileInfo 相關資料失敗，請確認 JobLog(Log\\JobOS06_AtDailyJob\\) 或 DefaultLog(Log\\Default\\)";
                return "";
            }

            if (tblFileInfo.Rows.Count <= 0)
            {
                isDownload = false;
                errorMsg = "JOB tblFileInfo 沒有資料";
                return "";
            }

            #region rerun mechanism
            if (!string.IsNullOrEmpty(tblFileInfo.Rows[0]["Parameter"].ToString()))
            {
                date = tblFileInfo.Rows[0]["Parameter"].ToString().Trim();
            }
            #endregion

            fileName = string.Format("OS06{0}.{1}", date, extension);
            unZipPwd = RedirectHelper.GetDecryptString(tblFileInfo.Rows[0]["ZipPwd"].ToString());//待確認
            string ftpPwd = RedirectHelper.GetDecryptString(tblFileInfo.Rows[0]["FtpPwd"].ToString());
            FTPFactory objFtp = new FTPFactory(tblFileInfo.Rows[0]["FtpIP"].ToString(), "", tblFileInfo.Rows[0]["FtpUserName"].ToString(), ftpPwd, "21", localPath, "Y");
            bool isNotFound = false;
            isDownload = objFtp.DownloadWithJob(tblFileInfo.Rows[0]["FtpPath"].ToString(), fileName, localPath, fileName, ref isNotFound, this.jobID);

            if (isDownload)
            {
                JobHelper.SaveLog(fileName + " FTP 取檔成功", LogState.Info);
            }
            else
            {
                if (isNotFound)
                {
                    //找不到檔案
                    JobHelper.SaveLog("[FAIL] 檔案: " + fileName + " FTP 取檔失敗，找不到檔案", LogState.Error);
                    errorMsg += "[FAIL] 檔案: " + fileName + " FTP 取檔失敗，找不到檔案";
                }
                else
                {
                    //下載失敗
                    JobHelper.SaveLog("[FAIL] 檔案: " + fileName + " FTP 取檔失敗，下載失敗", LogState.Error);
                    errorMsg += "[FAIL] 檔案: " + fileName + " FTP 取檔失敗，下載失敗";
                }
            }
        }
        catch (Exception ex)
        {
            Logging.Log(ex.Message);
            JobHelper.SaveLog("下載檔案時發生例外錯誤：" + ex.Message);
            errorMsg = "下載檔案時發生例外錯誤";
        }

        return fileName;
    }

    /// <summary>
    /// 功能說明:從FTP下載多個檔案
    /// 作    者:Kelton
    /// 創建時間:2022/10/04
    /// 修改時間:
    /// </summary>
    /// <param name="date">檔案日期</param>
    /// <param name="localPath">本地路徑</param>
    /// <param name="extension">副檔名</param>
    /// <param name="isDownload">是否下載成功</param>
    /// <returns></returns>
    public List<string> DownloadFromFTP_Multiple(string date, string localPath, string extension, ref bool isDownload, ref string unZipPwd, ref string errorMsg, ref int fileOKDataCount,
                                                 ref int settingfileCount, ref bool doCheck)
    {
        List<string> fileNames = new List<string>();
        try
        {
            DataTable tblFileInfo = new DataTable();

            if (!CSIPCardMaintain.BusinessRules.BRM_FileInfo.GetFTPFileInfo(this.jobID, ref tblFileInfo))
            {
                isDownload = false;
                errorMsg = "取得 JOB tblFileInfo 相關資料失敗，請確認 JobLog(Log\\JobOS06_AtDailyJob\\) 或 DefaultLog(Log\\Default\\)";
                return null;
            }

            if (tblFileInfo.Rows.Count <= 0)
            {
                isDownload = false;
                errorMsg = "JOB tblFileInfo 沒有資料";
                return null;
            }

            #region rerun mechanism
            if (!string.IsNullOrEmpty(tblFileInfo.Rows[0]["Parameter"].ToString()))
            {
                date = tblFileInfo.Rows[0]["Parameter"].ToString().Trim();
            }
            #endregion

            string remFileName = string.Format("OS06{0}.FILEOK", date);
            unZipPwd = RedirectHelper.GetDecryptString(tblFileInfo.Rows[0]["ZipPwd"].ToString());//待確認
            string ftpPwd = RedirectHelper.GetDecryptString(tblFileInfo.Rows[0]["FtpPwd"].ToString());
            FTPFactory objFtp = new FTPFactory(tblFileInfo.Rows[0]["FtpIP"].ToString(), "", tblFileInfo.Rows[0]["FtpUserName"].ToString(), ftpPwd, "21", localPath, "Y");

            string strReadLine = "";
            int dataCount = 0;  // FILEOK 檔案內容紀錄的筆數
            int fileCount = 0;  // FTP上實際應該要存在的檔案數
            int realfileCount = 0; // FTP上存在的檔案數
            int fileMaxDataCount = 0; //單一檔案最大筆數

            SqlCommand sqlcmd = new SqlCommand();
            sqlcmd.CommandType = CommandType.Text;
            sqlcmd.CommandText = string.Format("SELECT PROPERTY_CODE FROM M_PROPERTY_CODE WHERE FUNCTION_KEY = '{0}' AND PROPERTY_KEY = 'JobOS06_Check'", this.eAgentInfo.functionkey);
            DataSet ds = BRM_PROPERTY_CODE.SearchOnDataSet(sqlcmd, "Connection_CSIP");

            if (ds != null)
            {
                DataTable dt = ds.Tables[0];
                if (dt != null && dt.Rows.Count > 0)
                {
                    if (dt.Rows[0][0].ToString().ToUpper() == "N")
                    {
                        doCheck = false;
                    }
                }
                else
                {
                    JobHelper.SaveLog("[FAIL] M_PROPERTY_CODE 查無資料拆檔功能檢核開關設定資料", LogState.Error);
                }
            }
            else
            {
                JobHelper.SaveLog("[FAIL] 取得 M_PROPERTY_CODE 設定資料失敗", LogState.Error);
                errorMsg += "[FAIL] 取得 M_PROPERTY_CODE 設定資料失敗";
            }

            string[] tempfileNames = Array.FindAll(objFtp.GetFileList(tblFileInfo.Rows[0]["FtpPath"].ToString()), (v) => { return v.StartsWith(string.Format("OS06{0}", date)); });

            List<string> realFileNames = new List<string>();
            string[] splitstr;
            // 計算存在的檔案數量
            foreach (var name in tempfileNames)
            {
                splitstr = name.Split('.');
                if (splitstr[0] != string.Format("OS06{0}", date) && splitstr[1].ToUpper() == extension)
                {
                    realfileCount++;
                    realFileNames.Add(name);
                }
            }

            // 判斷是否要進行相關檢核
            if (doCheck)
            {
                // 取得設定資料          
                sqlcmd.CommandType = CommandType.Text;
                sqlcmd.CommandText = string.Format("SELECT PROPERTY_CODE, SEQUENCE FROM M_PROPERTY_CODE WHERE FUNCTION_KEY = '{0}' AND PROPERTY_KEY = '{1}'", this.eAgentInfo.functionkey, this.jobID);
                ds = BRM_PROPERTY_CODE.SearchOnDataSet(sqlcmd, "Connection_CSIP");

                if (ds != null)
                {
                    DataTable dt = ds.Tables[0];
                    // 取得可收的最大檔案數資料
                    if (dt != null && dt.Select("SEQUENCE = '1'").Length > 0)
                    {
                        settingfileCount = int.Parse(dt.Select("SEQUENCE = '1'")[0]["PROPERTY_CODE"].ToString());
                    }
                    else
                    {
                        JobHelper.SaveLog("[FAIL] M_PROPERTY_CODE 可收的最大檔案數設定資料有誤", LogState.Error);
                        errorMsg += "[FAIL] M_PROPERTY_CODE 可收的最大檔案數設定資料有誤";
                        return null;
                    }

                    // 取得單一檔案最大資料筆數資料
                    if (dt != null && dt.Select("SEQUENCE = '2'").Length > 0)
                    {
                        fileMaxDataCount = int.Parse(dt.Select("SEQUENCE = '2'")[0]["PROPERTY_CODE"].ToString());
                    }
                    else
                    {
                        JobHelper.SaveLog("[FAIL] M_PROPERTY_CODE 單一檔案最大資料筆數設定資料有誤", LogState.Error);
                        errorMsg += "[FAIL] M_PROPERTY_CODE 單一檔案最大資料筆數設定資料有誤";
                        return null;
                    }
                }
                else
                {
                    JobHelper.SaveLog("[FAIL] 取得 M_PROPERTY_CODE 設定資料失敗", LogState.Error);
                    errorMsg += "[FAIL] 取得 M_PROPERTY_CODE 設定資料失敗";
                    return null;
                }

                //檢查並下載 FILEOK 檔案
                JobHelper.SaveLog(string.Format("檔案 {0} 下載開始", remFileName), LogState.Info);
                bool returnFlag = objFtp.Download(tblFileInfo.Rows[0]["FtpPath"].ToString(), remFileName, localPath, remFileName);
                if (!returnFlag)
                {
                    errorMsg += "[FAIL] 檔案: " + remFileName + " FTP 取檔失敗，下載失敗";
                    return null;
                }

                StreamReader sr = new StreamReader(localPath + remFileName, System.Text.Encoding.Default);

                strReadLine = sr.ReadToEnd().Trim();
                if (strReadLine == string.Empty)
                {
                    JobHelper.SaveLog("[FAIL] 檔案: " + remFileName + " 內容為空", LogState.Error);
                    errorMsg += "[FAIL] 檔案: " + remFileName + " 內容為空";
                    return null;
                }

                // 取得總筆數
                dataCount = Convert.ToInt32(strReadLine);
                fileOKDataCount = dataCount;

                // 計算實際應有的檔案數
                if (dataCount % fileMaxDataCount == 0)
                {
                    fileCount = dataCount / fileMaxDataCount;
                }
                else
                {
                    fileCount = (dataCount / fileMaxDataCount) + 1;
                }

                // 因目前主機端只能拆成 3 個檔案，有可能最後一個檔案的筆數會大於單一檔案最大筆數的設定值
                // 故計算出來的檔案數量若大於可收的最大檔案數時，將計算的檔案數設成可收的最大檔案數
                if (settingfileCount < fileCount)
                {
                    fileCount = settingfileCount;
                }

                if (realfileCount != fileCount)
                {
                    JobHelper.SaveLog("[FAIL] 檔案數量有誤", LogState.Error);
                    errorMsg += "[FAIL] 檔案數量有誤";
                    return null;
                }

                if (settingfileCount < realfileCount)
                {
                    JobHelper.SaveLog("[FAIL] 實際存在的檔案數量大於 M_PROPERTY_CODE 設定的可收最大檔案數", LogState.Error);
                    errorMsg += "[FAIL] 實際存在的檔案數量大於 M_PROPERTY_CODE 設定的可收最大檔案數";
                    return null;
                }

                bool havefile = false;
                for (int i = 1; i <= fileCount; i++)
                {
                    havefile = false;
                    // 檢查檔案是否有跳號或缺少
                    foreach (var name in realFileNames)
                    {
                        splitstr = name.Split('.');
                        if (splitstr[0] == string.Format("OS06{0}{1}", date, i.ToString().PadLeft(2, '0')))
                        {
                            havefile = true;
                            break;
                        }
                    }
                    if (!havefile)
                    {
                        JobHelper.SaveLog("[FAIL] 檔案編號順序有跳號", LogState.Error);
                        errorMsg += "[FAIL] 檔案編號順序有跳號";
                        return null;
                    }
                }
            }

            bool isNotFound = false;
            // 下載檔案
            foreach (var fileName in realFileNames)
            {
                JobHelper.SaveLog(string.Format("檔案 {0} 下載開始", fileName), LogState.Info);
                isNotFound = false;

                isDownload = objFtp.DownloadWithJob(tblFileInfo.Rows[0]["FtpPath"].ToString(), fileName, localPath, fileName, ref isNotFound, this.jobID);

                if (isDownload)
                {
                    fileNames.Add(fileName);
                }
                else
                {
                    if (isNotFound)
                    {
                        //找不到檔案
                        JobHelper.SaveLog("[FAIL] 檔案: " + fileName + " FTP 取檔失敗，找不到檔案", LogState.Error);
                        errorMsg += "[FAIL] 檔案: " + fileName + " FTP 取檔失敗，找不到檔案";
                    }
                    else
                    {
                        //下載失敗
                        JobHelper.SaveLog("[FAIL] 檔案: " + fileName + " FTP 取檔失敗，下載失敗", LogState.Error);
                        errorMsg += "[FAIL] 檔案: " + fileName + " FTP 取檔失敗，下載失敗";
                    }
                    return null;
                }
            }
        }
        catch (Exception ex)
        {
            Logging.Log(ex.Message);
            JobHelper.SaveLog("下載檔案時發生例外錯誤：" + ex.Message);
            errorMsg = "下載檔案時發生例外錯誤";
        }

        return fileNames;
    }

    /// <summary>
    /// 解壓縮檔案
    /// </summary>
    /// <param name="jobID">jobID</param>
    /// <param name="filePath">檔案路徑</param>
    /// <param name="zipFileName">壓縮檔名稱</param>
    /// <param name="pwd">解壓縮密碼</param>
    /// <returns></returns>
    public bool DecompressFile(string jobID, string filePath, string zipFileName, string pwd, ref string errorMsg)
    {
        JobHelper JobHelper = new JobHelper();
        JobHelper.strJobID = jobID;
        bool unZipResult = false;

        try
        {
            int ZipCount = 0;
            unZipResult = JobHelper.ZipExeFile(filePath, filePath + "\\" + zipFileName, pwd, ref ZipCount);

        }
        catch (Exception ex)
        {
            Logging.Log(ex.Message);
            JobHelper.SaveLog("解壓縮時發生例外錯誤：" + ex.Message);
            errorMsg = "解壓縮時發生例外錯誤";
        }

        return unZipResult;
    }

    /// <summary>
    /// 解壓縮EXE檔
    /// </summary>
    /// <param name="destFolder"></param>
    /// <param name="srcZipFile"></param>
    /// <param name="password"></param>
    /// <returns></returns>
    public bool ZipExeFile(string destFolder, string srcZipFile, string password, ref string errorMsg)
    {
        try
        {
            string strDATFileName = string.Empty;
            string strExeFileName = srcZipFile.Substring(0, srcZipFile.Trim().Length - 4);


            strDATFileName = srcZipFile.Replace("EXE", "dat");

            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            startInfo.CreateNoWindow = true; //不顯示視窗
            startInfo.UseShellExecute = false; //不使用Shell
            startInfo.FileName = "cmd.exe"; //要啟動的應用程式名稱
            startInfo.WorkingDirectory = destFolder; //工作目錄
            if (!string.IsNullOrEmpty(password))
            {
                startInfo.Arguments = "/C " + destFolder + srcZipFile + " -y -g" + password;
            }
            else
            {
                startInfo.Arguments = "/C " + destFolder + srcZipFile + " -y ";
            }

            process.StartInfo = startInfo;
            process.Start();
            process.WaitForExit(1800000);
            process.Close();

            Thread.Sleep(30000);

            if (File.Exists(destFolder + strDATFileName))
            {
                return true;
            }
            else
            {
                errorMsg = string.Format("目的資料夾中沒有檔案 {0}", strDATFileName);
                return false;
            }
        }
        catch (Exception ex)
        {
            Logging.Log(ex.Message);
            JobHelper.SaveLog("解壓縮檔案時發生例外錯誤：" + ex.Message);
            errorMsg = "解壓縮檔案時發生例外錯誤";
            return false;
        }

    }

    /// <summary>
    /// 確認Import_Log是否有資料，若無則新增
    /// </summary>
    /// <param name="date">執行日期</param>
    /// <param name="fileName">檔案名稱</param>
    /// <param name="errorMsg">錯誤訊息</param>
    /// <returns></returns>
    public bool CheckImportLog(string date, string fileName, ref string errorMsg)
    {
        bool isSuccess = false;
        try
        {
            const string sqlSearchText = @"SELECT * FROM Import_Log WHERE	INDate = @INDate AND FileName = @FileName";
            const string sqlInsertText = @"INSERT Import_Log ( INDate, FileName ) VALUES ( @INDate,@FileName )";

            SqlCommand sqlcmd = new SqlCommand();
            sqlcmd.CommandText = sqlSearchText;
            sqlcmd.Parameters.Add(new SqlParameter("@INDate", date));
            sqlcmd.Parameters.Add(new SqlParameter("@FileName", fileName));

            DataSet ds = CSIPCardMaintain.BusinessRules.BRImprot_Log.SearchOnDataSet(sqlcmd);

            if (ds == null || ds.Tables.Count <= 0)
            {
                errorMsg = "查詢時發生錯誤，請確認 JobLog(Log\\JobOS06_AtDailyJob\\) 或 DefaultLog(Log\\Default\\)";
                return false;
            }

            DataTable dt = ds.Tables[0];

            if (dt.Rows.Count <= 0)
            {
                sqlcmd = new SqlCommand();
                sqlcmd.CommandText = sqlInsertText;
                sqlcmd.Parameters.Add(new SqlParameter("@INDate", date));
                sqlcmd.Parameters.Add(new SqlParameter("@FileName", fileName));
                isSuccess = CSIPCardMaintain.BusinessRules.BRImprot_Log.Add(sqlcmd);
            }
            else
            {
                return true;
            }

            if (!isSuccess)
            {
                errorMsg = "INSERT Import_Log 失敗！請確認 JobLog(Log\\JobOS06_AtDailyJob\\) 或 DefaultLog(Log\\Default\\)";
            }
        }
        catch (Exception ex)
        {
            Logging.Log(ex.Message);
            JobHelper.SaveLog("新增 Import_Log 資料時發生例外錯誤：" + ex.Message);
            errorMsg = "新增 Import_Log 資料時發生例外錯誤";
            isSuccess = false;
        }
        return isSuccess;
    }

    /// <summary>
    /// 功能說明:確認Import_Log是否有資料，若無則新增，檢查並回傳檔案是否已匯入
    /// 作    者:Kelton
    /// 創建時間:2022/10/20
    /// 修改時間: 
    /// </summary>
    /// <param name="date">執行日期</param>
    /// <param name="fileName">檔案名稱</param>
    /// <param name="errorMsg">錯誤訊息</param>
    /// <returns></returns>
    public bool CheckImportLog_New(string date, string fileName, ref string errorMsg, ref bool imported)
    {
        bool isSuccess = false;
        try
        {
            const string sqlSearchText = @"SELECT * FROM Import_Log WHERE	INDate = @INDate AND FileName = @FileName";
            const string sqlInsertText = @"INSERT Import_Log ( INDate, FileName ) VALUES ( @INDate,@FileName )";

            SqlCommand sqlcmd = new SqlCommand();
            sqlcmd.CommandText = sqlSearchText;
            sqlcmd.Parameters.Add(new SqlParameter("@INDate", date));
            sqlcmd.Parameters.Add(new SqlParameter("@FileName", fileName));

            DataSet ds = CSIPCardMaintain.BusinessRules.BRImprot_Log.SearchOnDataSet(sqlcmd);

            if (ds == null || ds.Tables.Count <= 0)
            {
                errorMsg = "查詢時發生錯誤，請確認 JobLog(Log\\JobOS06_AtDailyJob\\) 或 DefaultLog(Log\\Default\\)";
                return false;
            }

            DataTable dt = ds.Tables[0];

            if (dt.Rows.Count <= 0)
            {
                sqlcmd = new SqlCommand();
                sqlcmd.CommandText = sqlInsertText;
                sqlcmd.Parameters.Add(new SqlParameter("@INDate", date));
                sqlcmd.Parameters.Add(new SqlParameter("@FileName", fileName));
                isSuccess = CSIPCardMaintain.BusinessRules.BRImprot_Log.Add(sqlcmd);
            }
            else
            {
                // 檢查 Log 資料是否完整，用來判斷檔案是否已匯入過
                if (dt != null && !string.IsNullOrWhiteSpace(dt.Rows[0]["RecordNums"].ToString()) &&
                    !string.IsNullOrWhiteSpace(dt.Rows[0]["Active_Status"].ToString()) && !string.IsNullOrWhiteSpace(dt.Rows[0]["ErrorNums"].ToString()))
                {
                    imported = true;
                }
                return true;
            }

            if (!isSuccess)
            {
                errorMsg = "INSERT Import_Log 失敗！請確認 JobLog(Log\\JobOS06_AtDailyJob\\) 或 DefaultLog(Log\\Default\\)";
            }
        }
        catch (Exception ex)
        {
            Logging.Log(ex.Message);
            JobHelper.SaveLog("新增 Import_Log 資料時發生例外錯誤：" + ex.Message);
            errorMsg = "新增 Import_Log 資料時發生例外錯誤";
            isSuccess = false;
        }
        return isSuccess;
    }

    /// <summary>
    /// TruncateTable
    /// </summary>
    /// <param name="tableName">Table名稱</param>
    /// <param name="errorMsg">錯誤訊息</param>
    /// <returns></returns>
    public bool TruncateTable(string tableName, ref string errorMsg)
    {
        bool result = false;

        try
        {
            string sqlTruncateText = @"TRUNCATE TABLE @TableName";
            string connection = UtilHelper.GetConnectionStrings("Connection_System");
            SqlConnection sqlconn = new SqlConnection(connection);
            SqlCommand sqlcmd = new SqlCommand(sqlTruncateText.Replace("@TableName", tableName), sqlconn);
            sqlcmd.Parameters.Add(new SqlParameter("@TableName", tableName));
            sqlconn.Open();
            result = (sqlcmd.ExecuteNonQuery() == -1);
            sqlconn.Close();
        }
        catch (Exception ex)
        {
            Logging.Log(ex.Message);
            JobHelper.SaveLog("Truncate 時發生例外錯誤：" + ex.Message);
            errorMsg = "Truncate 時發生例外錯誤";
            result = false;
        }

        if (!result && string.IsNullOrEmpty(errorMsg))
        {
            errorMsg = "Truncate 時發生錯誤，請確認 JobLog(Log\\JobOS06_AtDailyJob\\) 或 DefaultLog(Log\\Default\\)";
        }
        return result;
    }

    /// <summary>
    /// 檢查CPMAST_TMP資料的日期是否正確，若否則複製CPMAST_TMP的錯誤資料到CPMAST_ERR並刪除CPMAST_TMP中的錯誤資料
    /// </summary>
    /// <param name="fileName">檔案名稱</param>
    /// <param name="errorDataCount">錯誤資料筆數</param>
    /// <param name="errorMsg">錯誤訊息</param>
    /// <returns></returns>
    public bool CheckCpmastTmp(string fileName, ref int errorDataCount, ref string errorMsg)
    {
        bool checkResult = false;
        try
        {
            const string sqlSearhText = @"SELECT COUNT ( * ) FROM	cpmast_tmp WHERE ( len( ltrim( MAINT_D ) ) = 8 	AND ( SUBSTRING ( MAINT_D, 3, 1 ) <> '/' OR SUBSTRING ( MAINT_D, 6, 1 ) <> '/' ) )";

            DataHelper dh = new DataHelper();
            DataSet ds = dh.ExecuteDataSet(sqlSearhText);
            if (ds == null || ds.Tables.Count <= 0)
            {
                errorMsg = "查詢 cpmast_tmp 發生錯誤，請確認 JobLog(Log\\JobOS06_AtDailyJob\\) 或 DefaultLog(Log\\Default\\)";
                return false;
            }

            DataTable dt = ds.Tables[0];
            int result = 0;
            if (!int.TryParse(ds.Tables[0].Rows[0][0].ToString(), out result))
            {
                errorMsg = "轉換錯誤筆數時發生錯誤";
                return false;
            }
            errorDataCount = result;

            if (errorDataCount <= 0)
            {
                return true;
            }
            else
            {
                //將錯誤資料另存至cpmast_Err，並從cpmast_tmp刪除錯誤資料
                if (!CopyDataToCpmastErrAndDeleteCpmastTmp(fileName, ref errorMsg))
                {
                    if (string.IsNullOrEmpty(errorMsg))
                        errorMsg = "將錯誤資料另存至cpmast_Err，並從cpmast_tmp刪除錯誤資料時發生錯誤";
                    return false;
                }
                else
                {
                    checkResult = true;
                }
            }
        }
        catch (Exception ex)
        {
            Logging.Log(ex.Message);
            JobHelper.SaveLog("將錯誤資料另存至cpmast_Err，並從cpmast_tmp刪除錯誤資料時發生例外錯誤：" + ex.Message);
            errorMsg = "將錯誤資料另存至cpmast_Err，並從cpmast_tmp刪除錯誤資料時發生例外錯誤";
            return false;
        }
        return checkResult;
    }

    /// <summary>
    /// 複製CPMAST_TMP的錯誤資料到CPMAST_ERR並刪除CPMAST_TMP中的錯誤資料
    /// </summary>
    /// <param name="fileName">檔案名稱</param>
    /// <param name="errorMsg">錯誤訊息</param>
    /// <returns></returns>
    public bool CopyDataToCpmastErrAndDeleteCpmastTmp(string fileName, ref string errorMsg)
    {
        bool copyAndDeleteStatus = false;
        try
        {
            const string sqlCopyText = @"
            INSERT INTO cpmast_Err ( TYPE, CUST_ID, CARD_TYPE, FLD_NAME, BEFOR_UPD, AFTER_UPD, LST_LIMIT, CUR_LIMIT, MAINT_D, MAINT_T, USER_ID, TER_ID, EXE_Name ) 
            SELECT
            TYPE,
            CUST_ID,
            CARD_TYPE,
            FLD_NAME,
            BEFOR_UPD,
            AFTER_UPD,
            0,
            0,
            ltrim( MAINT_D ),
            MAINT_T,
            isnull( USER_ID, '' ),
            '',
            SUBSTRING ( @fileName, 1, 12 ) 
            FROM
	            cpmast_tmp 
            WHERE
	            (
		            len( ltrim( MAINT_D ) ) = 8 
	            AND ( SUBSTRING ( MAINT_D, 3, 1 ) <> '/' OR SUBSTRING ( MAINT_D, 6, 1 ) <> '/' ) 
	            )";

            const string sqlDelText = @"DELETE FROM cpmast_tmp WHERE ( len( ltrim( MAINT_D ) ) = 8 	AND ( SUBSTRING ( MAINT_D, 3, 1 ) <> '/' OR SUBSTRING ( MAINT_D, 6, 1 ) <> '/' ) )";

            SqlCommand sqlcmd = new SqlCommand();
            sqlcmd.CommandText = sqlCopyText;
            sqlcmd.Parameters.Add(new SqlParameter("@fileName", fileName));
            //新增錯誤資料至CPMAST_ERR
            copyAndDeleteStatus = CSIPCardMaintain.BusinessRules.BRCPMAST_Err.Add(sqlcmd);

            if (!copyAndDeleteStatus)
            {
                errorMsg = "新增錯誤資料至 CPMAST_ERR 失敗，請確認 JobLog(Log\\JobOS06_AtDailyJob\\) 或 DefaultLog(Log\\Default\\)";
                return false;
            }

            DataHelper dh = new DataHelper();
            DataSet ds = new DataSet();
            //從CPMAST_TMP刪除錯誤資料
            ds = dh.ExecuteDataSet(sqlDelText);
            if (ds == null)
            {
                errorMsg = "刪除CPMAST_TMP錯誤資料失敗，請確認 JobLog(Log\\JobOS06_AtDailyJob\\) 或 DefaultLog(Log\\Default\\)";
                return false;
            }
            else
            {
                copyAndDeleteStatus = true;
            }
        }
        catch (Exception ex)
        {
            Logging.Log(ex.Message);
            JobHelper.SaveLog("複製 CPMAST_TMP 的錯誤資料到 CPMAST_ERR 並刪除 CPMAST_TMP 中的錯誤資料時發生例外錯誤：" + ex.Message);
            errorMsg = "複製 CPMAST_TMP 的錯誤資料到 CPMAST_ERR 並刪除 CPMAST_TMP 中的錯誤資料時發生例外錯誤";
            return false;
        }
        return copyAndDeleteStatus;
    }

    /// <summary>
    /// 轉換CPMAST_TMP日期格式
    /// </summary>
    /// <param name="correctDataCount">正確資料筆數</param>
    /// <param name="errorMsg">錯誤訊息</param>
    /// <returns></returns>
    public bool ConvertCpmastTmpData(ref int correctDataCount, ref string errorMsg)
    {
        try
        {
            const string sqlSearchText = @"select count(*) from  cpmast_tmp";
            const string sqlUpdateText1 = @"update cpmast_tmp Set MAINT_D = '20' + substring( MAINT_D,7,2) + substring( MAINT_D,1,2) + substring( MAINT_D,4,2) where len(ltrim(MAINT_D))=8";
            const string sqlUpdateText2 = @"update cpmast_tmp Set MAINT_D = ''  where len(ltrim(MAINT_D))<>8";
            const string sqlUpdateText3 = @"update cpmast_tmp Set USER_ID = ''  where USER_ID is null";
            DataHelper dh = new DataHelper();
            DataSet ds = new DataSet();

            ds = dh.ExecuteDataSet(sqlSearchText);
            if (ds == null || ds.Tables.Count <= 0)
            {
                errorMsg = "查詢 table[CPMAST_TMP]時發生錯誤，請確認 JobLog(Log\\JobOS06_AtDailyJob\\) 或 DefaultLog(Log\\Default\\)";
                return false;
            }

            int result = 0;
            if (!int.TryParse(ds.Tables[0].Rows[0][0].ToString(), out result))
            {
                errorMsg = "轉換正確筆數時發生錯誤";
                return false;
            }
            correctDataCount = result;

            SqlCommand sqlcmd = new SqlCommand();
            sqlcmd.CommandTimeout = int.Parse(UtilHelper.GetAppSettings("SqlCmdTimeoutMax")); //Timeout 調整為 webconfig參數 by Ares Stanley 20220621

            dh = new DataHelper();
            sqlcmd.CommandText = sqlUpdateText1;
            dh.ExecuteNonQuery(sqlcmd);

            dh = new DataHelper();
            sqlcmd.CommandText = sqlUpdateText2;
            dh.ExecuteNonQuery(sqlcmd);

            dh = new DataHelper();
            sqlcmd.CommandText = sqlUpdateText3;
            dh.ExecuteNonQuery(sqlcmd);

            return true;
        }
        catch (Exception ex)
        {
            Logging.Log(ex.Message);
            JobHelper.SaveLog("轉換 CPMAST_TMP 日期格式時發生例外錯誤：" + ex.Message);
            errorMsg = "轉換 CPMAST_TMP 日期格式時發生例外錯誤";
            return false;
        }
    }

    /// <summary>
    /// 將正確資料Insert到CPMAST
    /// </summary>
    /// <param name="fileName">檔案名稱</param>
    /// <param name="errorMsg">錯誤訊息</param>
    /// <returns></returns>
    public bool InsertCorrectDataToCpmast(string fileName, ref string errorMsg)
    {
        try
        {
            //查詢 cpmast_tmp有無資料
            const string sqlSearhText = @"SELECT COUNT ( * ) FROM	cpmast_tmp";

            DataHelper dh = new DataHelper();
            DataSet ds = dh.ExecuteDataSet(sqlSearhText);
            if (ds == null || ds.Tables.Count <= 0)
            {
                errorMsg = "查詢 cpmast_tmp 發生錯誤，請確認 JobLog(Log\\JobOS06_AtDailyJob\\) 或 DefaultLog(Log\\Default\\)";
                return false;
            }

            DataTable dt = ds.Tables[0];
            int result = 0;
            if (!int.TryParse(ds.Tables[0].Rows[0][0].ToString(), out result))
            {
                errorMsg = "轉換查詢筆數時發生錯誤";
                return false;
            }
            // 20221020 調整將資料拆成多個檔案 By Kelton start
            const string sqlInsertText = @"
            INSERT INTO cpmast ( TYPE, CUST_ID, CARD_TYPE, FLD_NAME, BEFOR_UPD, AFTER_UPD, LST_LIMIT, CUR_LIMIT, MAINT_D, MAINT_T, USER_ID, TER_ID, EXE_Name ) 
            SELECT
            TYPE,
            CUST_ID,
            CARD_TYPE,
            FLD_NAME,
            BEFOR_UPD,
            AFTER_UPD,
            0,
            0,
            MAINT_D,
            MAINT_T,
            USER_ID,
            '',
            SUBSTRING ( @fileName, 1, CHARINDEX('.', @fileName) - 1 ) 
            FROM
	            cpmast_tmp";
            // 20221004 調整將資料拆成多個檔案 By Kelton end

            SqlCommand sqlcmd = new SqlCommand();
            sqlcmd.CommandText = sqlInsertText;
            sqlcmd.Parameters.Add(new SqlParameter("@fileName", fileName));
            sqlcmd.CommandTimeout = 600; //Timeout 調整為 600 秒 by Ares Stanley 20220421

            if (result <= 0)
            {
                JobHelper.SaveLog("table[CPMAST_TMP]無資料！", LogState.Info);
                return true;
            }
            else
            {
                if (!CSIPCardMaintain.BusinessRules.BRCPMAST.Add(sqlcmd))
                {
                    errorMsg = "新增正確資料至table[CPMAST]時發生錯誤，請確認 JobLog(Log\\JobOS06_AtDailyJob\\) 或 DefaultLog(Log\\Default\\)";
                    return false;
                }
            }
            return true;
        }
        catch (Exception ex)
        {
            Logging.Log(ex.Message);
            JobHelper.SaveLog("將正確資料 Insert 到 CPMAST 時發生例外錯誤：" + ex.Message);
            errorMsg = "將正確資料 Insert 到 CPMAST 時發生例外錯誤";
            return false;
        }
    }

    /// <summary>
    /// 更新Import_Log
    /// 修改紀錄：調整錯誤數量為日期條件錯誤筆數+檢核失敗筆數 by Ares Stanley 20220325
    /// </summary>
    /// <param name="correctDataCount"></param>
    /// <param name="errorDataCount"></param>
    /// <param name="date"></param>
    /// <param name="fileName"></param>
    /// <param name="errorMsg"></param>
    /// <returns></returns>
    public bool UpdateImportLog(int correctDataCount, int errorDataCount, int datErrorDataCount, string date, ref string errorMsg)
    {
        try
        {
            if (correctDataCount > 0)
            {
                const string sqlUpdateText = @"
                UPDATE import_log 
                SET RecordNums = @COUNT,
                Active_Status = '匯檔成功',
                ErrorNums = @ErrCnt 
                WHERE
	                INDate = @INDate 
	                AND FileName = @fileName
                ";
                SqlCommand sqlcmd = new SqlCommand();
                sqlcmd.CommandText = sqlUpdateText;
                sqlcmd.Parameters.Add(new SqlParameter("@COUNT", correctDataCount));
                sqlcmd.Parameters.Add(new SqlParameter("@ErrCnt", errorDataCount + datErrorDataCount));
                sqlcmd.Parameters.Add(new SqlParameter("@INDate", date));
                sqlcmd.Parameters.Add(new SqlParameter("@fileName", string.Format("OS06{0}.EXE", date)));
                bool updateStatus = CSIPCardMaintain.BusinessRules.BRImprot_Log.Update(sqlcmd);
                if (!updateStatus)
                {
                    errorMsg += "　更新table[Import_log]失敗，請確認 JobLog(Log\\JobOS06_AtDailyJob\\) 或 DefaultLog(Log\\Default\\)";
                    return false;
                }
            }
            else
            {
                const string sqlUpdateText = @"
                UPDATE import_log 
                SET RecordNums = 0,
                Active_Status = '匯檔失敗',
                ErrorNums = @ErrCnt
                WHERE
	                INDate = @INDate 
	                AND FileName = @fileName
                ";
                SqlCommand sqlcmd = new SqlCommand();
                sqlcmd.CommandText = sqlUpdateText;
                sqlcmd.Parameters.Add(new SqlParameter("@ErrCnt", errorDataCount + datErrorDataCount));
                sqlcmd.Parameters.Add(new SqlParameter("@INDate", date));
                sqlcmd.Parameters.Add(new SqlParameter("@fileName", string.Format("OS06{0}.EXE", date)));
                bool updateStatus = CSIPCardMaintain.BusinessRules.BRImprot_Log.Update(sqlcmd);
                if (!updateStatus)
                {
                    errorMsg += "　更新table[Import_log]失敗，請確認 JobLog(Log\\JobOS06_AtDailyJob\\) 或 DefaultLog(Log\\Default\\)";
                    return false;
                }
            }
            return true;
        }
        catch (Exception ex)
        {
            Logging.Log(ex.Message);
            JobHelper.SaveLog("更新 table[Import_Log] 時發生例外錯誤：" + ex.Message);
            errorMsg += "　更新 table[Import_Log] 時發生例外錯誤";
            return false;
        }
    }

    /// <summary>
    /// 功能說明:更新Import_Log
    /// 作    者:Kelton
    /// 創建時間:2022/10/04
    /// 修改時間:
    /// </summary> 
    /// <param name="correctDataCount"></param>
    /// <param name="errorDataCount"></param>
    /// <param name="date"></param>
    /// <param name="fileName"></param>
    /// <param name="errorMsg"></param>
    /// <returns></returns>
    public bool UpdateImportLogByFileName(int correctDataCount, int errorDataCount, int datErrorDataCount, string date, string fileName, ref string errorMsg)
    {
        try
        {
            if (correctDataCount > 0)
            {
                const string sqlUpdateText = @"
                UPDATE import_log 
                SET RecordNums = @COUNT,
                Active_Status = '匯檔成功',
                ErrorNums = @ErrCnt 
                WHERE
	                INDate = @INDate 
	                AND FileName = @fileName
                ";
                SqlCommand sqlcmd = new SqlCommand();
                sqlcmd.CommandText = sqlUpdateText;
                sqlcmd.Parameters.Add(new SqlParameter("@COUNT", correctDataCount));
                sqlcmd.Parameters.Add(new SqlParameter("@ErrCnt", errorDataCount + datErrorDataCount));
                sqlcmd.Parameters.Add(new SqlParameter("@INDate", date));
                sqlcmd.Parameters.Add(new SqlParameter("@fileName", fileName));
                bool updateStatus = CSIPCardMaintain.BusinessRules.BRImprot_Log.Update(sqlcmd);
                if (!updateStatus)
                {
                    errorMsg += "　更新table[Import_log]失敗，請確認 JobLog(Log\\JobOS06_AtDailyJob\\) 或 DefaultLog(Log\\Default\\)";
                    return false;
                }
            }
            else
            {
                const string sqlUpdateText = @"
                UPDATE import_log 
                SET RecordNums = 0,
                Active_Status = '匯檔失敗',
                ErrorNums = @ErrCnt
                WHERE
	                INDate = @INDate 
	                AND FileName = @fileName
                ";
                SqlCommand sqlcmd = new SqlCommand();
                sqlcmd.CommandText = sqlUpdateText;
                sqlcmd.Parameters.Add(new SqlParameter("@ErrCnt", errorDataCount + datErrorDataCount));
                sqlcmd.Parameters.Add(new SqlParameter("@INDate", date));
                sqlcmd.Parameters.Add(new SqlParameter("@fileName", fileName));
                bool updateStatus = CSIPCardMaintain.BusinessRules.BRImprot_Log.Update(sqlcmd);
                if (!updateStatus)
                {
                    errorMsg += "　更新table[Import_log]失敗，請確認 JobLog(Log\\JobOS06_AtDailyJob\\) 或 DefaultLog(Log\\Default\\)";
                    return false;
                }
            }
            return true;
        }
        catch (Exception ex)
        {
            Logging.Log(ex.Message);
            JobHelper.SaveLog("更新 table[Import_Log] 時發生例外錯誤：" + ex.Message);
            errorMsg += "　更新 table[Import_Log] 時發生例外錯誤";
            return false;
        }
    }

    /// <summary>
    /// 取得DAT檔資料
    /// </summary>
    /// <param name="filePath"></param>
    /// <param name="fileName"></param>
    /// <param name="errorMsg"></param>
    /// <returns></returns>
    public DataTable GetMaintainData(string filePath, ref string fileName, ref string errorMsg, ref bool makeErrorFile, ref string errorFilePath, ref int errorDataCount, ref int datDataCount)
    {
        try
        {
            DirectoryInfo di = new DirectoryInfo(filePath);
            DataTable dt = new DataTable();

            foreach (FileInfo fi in di.GetFiles())
            {
                try
                {
                    if (fi.Extension == ".dat")
                    {
                        fileName = fi.Name;
                        dt = ValidateFileLength(fi, 159, ref errorMsg, ref makeErrorFile, ref errorFilePath, ref errorDataCount, ref datDataCount);
                    }
                }
                catch (Exception ex)
                {
                    Logging.Log(ex.Message);
                    JobHelper.SaveLog("取得DAT檔資料時發生例外錯誤：" + ex.Message);
                    errorMsg = "取得DAT檔資料時發生例外錯誤";
                    return null;
                }
            }
            return dt;
        }
        catch (Exception ex)
        {
            Logging.Log(ex.Message);
            JobHelper.SaveLog("取得DAT檔資料時發生例外錯誤：" + ex.Message);
            errorMsg = "取得DAT檔資料時發生例外錯誤";
            return null;
        }
    }

    /// <summary>
    /// 功能說明:取得指定DAT檔資料
    /// 作    者:Kelton
    /// 創建時間:2022/10/04
    /// 修改時間:
    /// </summary>
    /// <param name="filePath"></param>
    /// <param name="fileName"></param>
    /// <param name="errorMsg"></param>
    /// <returns></returns>
    public DataTable GetMaintainDataByFileName(string filePath, string fileName, ref string errorMsg, ref bool makeErrorFile, ref string errorFilePath, ref int errorDataCount, ref int datDataCount)
    {
        try
        {
            DirectoryInfo di = new DirectoryInfo(filePath);
            DataTable dt = new DataTable();

            foreach (FileInfo fi in di.GetFiles())
            {
                try
                {
                    if (fi.Extension == ".dat")
                    {
                        if (fi.Name.Substring(0, fi.Name.Length - 4) == fileName.Substring(0, fileName.Length - 4))
                        {
                            dt = ValidateFileLength(fi, 159, ref errorMsg, ref makeErrorFile, ref errorFilePath, ref errorDataCount, ref datDataCount);
                            break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logging.Log(ex.Message);
                    JobHelper.SaveLog("取得DAT檔資料時發生例外錯誤：" + ex.Message);
                    errorMsg = "取得DAT檔資料時發生例外錯誤";
                    return null;
                }
            }
            return dt;
        }
        catch (Exception ex)
        {
            Logging.Log(ex.Message);
            JobHelper.SaveLog("取得DAT檔資料時發生例外錯誤：" + ex.Message);
            errorMsg = "取得DAT檔資料時發生例外錯誤";
            return null;
        }
    }

    /// <summary>
    /// 取得DAT的Table Header
    /// </summary>
    /// <returns></returns>
    private DataTable SetDatTableHeader()
    {
        DataTable result = new DataTable();

        result.Columns.Add("TYPE", typeof(System.String));
        result.Columns.Add("COMPANY", typeof(System.String));
        result.Columns.Add("CARD_TYPE", typeof(System.String));
        result.Columns.Add("CUST_ID", typeof(System.String));
        result.Columns.Add("FILLER1", typeof(System.String));
        result.Columns.Add("TC_CODE", typeof(System.String));
        result.Columns.Add("FILLER2", typeof(System.String));
        result.Columns.Add("FLD_NAME", typeof(System.String));
        result.Columns.Add("FILLER3", typeof(System.String));
        result.Columns.Add("BEFOR_UPD", typeof(System.String));
        result.Columns.Add("FILLER4", typeof(System.String));
        result.Columns.Add("AFTER_UPD", typeof(System.String));
        result.Columns.Add("FILLER5", typeof(System.String));
        result.Columns.Add("LST_LIMIT", typeof(System.Int32));
        result.Columns.Add("CUR_LIMIT", typeof(System.Int32));
        result.Columns.Add("TER_ID", typeof(System.String));
        result.Columns.Add("FILLER6", typeof(System.String));
        result.Columns.Add("MAINT_D", typeof(System.String));
        result.Columns.Add("FILLER7", typeof(System.String));
        result.Columns.Add("MAINT_T", typeof(System.String));
        result.Columns.Add("USER_ID", typeof(System.String));

        return result;
    }

    /// <summary>
    /// 從bytes取得資料
    /// </summary>
    /// <param name="bytes">資料</param>
    /// <param name="startPoint">開始</param>
    /// <param name="length">長度</param>
    /// <returns></returns>
    public string NewString(byte[] bytes, int startPoint, int length)
    {
        string result = "";
        try
        {
            char[] chars = Encoding.Default.GetChars(bytes, startPoint, length);

            foreach (char chr in chars)
            {
                result = result + chr;
            }

            return result;
        }
        catch (Exception ex)
        {
            Logging.Log(ex.Message);
            JobHelper.SaveLog("從bytes取得資料發生例外錯誤：" + ex.Message);
            return result;
        }
    }

    /// <summary>
    /// 檢查每列資料長度並寫入Datatable
    /// 修改紀錄：調整Log紀錄 by Ares Stanley 20220310
    /// </summary>
    /// <param name="file">檔案</param>
    /// <param name="filerightlength">正確長度</param>
    /// <param name="errorMsg">錯誤訊息</param>
    /// <returns></returns>
    private DataTable ValidateFileLength(FileInfo file, int filerightlength, ref string errorMsg, ref bool makeErrorFile, ref string errorFilePath, ref int errorDataCount, ref int datDataCount)
    {
        int intcount = 0;
        int intcounterror = 0;
        bool isDatOK = true;
        string fileRow = "";
        int fileline = 0;

        DataTable result = SetDatTableHeader();
        StreamReader streamReader = new StreamReader(file.FullName, System.Text.Encoding.Default);
        StringBuilder errorFileRow = new StringBuilder();
        try
        {
            while ((fileRow = streamReader.ReadLine()) != null)
            {

                isDatOK = true;

                fileline = fileline + 1;

                byte[] bytes = Encoding.Default.GetBytes(fileRow);


                if (bytes.Length != filerightlength)
                {
                    //長度錯誤
                    intcounterror = intcounterror + 1;
                    isDatOK = false;
                    errorFileRow.AppendLine(fileRow);
                }
                else
                {
                    //長度正確
                    intcount += 1;

                    DataRow dtRow = null;
                    dtRow = result.NewRow();

                    dtRow["TYPE"] = NewString(bytes, 0, 1).Trim();
                    dtRow["COMPANY"] = NewString(bytes, 1, 3).Trim();
                    dtRow["CARD_TYPE"] = NewString(bytes, 4, 3).Trim();
                    dtRow["CUST_ID"] = NewString(bytes, 7, 16).Trim();
                    dtRow["FILLER1"] = NewString(bytes, 23, 1).Trim();
                    dtRow["TC_CODE"] = NewString(bytes, 24, 3).Trim();
                    dtRow["FILLER2"] = NewString(bytes, 27, 1).Trim();
                    dtRow["FLD_NAME"] = NewString(bytes, 28, 20).Trim();
                    dtRow["FILLER3"] = NewString(bytes, 48, 1).Trim();
                    dtRow["BEFOR_UPD"] = NewString(bytes, 49, 30).Trim();
                    dtRow["FILLER4"] = NewString(bytes, 79, 1).Trim();
                    dtRow["AFTER_UPD"] = NewString(bytes, 80, 30).Trim();
                    dtRow["FILLER5"] = NewString(bytes, 110, 12).Trim();
                    dtRow["LST_LIMIT"] = 0;
                    dtRow["CUR_LIMIT"] = 0;
                    dtRow["TER_ID"] = "";
                    dtRow["FILLER6"] = "";
                    dtRow["MAINT_D"] = NewString(bytes, 122, 8).Trim();
                    dtRow["FILLER7"] = NewString(bytes, 130, 1).Trim();
                    dtRow["MAINT_T"] = NewString(bytes, 131, 8).Trim();
                    dtRow["USER_ID"] = NewString(bytes, 139, 8).Trim();

                    result.Rows.Add(dtRow);
                    isDatOK = true;
                }

                if (isDatOK == false)
                {
                    JobHelper.SaveLog("檔案" + file.FullName + "在第" + fileline + "筆發生長度不正確，" + "正確長度需為" + filerightlength, LogState.Info);
                }
            }

            if (intcounterror > 0)
            {
                //將錯誤資料寫入檔案 OSyyyyMMdd.dat_ERROR.TXT
                File.WriteAllText(file.Directory.FullName + "\\" + file.Name + "_ERROR.TXT", errorFileRow.ToString(), Encoding.Default);
                makeErrorFile = true;
                errorFilePath = file.Directory.FullName + "\\" + file.Name + "_ERROR.TXT";
                JobHelper.SaveLog("共有" + intcounterror + "筆長度不正確，錯誤資料檔路徑為：" + errorFilePath, LogState.Info);
            }

            datDataCount = fileline;
            errorDataCount = intcounterror;

            return result;
        }
        catch (Exception ex)
        {
            Logging.Log(ex.Message);
            JobHelper.SaveLog("檢核資料長度時發生例外錯誤：" + ex.Message);
            errorMsg = "檢核資料長度時發生例外錯誤";
            return null;
        }
        finally
        {
            streamReader.Dispose();//auto close
            //若有檢核失敗資料則 Log 層級為 Error By Ares Stanley 20220503
            if (intcounterror > 0)
            {
                //有檢核失敗資料
                JobHelper.SaveLog(string.Format("檢核結果：檢核總筆數：{0}筆，成功：{1}筆，失敗：{2}筆", intcount + intcounterror, intcount, intcounterror), LogState.Error);
            }
            else
            {
                //無檢核失敗資料
                JobHelper.SaveLog(string.Format("檢核結果：檢核總筆數：{0}筆，成功：{1}筆，失敗：{2}筆", intcount + intcounterror, intcount, intcounterror), LogState.Info);
            }

        }
    }

    /// <summary>
    /// 將資料匯入CPMAST_TMP
    /// </summary>
    /// <param name="tableName">Table名稱</param>
    /// <param name="sourceData">資料</param>
    /// <param name="errorMsg">錯誤訊息</param>
    /// <returns></returns>
    public bool InsertCpmastTmp(string tableName, DataTable sourceData, ref string errorMsg)
    {
        bool result = false;
        string connnection = UtilHelper.GetConnectionStrings("Connection_System");
        SqlConnection conn = new SqlConnection(connnection);
        SqlBulkCopy sbc = new SqlBulkCopy(connnection);
        sbc.DestinationTableName = tableName;
        try
        {
            conn.Open();
            sbc.BulkCopyTimeout = 600;
            sbc.ColumnMappings.Add("TYPE", "TYPE");
            sbc.ColumnMappings.Add("COMPANY", "COMPANY");
            sbc.ColumnMappings.Add("CARD_TYPE", "CARD_TYPE");
            sbc.ColumnMappings.Add("CUST_ID", "CUST_ID");
            sbc.ColumnMappings.Add("FILLER1", "FILLER1");
            sbc.ColumnMappings.Add("TC_CODE", "TC_CODE");
            sbc.ColumnMappings.Add("FILLER2", "FILLER2");
            sbc.ColumnMappings.Add("FLD_NAME", "FLD_NAME");
            sbc.ColumnMappings.Add("FILLER3", "FILLER3");
            sbc.ColumnMappings.Add("BEFOR_UPD", "BEFOR_UPD");
            sbc.ColumnMappings.Add("FILLER4", "FILLER4");
            sbc.ColumnMappings.Add("AFTER_UPD", "AFTER_UPD");
            sbc.ColumnMappings.Add("FILLER5", "FILLER5");
            sbc.ColumnMappings.Add("LST_LIMIT", "LST_LIMIT");
            sbc.ColumnMappings.Add("CUR_LIMIT", "CUR_LIMIT");
            sbc.ColumnMappings.Add("TER_ID", "TER_ID");
            sbc.ColumnMappings.Add("FILLER6", "FILLER6");
            sbc.ColumnMappings.Add("MAINT_D", "MAINT_D");
            sbc.ColumnMappings.Add("FILLER7", "FILLER7");
            sbc.ColumnMappings.Add("MAINT_T", "MAINT_T");
            sbc.ColumnMappings.Add("USER_ID", "USER_ID");


            sbc.WriteToServer(sourceData);

            result = true;
            return result;
        }
        catch (Exception ex)
        {
            Logging.Log(ex.Message);
            JobHelper.SaveLog("將資料匯入 CPMAST_TMP 時發生例外錯誤：" + ex.Message);
            errorMsg = "將資料匯入 CPMAST_TMP 時發生例外錯誤";
            return false;
        }
        finally
        {
            sbc.Close();
            conn.Close();
            conn.Dispose();
        }
    }
}