//******************************************************************
//* 作    者：
//* 功能說明：
//* 創建日期：
//* 修改紀錄：
//* <author>            <time>            <TaskID>                <desc>
//* Ares Luke          2020/11/19         20200031-CSIP EOS       調整取web.config加解密參數
//*******************************************************************

using System;
using System.Data;
using Framework.Common.Logging;
using Framework.Common.Message;
using Framework.Common.IO;
using System.IO;
using System.Collections;
using System.Text;
using ICSharpCode.SharpZipLib.Zip;
using ICSharpCode.SharpZipLib.Checksum;
using System.Data.SqlClient;
using CSIPCommonModel.BusinessRules;
using CSIPCommonModel.EntityLayer;
using Framework.Common.Utility;
using CSIPCardMaintain.BusinessRules;

/// <summary>
/// JobHelper 的摘要描述
/// </summary>
public class JobHelper
{
    protected string strFunctionKey = UtilHelper.GetAppSettings("FunctionKey").ToString();
    /// <summary>
    /// 緩衝區大小(2048)
    /// </summary>
    private const int BUFFER_SIZE = 2048;


    /// <summary>
    /// 專案代號:20200031-CSIP EOS
    /// 功能說明:白箱掃描修正-HardCode
    /// 作    者:Ares Rick
    /// 修改時間:2021/01/27
    /// </summary>
    /// <summary>
    /// 預設的壓縮密碼
    /// </summary>


    public String strJobID = "";
    /// <summary>
    /// 功能說明:查詢Job啟動狀態
    /// 作    者:Simba Liu
    /// 創建時間:2010/05/04
    /// 修改記錄:
    /// </summary>
    /// <param name="jobId"></param>
    /// <returns></returns>
    public string SerchJobStatus(string jobId)
    {
        string strJobStatus = string.Empty;
        string strMsgID = string.Empty;
        Com_AutoJob.SearchJobStatus(strFunctionKey, ref strJobStatus, jobId, ref strMsgID);
        return strJobStatus;
    }

    /// <summary>
    /// 功能說明:獲取Job執行狀態
    /// 作    者:Simba Liu
    /// 創建時間:2010/05/04
    /// 修改記錄:
    /// </summary>
    /// <param name="jobId"></param>
    /// <returns></returns>
    public string GetJobStatus(int SCount, int FCcunt)
    {
        string strJobStatus = string.Empty;
        if (SCount > 0 && FCcunt == 0)
        {
            strJobStatus = "S";
        }
        else
        {
            strJobStatus = "F";
        }
        return strJobStatus;
    }

    /// <summary>
    /// 功能說明:保存FTP登陸失敗信息至DB
    /// 作    者:Simba Liu
    /// 創建時間:2010/05/04
    /// 修改記錄:
    /// </summary>
    /// <param name="jobId"></param>
    /// <returns></returns>
    public void WriteLogToDB(string jobId)
    {
        EntityL_BATCH_LOG LBatchLog = new EntityL_BATCH_LOG();
        LBatchLog.FUNCTION_KEY = strFunctionKey;
        LBatchLog.JOB_ID = jobId;
        LBatchLog.START_TIME = DateTime.Now;
        LBatchLog.END_TIME = DateTime.Now;
        LBatchLog.STATUS = "F";
        LBatchLog.RETURN_MESSAGE = "FTP連接失敗";
        BRL_BATCH_LOG.Insert(LBatchLog.FUNCTION_KEY, LBatchLog.JOB_ID, Convert.ToDateTime(LBatchLog.START_TIME), LBatchLog.STATUS, LBatchLog.RETURN_MESSAGE);
    }

    /// <summary>
    /// 功能說明:保存匯入信息至DB
    /// 作    者:Simba Liu
    /// 創建時間:2010/05/04
    /// 修改記錄:
    /// </summary>
    /// <param name="jobId"></param>
    /// <returns></returns>
    public void WriteLogToDB(string jobId, DateTime StartTime, DateTime EndTime, string strStatus, string strReturnMsg)
    {
        EntityL_BATCH_LOG LBatchLog = new EntityL_BATCH_LOG();
        LBatchLog.FUNCTION_KEY = strFunctionKey;
        LBatchLog.JOB_ID = jobId;
        LBatchLog.START_TIME = StartTime;
        LBatchLog.END_TIME = EndTime;
        LBatchLog.STATUS = strStatus;
        LBatchLog.RETURN_MESSAGE = strReturnMsg;
        BRL_BATCH_LOG.Insert(LBatchLog.FUNCTION_KEY, LBatchLog.JOB_ID, Convert.ToDateTime(LBatchLog.START_TIME), LBatchLog.STATUS, LBatchLog.RETURN_MESSAGE);
    }

    /// <summary>
    /// 記錄匯入日志
    /// </summary>
    /// <param name="eLUploadDetail">匯入錯誤日志</param>
    /// <param name="intRow">錯誤行號</param>
    /// <param name="strMsg">錯誤信息</param>
    /// <returns>int</returns>
    public static void LogUpload(EntityL_UPLOAD_DETAIL eLUploadDetail, int intRow, string strMsg)
    {
        eLUploadDetail.FAIL_REC_NO = intRow.ToString();
        eLUploadDetail.FAIL_REASON = strMsg;

        BRL_UPLOAD_DETAIL.Add(eLUploadDetail, ref  strMsg);
    }

    /// <summary>
    /// 功能說明:發送Mail
    /// 作    者:Simba Liu
    /// 創建時間:2010/05/04
    /// 修改記錄:
    /// </summary>
    public bool SendMail(string strFrom, string[] strTo, string strSubject, string strBody)
    {
        try
        {
            MailService MailService = new MailService();
            MailService.MailSenderBasic(strTo, strFrom, strSubject, strBody, string.Empty);
            SaveLog("成功發送mail", LogState.Info);
            return true;
        }
        catch (Exception exp)
        {
            SaveLog("發送mail中發生以下異常:" + exp.Message);
            return false;
        }
    }

    /// <summary>
    /// 功能說明:發送Mail
    /// 作    者:Simba Liu
    /// 創建時間:2010/05/04
    /// 修改記錄:
    /// </summary>
    public bool SendMail(string[] strTo, string[] strCc, string strFrom, string strSubject, string strBody)
    {
        try
        {
            MailService MailService = new MailService();
            if (null != strCc || strCc.Length >= 0)
            {
                if (!string.IsNullOrEmpty(strCc[0]))
                {
                    MailService.MailSenderBasic(strTo, strCc, strFrom, strSubject, strBody);
                }
                else
                {
                    MailService.MailSenderBasic(strTo, strFrom, strSubject, strBody, string.Empty);
                }
            }
            else
            {
                MailService.MailSenderBasic(strTo, strFrom, strSubject, strBody, string.Empty);
            }
            SaveLog("成功發送mail", LogState.Info);
            return true;
        }
        catch (Exception exp)
        {
            SaveLog("發送mail中發生以下異常:" + exp.Message);
            return false;
        }
    }

    public void Write(string JobID, string strMsg, LogState state = LogState.Error)
    {
        Logging.Log(strMsg, JobID, state);
    }

#if false
    public void Write(string JobID, string strMsg, LogState state = LogState.Error, JobExecutionContext context = null)
    {
        //string strFilePath = ConfigurationManager.AppSettings["JobLogPath"].ToString();
        string strFilePath = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "Log\\";
        string DataName = DateTime.Now.ToString("yyyyMMdd") + "_" + JobID + ".TXT";
        string strDataPath = strFilePath + DataName;
        FileStream fs = new FileStream(strDataPath, FileMode.Append);
        StreamWriter sw = new StreamWriter(fs);
        //開始寫入
        sw.Write(DateTime.Now.ToString() + " " + strMsg + "\r\n");
        //清空緩衝區
        sw.Flush();
        //關閉流
        sw.Close();
        fs.Close();

        //if (context != null)
        //{
        //    string strMail = context.JobDetail.JobDataMap.GetString("mail").Trim();
        //    string strTitle = context.JobDetail.JobDataMap.GetString("title").Trim();
        //    if (!string.IsNullOrEmpty(strMail))
        //        BRM_FileInfo.SendFailMail(strMail, strTitle, strMsg, DateTime.Now);
        //}
    }
#endif


    public void WriteError(string JobID, string strMsg)
    {
        Logging.Log(strMsg, JobID, LogState.Error);
        // WriteError(JobID, strMsg, null);
    }

#if false
    public void WriteError(string JobID, string strMsg, JobExecutionContext context)
    {
        //string strFilePath = ConfigurationManager.AppSettings["JobLogPath"].ToString();
        string strFilePath = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "Log\\";
        string DataName = DateTime.Now.ToString("yyyyMMdd") + "_" + JobID + ".error";
        string strDataPath = strFilePath + DataName;
        FileStream fs = new FileStream(strDataPath, FileMode.Append);
        StreamWriter sw = new StreamWriter(fs);
        //開始寫入
        sw.Write(DateTime.Now.ToString() + " " + strMsg + "\r\n");
        //清空緩衝區
        sw.Flush();
        //關閉流
        sw.Close();
        fs.Close();

        //if (context != null)
        //{
        //    string strMail = context.JobDetail.JobDataMap.GetString("mail").Trim();
        //    string strTitle = context.JobDetail.JobDataMap.GetString("title").Trim();
        //    if (!string.IsNullOrEmpty(strMail))
        //        BRM_FileInfo.SendFailMail(strMail, strTitle, strMsg, DateTime.Now);
        //}
    }
#endif

    /// <summary>
    /// 功能說明:記錄系統日誌
    /// 作    者:Simba Liu
    /// 創建時間:2010/05/04
    /// 修改記錄:
    /// </summary>
    /// <param name="Message">Message</param>
    public void SaveLog(string Message, LogState state = LogState.Error)
    {
        string strErrMsg = Message + ",\r\n";
        Logging.Log(strErrMsg, strJobID, state, LogLayer.BusinessRule);
    }

    /// <summary>
    /// 功能說明:記錄異常消息
    /// 作    者:Simba Liu
    /// 創建時間:2010/05/04
    /// 修改記錄:
    /// </summary>
    /// <param name="exp">exp</param>
    public void SaveLog(Exception exp)
    {
        string strErrMsg = exp.Message + ",\r\n";
        strErrMsg = strErrMsg + "Source:" + exp.Source.ToString() + ",\r\n";
        strErrMsg = strErrMsg + "StackTrace:" + exp.StackTrace;
        Logging.Log(strErrMsg, strJobID, LogState.Error, LogLayer.BusinessRule);
    }

    /// <summary>
    /// 功能說明:查詢交換檔資料
    /// 作    者:Simba Liu
    /// 創建時間:2010/05/04
    /// 修改記錄:
    /// </summary>
    /// <param name="dtFileInfo"></param>
    /// <param name="jobId"></param>
    /// <returns></returns>
    public bool SearchFileInfo(ref  DataTable dtFileInfo, string jobId)
    {
        return BRM_FileInfo.selectFileInfo(ref dtFileInfo, jobId);
    }

    /// <summary>
    /// 功能說明:字串向右補空格到最大長度
    /// 作    者:Simba Liu
    /// 創建時間:2010/05/04
    /// 修改記錄:
    /// </summary>
    /// <param name="strValue">待處理的字符串</param>
    /// <param name="iCount">字符串總長度</param>
    /// <returns></returns>
    public static string SetStrngValue(string strValue, uint iCount)
    {
        if (iCount == 0)
        {
            return strValue;
        }

        StringBuilder strbTemp = new StringBuilder();
        if (string.IsNullOrEmpty(strValue))
        {
            for (int iStepTo = 0; iStepTo < iCount; iStepTo++)
            {
                strbTemp.Append(" ");
            }
            strValue = strbTemp.ToString();

            return strValue;
        }
        else
        {
            int ilen = 0;

            for (int i = 0; i < strValue.Length; i++)
            {
                char c = strValue[i];
                int ibyte_len = System.Text.Encoding.Default.GetByteCount(c.ToString());
                if (ibyte_len > 1) //判断是否为汉字或全脚符号 
                {
                    ilen += 2;
                }
                else
                {
                    ilen += 1;
                }
            }

            if (ilen >= iCount)
            {
                return strValue;
            }

            for (int iStepTo = ilen; iStepTo < iCount; iStepTo++)
            {
                strbTemp.Append(" ");
            }
            strValue += strbTemp.ToString();
            return strValue;
        }
    }

    /// <summary>
    /// 功能說明:根據數值產生固定的長度(太短補齊，太常截斷)
    /// 作    者:Max Lu
    /// 創建時間:2011/01/23
    /// 修改記錄:
    /// </summary>
    /// <param name="strValue">待處理的字符串</param>
    /// <param name="iCount">字符串總長度</param>
    /// <returns></returns>
    public static string SetStrngValue2(string strValue, uint iCount)
    {
        if (iCount == 0)
        {
            strValue = "";
            return strValue;
        }

        StringBuilder strbTemp = new StringBuilder();
        if (string.IsNullOrEmpty(strValue))
        {
            for (int iStepTo = 0; iStepTo < iCount; iStepTo++)
            {
                strbTemp.Append(" ");
            }
            strValue = strbTemp.ToString();

            return strValue;
        }
        else
        {
            int ilen = 0;

            for (int i = 0; i < strValue.Length; i++)
            {
                char c = strValue[i];
                int ibyte_len = System.Text.Encoding.Default.GetByteCount(c.ToString());
                if (ibyte_len > 1) //判断是否为汉字或全脚符号 
                {
                    ilen += 2;
                }
                else
                {
                    ilen += 1;
                }
            }

            if (ilen >= iCount)
            {
                strValue = strValue.Substring(0, Convert.ToInt32(iCount));
                return strValue;
            }

            for (int iStepTo = ilen; iStepTo < iCount; iStepTo++)
            {
                strbTemp.Append(" ");
            }
            strValue += strbTemp.ToString();
            return strValue;
        }
    }

    /// <summary>
    /// 功能說明:解壓文件
    /// 作    者:Simba Liu
    /// 創建時間:2010/05/04
    /// 修改記錄:
    /// </summary>
    /// <param name="destFolder">解壓后的文件夾路徑</param>
    /// <param name="srcZipFile">需要解壓的文件</param>
    /// <param name="password">解壓密碼</param>
    /// <returns>解壓的文件個數</returns>
    public bool ZipExeFile(string destFolder, string srcZipFile, string password, ref int ZipCount)
    {
        if (string.IsNullOrEmpty(destFolder) || string.IsNullOrEmpty(srcZipFile))
        {
            return false;
        }
        try
        {
            ZipCount = CompressToZip.Unzip(destFolder, srcZipFile, password);
            if (ZipCount > 0)
            {
                //*解壓成功刪除壓縮文檔
                File.Delete(srcZipFile);
                SaveLog("解壓成功", LogState.Info);
                return true;
            }
            else
            {
                this.SaveLog("解壓失敗");
                return false;
            }

        }
        catch (Exception exp)
        {
            SaveLog("解壓文件失敗，失敗原因:" + exp.Message);
            return false;
        }
    }


#region 壓縮
    /// <summary>
    /// 以最好壓縮模式壓縮文件Zip文件,返回壓縮后的文件路徑
    /// </summary>
    /// <param name="strFile">需要壓縮的文件路徑</param>
    /// <param name="strZip">(傳出)壓縮后的zip文件路徑</param>
    public static void ZipFile(string strFile, string strZip)
    {
        if (strFile[strFile.Length - 1] != Path.DirectorySeparatorChar)
            strFile += Path.DirectorySeparatorChar;
        CompressToZip.ZipFile(strFile, strZip);
    }

    /// <summary>
    /// 壓縮文件夾成zip壓縮檔
    /// </summary>
    /// <param name="destFolder">壓縮后zip壓縮檔的路徑</param>
    /// <param name="srcFolder">需要壓縮的文件夾路徑</param>
    /// <param name="password">指定的壓縮密碼</param>
    /// <param name="level">壓縮級別</param>
    /// 等級 0 是最低等級, 只做歸檔, 不壓縮 
    /// 等級 1 壓縮率低, 但速度很快 
    /// 等級 6 是默認的壓縮等級 
    /// 等級 9 的壓縮率最高, 但它耗時也多
    /// <returns>壓縮的筆數</returns>
    public static bool Zip(string destFolder, string srcFolder, string password, CompressToZip.CompressLevel level)
    {
        int count = 0;
        bool ZipFlg = false;
        try
        {
            count = CompressToZip.Zip(destFolder, srcFolder, password, level);
        }
        catch
        {
            return false;
        }
        if (count > 0)
        {
            ZipFlg = true;
        }
        return ZipFlg;
    }

    /// <summary>
    /// 壓縮多分文件成zip壓縮檔
    /// </summary>
    /// <param name="destFolder">壓縮后zip壓縮檔的路徑</param>
    /// <param name="srcFiles">需要壓縮的文件(string元組)</param>
    /// <param name="folderName">zip中的文件夾名</param>
    /// <param name="password">指定的壓縮密碼</param>
    /// <param name="level">壓縮級別</param>
    /// <returns>壓縮的筆數</returns>
    public static int Zip(string destFolder, string[] srcFiles, string folderName, string password, CompressToZip.CompressLevel level)
    {
        ZipOutputStream zipStream = null;
        FileStream streamWriter = null;
        int count = 0;

        try
        {
            //Use Crc32
            Crc32 crc32 = new Crc32();

            //Create Zip File
            zipStream = new ZipOutputStream(File.Create(destFolder));

            //Specify Level
            zipStream.SetLevel(Convert.ToInt32(level));

            //Specify Password
            if (password != null && password.Trim().Length > 0)
            {
                zipStream.Password = password;
            }

            //Foreach File
            foreach (string file in srcFiles)
            {
                //Check Whether the file exists
                if (!File.Exists(file))
                {
                    throw new FileNotFoundException(file);
                }
                //Read the file to stream
                //streamWriter = File.OpenRead(file);
                streamWriter = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

                byte[] buffer = new byte[streamWriter.Length];
                streamWriter.Read(buffer, 0, buffer.Length);
                streamWriter.Close();

                //Specify ZipEntry
                crc32.Reset();
                crc32.Update(buffer);
                ZipEntry zipEntry = new ZipEntry(Path.Combine(folderName, Path.GetFileName(file)));
                zipEntry.DateTime = DateTime.Now;
                zipEntry.Size = buffer.Length;
                zipEntry.Crc = crc32.Value;
                zipEntry.IsUnicodeText = true;
                //Put file info into zip stream
                zipStream.PutNextEntry(zipEntry);

                //Put file data into zip stream
                zipStream.Write(buffer, 0, buffer.Length);

                count++;
            }
        }
        catch (Exception ex)
        {
            string msg = ex.Message;
            throw;
        }
        finally
        {
            //Clear Resource
            if (streamWriter != null)
            {
                streamWriter.Close();
            }
            if (zipStream != null)
            {
                zipStream.Finish();
                zipStream.Close();
            }
        }

        return count;
    }
#endregion


    /// <summary>
    /// 功能說明:在local創建目錄
    /// 作    者:Simba Liu
    /// 創建時間:2010/05/04
    /// 修改記錄:
    /// </summary>
    /// <param name="strPath">本地目錄路徑</param>
    public bool CreateLocalFolder(string strPath)
    {
        if (string.IsNullOrEmpty(strPath))
        {
            return false;
        }
        try
        {
            if (!Directory.Exists(strPath))
            {
                Directory.CreateDirectory(strPath);
                this.SaveLog("在local創建目錄成功", LogState.Info);
                return true;
            }
            else
            {
                return false;
            }
        }
        catch (Exception exp)
        {
            this.SaveLog("在local創建目錄失敗，失敗原因:" + exp.Message);
            return false;
        }
    }

    /// <summary>
    /// 功能說明:判斷當前時間是上午還是下午
    /// 作    者:Simba Liu
    /// 創建時間:2010/05/04
    /// 修改記錄:
    /// </summary>
    public void IsAmOrPm(DateTime time, ref string strAmOrPm)
    {
        if (time.Hour >= 12)
        {
            strAmOrPm = "pm";
        }
        else
        {
            strAmOrPm = "am";
        }
    }

    /// <summary>
    /// 功能說明:在本地新建一個包含年月日時分秒/AM/PM的目錄名稱
    /// 作    者:Simba Liu
    /// 創建時間:2010/05/04
    /// 修改記錄:
    /// </summary>
    /// <param name="strAmOrPm"></param>
    /// <returns></returns>
    public void CreateFolderName(string jobId, ref string strFolderName)
    {
        if (string.IsNullOrEmpty(jobId))
        {
            return;
        }
        strFolderName = jobId + DateTime.Now.ToString("yyyyMMddHHmmss");
    }

    /// <summary>
    /// 功能說明:判斷文件格式是否為txt格式
    /// 作    者:Simba Liu
    /// 創建時間:2010/05/04
    /// 修改記錄:
    /// </summary>
    /// <param name="strTxt"></param>
    /// <returns></returns>
    public bool ValidateTxt(string strFileName)
    {
        if (string.IsNullOrEmpty(strFileName))
        {
            return false;
        }
        System.IO.FileInfo Files = new System.IO.FileInfo(strFileName);
        if (Files.Extension.ToLower().Equals(".txt"))
        {
            return true;
        }
        else
        {
            return false;
        }
    }

    /// <summary>
    /// 功能說明:判斷文件格式是否為kk3或kk6格式
    /// 作    者:Grezz
    /// 創建時間:20180110
    /// 修改記錄:
    /// </summary>
    /// <param name="strFileName"></param>
    /// <returns></returns>
    public int Validatekk(string strFileName)
    {
        if (string.IsNullOrEmpty(strFileName))
        {
            return 2;
        }
        System.IO.FileInfo Files = new System.IO.FileInfo(strFileName);
        if (Files.Extension.ToLower().Equals(".kk3"))
        {
            return 0;
        }
        else if (Files.Extension.ToLower().Equals(".kk6"))
        {
            return 1;
        }
        else
        {
            return 2;
        }
    }

    /// <summary>
    /// 功能說明:讀取指定文件的內容
    /// 作    者:Simba Liu
    /// 創建時間:2010/06/09
    /// 修改記錄:
    /// </summary>
    /// <param name="path">文件路徑</param>
    /// <returns>按行排列文本內容</returns>
    public static string[] Read(string path)
    {
        string text1;
        ArrayList list1 = new ArrayList();
        FileStream stream1 = new FileStream(path, FileMode.Open);
        StreamReader reader1 = new StreamReader(stream1, Encoding.Default);
        while ((text1 = reader1.ReadLine()) != null)
        {
            list1.Add(text1);
        }
        reader1.Close();
        stream1.Close();
        return (string[])list1.ToArray(typeof(string));
    }

}
