//******************************************************************
//*  作    者：
//*  功能說明：
//*  創建日期：
//*  修改記錄：
//*<author>            <time>            <TaskID>                <desc>
//* Ares Stanley    2022/02/14    20210058-CSIP作業服務平台現代化II    調整webconfig取參數方式
//*******************************************************************

using System;
using System.Data;
using System.Collections;
using System.Configuration;
using System.IO;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using Framework.Common.IO;
using Framework.Data.OM.Collections;
using CSIPCardMaintain.BusinessRules;
using CSIPCardMaintain.EntityLayer;
using Quartz;
using Quartz.Impl;
using CSIPCommonModel.EntityLayer;
using CSIPCommonModel.BusinessRules;
using Framework.Common.Utility;

/// <summary>
/// OS06JOB 的摘要描述
/// </summary>
public class OS06JOB : Quartz.IJob
{
    public OS06JOB()
    {
        //
        // TODO: 在此加入建構函式的程式碼
        //
    }

    #region IJob 成員
    const string strJobN = "OS06JOB";
    string strJOBNAME = UtilHelper.GetAppSettings("JOBNAME_" + strJobN).Trim();      //OS06JOB 名稱
    public void Execute(Quartz.JobExecutionContext context)
    {
        #region 變數設定
        string strJOBDATE = DateTime.Now.ToString("yyyyMMdd");      //JOB時間
        string strLocalPath = UtilHelper.GetAppSettings("SubTotalFilesPath_" + strJobN).Trim();    //本地目錄
        string strZipName = UtilHelper.GetAppSettings("ZipName_" + strJobN).Trim();    //Zip自解壓檔名稱
        string strZipFileName = "maintain.txt";     //Zip自解壓檔解壓后檔名
        string strZipDate = "";     //Zip檔日期

        //獲取Quartz JobDataMap傳值
        JobDataMap jdmValue = context.JobDetail.JobDataMap;
        string strFK = jdmValue.GetString("FunctionKey");  //Function Key
        string strJobID = jdmValue.GetString("JobID");   //Job ID
        string strAgentID = jdmValue.GetString("userId"); //Agent ID

        //Batch_Log
        DateTime dtST = DateTime.Now;   //JOB開始時間
        string strRMsg = "";
        //DateTime dtET = DateTime.Now;   //JOB結束時間

        #endregion

        if (!strLocalPath.EndsWith(@"\"))
        {
            strLocalPath = strLocalPath + @"\";
        }

        strLocalPath = strLocalPath + strJobN + @"\";

        //如果今天還在  執行中  或  執行成功了  .或 因為其他問題結束了.今天就不再執行
        if (BRJOBLOG.Select(strJOBNAME, strJOBDATE))
        {
            //已經執行過了或者正在執行.不再執行
            return;
        }

        //寫執行log
        BRJOBLOG.Insert(strJOBNAME);
        BRJOBSTEPLOG.Insert(strJOBNAME, "STEP0->開始執行" + strJOBNAME, "");

        #region 下載Zip檔
        //獲取Zip自解壓檔名稱和Zip檔日期
        if (!FormatZipName(ref strZipName, ref strZipDate))
        {
            BRJOBLOG.Update(strJOBNAME, "失敗", "Zip自解壓包名稱處理失敗");
            BRL_BATCH_LOG.Insert(strFK, strJobID, dtST, "F", "Zip自解壓包名稱處理失敗");
            return;
        }

        //下載Zip檔
        BRJOBSTEPLOG.Insert(strJOBNAME, "STEP1.1->下傳檔案開始", "");
        FTPFactory ftpf = new FTPFactory(strJobN);

        //刪除舊處理資料夾
        FileTools.DeleteFolder(strLocalPath);
        //新建處理資料夾
        FileTools.EnsurePath(strLocalPath);
        //下載檔案
        if (ftpf.Download(strZipName, strLocalPath, strZipName))
        {
            BRJOBSTEPLOG.Insert(strJOBNAME, "STEP2.1->下傳成功", strZipName);
        }
        else
        {
            BRJOBSTEPLOG.Insert(strJOBNAME, "STEP2.1->下傳失敗", "");
            BRJOBLOG.Update(strJOBNAME, "失敗", "下傳失敗");
            BRL_BATCH_LOG.Insert(strFK, strJobID, dtST, "F", "下傳失敗");
            return;
        }
        #endregion

        #region 解壓檔案
        //解壓檔案
        string strZipPass = "13572468";
        string strArg = " -g" + strZipPass + " -y " + strLocalPath;
        if (CompressToZip.ZipExeFile(strLocalPath, strZipName, strArg))
        {
            BRJOBSTEPLOG.Insert(strJOBNAME, "STEP2.2->解壓縮成功", strZipName);
        }
        else
        {
            BRJOBSTEPLOG.Insert(strJOBNAME, "STEP2.2->解壓縮失敗", "");
            BRJOBLOG.Update(strJOBNAME, "失敗", "解壓縮失敗");
            BRL_BATCH_LOG.Insert(strFK, strJobID, dtST, "F", "解壓縮失敗");
            return;
        }

        //修改解壓檔檔名
        string[] arrystrFilelist = FileTools.GetFileList(strLocalPath);
        for (int i = 0; i < arrystrFilelist.Length; i++)
        {
            if (arrystrFilelist[i] != strLocalPath + strZipName)
            {
                FileTools.MoveFile(arrystrFilelist[i], strLocalPath + strZipFileName);
                break;
            }
        }

        BRImprot_Log.Insert(strZipDate, strZipName);

        if (!File.Exists(strLocalPath + strZipFileName))
        {
            BRImprot_Log.Update(strZipDate, strZipName, 0, "", 0);
            BRJOBSTEPLOG.Insert(strJOBNAME, "STEP2.3->修改解壓檔檔名失敗", "");
            BRJOBLOG.Update(strJOBNAME, "失敗", "修改解壓檔檔名失敗");
            BRL_BATCH_LOG.Insert(strFK, strJobID, dtST, "F", "修改解壓檔檔名失敗");
            return;
        }

        #endregion

        #region 檢查數據
        DataTable dtblHead = new DataTable();//* 表頭
        DataTable dtblBody = new DataTable();//* 詳細資料
        DataTable dtEnd = new DataTable();
        int iMaxRowCount = 15000;//* 最大行數
        ArrayList alError = new ArrayList();//* 錯誤訊息
        string strMsg = "";//* 錯誤消息
        string strUploadID = "03001";
        int iAllNum = 0;    //總筆數
        int iFiledNum = 0;  //錯誤筆數
        EntitySet<EntityCPMAST> esetCPMAST = new EntitySet<EntityCPMAST>();
        EntitySet<EntityCPMAST_Err> esetCPMASTErr = new EntitySet<EntityCPMAST_Err>();

        //數據檢查
        dtblBody = BaseHelper.UploadCheck(strAgentID, strFK, strUploadID, DateTime.Now, strJobN,
                        strLocalPath + strZipFileName, iMaxRowCount, alError, ref strMsg, dtblHead, dtEnd);

        if ("" != strMsg)
        {
            BRJOBSTEPLOG.Insert(strJOBNAME, "STEP3.1->數據檢查作業失敗", "");
            BRJOBLOG.Update(strJOBNAME, "失敗", "數據檢查作業失敗");
            BRL_BATCH_LOG.Insert(strFK, strJobID, dtST, "F", "數據檢查作業失敗");
            return;
        }

        iAllNum = dtblBody.Rows.Count;
        if (0 == iAllNum)
        {
            BRJOBSTEPLOG.Insert(strJOBNAME, "STEP3.2->讀取數據失敗", "");
            BRJOBLOG.Update(strJOBNAME, "失敗", "讀取數據失敗");
            BRL_BATCH_LOG.Insert(strFK, strJobID, dtST, "F", "讀取數據失敗");
            return;
        }

        iFiledNum = alError.Count;

        //將失敗的資料記錄到失敗訊息中
        if (null != alError && alError.Count > 0)
        {
            foreach (string strErr in alError)
            {
                strRMsg = strErr + ",";
            }

            strRMsg = strRMsg.Trim(',');
        }
        #endregion

        #region 匯入資料
        for (int i = 0; i < dtblBody.Rows.Count; i++)
        {
            EntityCPMAST eCPMAST = new EntityCPMAST();
            EntityCPMAST_Err eCPMASTErr = new EntityCPMAST_Err();

            //檢查MAINT_D欄位日期格式是否正確
            string strMAINT_D = dtblBody.Rows[i]["COL9"].ToString().Trim();

            if (8 == strMAINT_D.Length && ("/" != strMAINT_D.Substring(2, 1) || "/" != strMAINT_D.Substring(5, 1)))
            {
                eCPMASTErr.TYPE = dtblBody.Rows[i]["COL1"].ToString().Trim();
                eCPMASTErr.CUST_ID = dtblBody.Rows[i]["COL2"].ToString().Trim();
                eCPMASTErr.CARD_TYPE = dtblBody.Rows[i]["COL3"].ToString().Trim();
                eCPMASTErr.FLD_NAME = dtblBody.Rows[i]["COL4"].ToString().Trim();
                eCPMASTErr.BEFOR_UPD = dtblBody.Rows[i]["COL5"].ToString().Trim();
                eCPMASTErr.AFTER_UPD = dtblBody.Rows[i]["COL6"].ToString().Trim();
                eCPMASTErr.LST_LIMIT = 0;
                eCPMASTErr.CUR_LIMIT = 0;
                eCPMASTErr.MAINT_D = strMAINT_D;
                eCPMASTErr.MAINT_T = dtblBody.Rows[i]["COL10"].ToString().Trim();
                eCPMASTErr.USER_ID = dtblBody.Rows[i]["COL11"].ToString().Trim();
                eCPMASTErr.TER_ID = "";
                eCPMASTErr.EXE_Name = strZipName;

                esetCPMASTErr.Add(eCPMASTErr);
                iFiledNum++;
            }
            else
            {
                eCPMAST.TYPE = dtblBody.Rows[i]["COL1"].ToString().Trim();
                eCPMAST.CUST_ID = dtblBody.Rows[i]["COL2"].ToString().Trim();
                eCPMAST.CARD_TYPE = dtblBody.Rows[i]["COL3"].ToString().Trim();
                eCPMAST.FLD_NAME = dtblBody.Rows[i]["COL4"].ToString().Trim();
                eCPMAST.BEFOR_UPD = dtblBody.Rows[i]["COL5"].ToString().Trim();
                eCPMAST.AFTER_UPD = dtblBody.Rows[i]["COL6"].ToString().Trim();
                eCPMAST.LST_LIMIT = 0;
                eCPMAST.CUR_LIMIT = 0;

                if (8 == strMAINT_D.Length)
                {
                    strMAINT_D = (2000 + int.Parse(strMAINT_D.Substring(6, 2))).ToString() + strMAINT_D.Substring(0, 2) + strMAINT_D.Substring(3, 2);
                }
                else
                {
                    strMAINT_D = "";
                }
                eCPMAST.MAINT_D = strMAINT_D;

                eCPMAST.MAINT_T = dtblBody.Rows[i]["COL10"].ToString().Trim();
                eCPMAST.USER_ID = dtblBody.Rows[i]["COL11"].ToString().Trim();
                eCPMAST.TER_ID = "";
                eCPMAST.EXE_Name = strZipName;

                esetCPMAST.Add(eCPMAST);
            }
        }

        //日期錯誤資料寫入CPMAST_Err表
        BRCPMAST_Err.Insert(esetCPMASTErr);

        string strMemo = "";
        string strStatus = "";
        if (0 != esetCPMAST.Count)
        {
            if (BRCPMAST.Insert(esetCPMAST))
            {
                strMemo = "匯入成功,成功" + esetCPMAST.Count.ToString() + "筆,失敗" + iFiledNum.ToString() + "筆";
                BRJOBSTEPLOG.Insert(strJOBNAME, "STEP4->匯入成功", "");
                BRJOBLOG.Update(strJOBNAME, "成功", strMemo);
                if (esetCPMAST.Count == iAllNum)
                {
                    strStatus = "S";
                    strRMsg = "總筆數：" + iAllNum.ToString() + "成功筆數：" + iAllNum.ToString() + "失敗筆數：" + iFiledNum.ToString() + "失敗訊息：" + strRMsg;
                }
                else
                {
                    strStatus = "P";
                    strRMsg = "總筆數：" + iAllNum.ToString() + "成功筆數：" + iAllNum.ToString() + "失敗筆數：" + iFiledNum.ToString() + "失敗訊息：" + strRMsg;
                }
                BRL_BATCH_LOG.Insert(strFK, strJobID, dtST, strStatus, strRMsg);
                BRImprot_Log.Update(strZipDate, strZipName, esetCPMAST.Count, "匯檔成功", iFiledNum);
            }
            else
            {
                BRJOBSTEPLOG.Insert(strJOBNAME, "STEP4->匯入失敗", "");
                BRJOBLOG.Update(strJOBNAME, "失敗", "匯入失敗");
                BRL_BATCH_LOG.Insert(strFK, strJobID, dtST, "F", "匯入失敗");
            }
        }
        else
        {
            strMemo = "匯入失敗,成功" + esetCPMAST.Count.ToString() + "筆,失敗" + iFiledNum.ToString() + "筆";
            BRJOBSTEPLOG.Insert(strJOBNAME, "STEP4->匯入結束", "");
            BRJOBLOG.Update(strJOBNAME, "結束", strMemo);
            strStatus = "F";
            strRMsg = "總筆數：" + iAllNum.ToString() + " 成功筆數：" + iAllNum.ToString() + " 失敗筆數：" + iFiledNum.ToString() + " 失敗訊息：" + strRMsg;
            BRL_BATCH_LOG.Insert(strFK, strJobID, dtST, strStatus, strRMsg);
            BRImprot_Log.Update(strZipDate, strZipName, 0, "", 0);
        }


        #endregion
    }

    #endregion

    /// <summary>
    /// 獲取zip文檔名稱和日期(如果名稱中[]中無值，則自動帶出上一天的日期，如有值，則使用已有的日期)
    /// </summary>
    /// <param name="strZipName">Zip檔名稱</param>
    /// <param name="strZipDate">Zip檔日期</param>
    /// <returns>是否成功</returns>
    private bool FormatZipName(ref string strZipName, ref string strZipDate)
    {
        try
        {
            strZipDate = strZipName.Substring(strZipName.IndexOf('[') + 1);
            strZipDate = strZipDate.Remove(strZipDate.IndexOf(']'));
            strZipDate = strZipDate.Trim();

            if (strZipDate.Equals(""))
            {
                strZipDate = DateTime.Now.AddDays(-1).ToString("yyyyMMdd");
                strZipName = strZipName.Replace("[]", DateTime.Now.AddDays(-1).ToString("MMdd"));
            }
            else
            {
                strZipDate = DateTime.Now.ToString("yyyy") + strZipDate;
                strZipName = strZipName.Replace("[", "").Replace("]", "");
            }

            return true;
        }
        catch (Exception ex)
        {
            BRJOBSTEPLOG.Insert(strJOBNAME, "失敗", "Zip自解壓包名稱處理失敗：" + ex.ToString());
            return false;
        }
    }
}
