//******************************************************************
//*  作    者：余洋
//*  功能說明：卡人維護記錄查詢
//*  創建日期：2009/09/28
//*  修改記錄：
//*<author>            <time>            <TaskID>                <desc>
//*******************************************************************
using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using CSIPCardMaintain.BusinessRules;
using CSIPCardMaintain.EntityLayer;
using Framework.Data.OM.Collections;
using Framework.Data.OM;
using Framework.WebControls;
using Framework.Common.Utility;
using Framework.Common.Message;
using Framework.Common.JavaScript;
using CSIPCommonModel.BaseItem;
// JaJa暫時移除
// using CrystalDecisions.CrystalReports.Engine;
using System.Data.Odbc;
// JaJa暫時移除
// using CrystalDecisions.Shared;
using System.IO;

public partial class Page_P030101010002 : PageBase
{
    // JaJa暫時移除
    // private ReportDocument rptResult;

    public static DateTime dtStartTime;
    protected void Page_Load(object sender, EventArgs e)
    {
        dtStartTime = DateTime.Now;
        WriteLog("BEGIN Load");
        try
        {
            Page.Title = BaseHelper.GetShowText("03_01010000_000");
            string strRptID = RedirectHelper.GetDecryptString(this.Page, "ID");                 //* 卡人ID
            string strRptPeople = RedirectHelper.GetDecryptString(this.Page, "People");         //* 維護員
            string strRptBeforeDate = RedirectHelper.GetDecryptString(this.Page, "BeforeDate"); //* 時間起
            string strRptEndDate = RedirectHelper.GetDecryptString(this.Page, "EndDate");       //* 時間迄
            string strRptSEQ = RedirectHelper.GetDecryptString(this.Page, "SEQ");               //* 排序欄位
            string strRptOld = RedirectHelper.GetDecryptString(this.Page, "Old");               //* 一年半以前資料
            CSIPCommonModel.EntityLayer.EntityAGENT_INFO eAgentinfo = (CSIPCommonModel.EntityLayer.EntityAGENT_INFO)Session["Agent"];
            string strAgentName = eAgentinfo.agent_name;        //* 業務員名字
            string strAgentID = eAgentinfo.agent_id;            //* 業務員ID
            bool blnOld = false;
            if (strRptOld.Trim().IndexOf("Old") > -1)
            {
                blnOld = true;
            }
            bool blnHadRecord = false;                          //* 是否有資料

            if (!IsPostBack)
            {
                WriteLog("BEGIN QUERY");
                //* 第一次進入頁面需要查詢
                if (BRReport.Report01010100_BySP(strRptID, strRptPeople, strRptBeforeDate, strRptEndDate, blnOld, strRptSEQ, strAgentID, ref blnHadRecord))
                {
                    if (!blnHadRecord)
                    {
                        //* 沒有資料
                        BaseHelper.GetScriptForWindowClose(this.Page);
                        return;
                    }
                }
                else
                {
                    //* 出錯關閉
                    BaseHelper.GetScriptForWindowErrorClose(this.Page);
                    return;
                }
                WriteLog("END QUERY");
            }

            WriteLog("BEGIN BIND");

            // JaJa暫時移除
            /*rptResult = new ReportDocument();
            string strRPTPathFile = AppDomain.CurrentDomain.BaseDirectory + ConfigurationManager.AppSettings["ReportTemplate"] + "member1.rpt";
            rptResult.Load(@strRPTPathFile);

            rptResult.DataDefinition.FormulaFields["Title1"].Text   = "'卡人ID'";
            rptResult.DataDefinition.FormulaFields["Nam"].Text      = "'" + strAgentName + "'";
            rptResult.DataDefinition.FormulaFields["Title"].Text    = "'維護記錄查詢'";
            rptResult.DataDefinition.FormulaFields["Con"].Text      = "'卡人:" + strRptID + "'";
            rptResult.DataDefinition.FormulaFields["Con1"].Text     = "'維護日期 : " + strRptBeforeDate + " ~ " + strRptEndDate + "'";

            rptResult.RecordSelectionFormula = "{CPMAST.CSIPAgentID} = '" + strAgentID + "' ";


            if (!BRReport.SetDBLogonForReport(rptResult))
            {
                BaseHelper.GetScriptForWindowErrorClose(this.Page);
                return;
            }


            this.crvReport.ReportSource = rptResult;*/

            WriteLog("END BIND");
        }
        catch (Exception exp)
        {
            BRReport.SaveLog(exp);
            MessageHelper.ShowMessage(this.Page, "00_00000000_000");
            return;
        }

        WriteLog("END LOAD");
    }

    protected void Page_UnLoad(object sender, EventArgs e)
    {
        WriteLog("BEGIN UNLOAD");
        //建立完页面时，释放报表文档资源
        // JaJa暫時移除
        /*if (rptResult != null)
        {
            rptResult.Close();
            rptResult.Dispose();
            rptResult = null;
        }*/
        WriteLog("END UNLOAD");
    }

    public static void WriteLog(string strMsgContext)
    {
        DateTime dtNow = DateTime.Now;
        TimeSpan ts = ((TimeSpan)(dtNow - dtStartTime));
        string strMsg = dtNow.ToLongTimeString() + "( + " + ts.TotalSeconds.ToString() + ")" + strMsgContext;

        string strApplicationPath = AppDomain.CurrentDomain.BaseDirectory + "Log/PullLog.txt";
        System.IO.StreamWriter swFile = File.AppendText(strApplicationPath);
        swFile.WriteLine(strMsg);
        swFile.Flush();
        swFile.Close();
    }
}
