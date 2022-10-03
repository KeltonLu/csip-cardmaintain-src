//******************************************************************
//*  作    者：偉林
//*  功能說明：卡人-自扣帳戶ID與卡人ID不同者查詢
//*  創建日期：2009/10/15
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
// JaJa暫時移除
// using CrystalDecisions.CrystalReports.Engine;
using System.Data.Odbc;
// JaJa暫時移除
// using CrystalDecisions.Shared;
using CSIPCommonModel.EntityLayer;
using Framework.Common.Utility;
using CSIPCardMaintain.BusinessRules;
using Framework.Common.JavaScript;
using Framework.Common.Message;

public partial class Page_P030101040202 : PageBase
{
    // JaJa暫時移除
    // private ReportDocument rptResult;

    protected void Page_Load(object sender, EventArgs e)
    {


        string strMsgID = "";
        try
        {

            Page.Title = BaseHelper.GetShowText("03_01010000_020");
            string strRptBeforeDate = RedirectHelper.GetDecryptString(this.Page, "BeforeDate");
            string strRptEndDate = RedirectHelper.GetDecryptString(this.Page, "EndDate");
            string strRptSEQ = RedirectHelper.GetDecryptString(this.Page, "SEQ");
            string strRptOld = RedirectHelper.GetDecryptString(this.Page, "Old");
            EntityAGENT_INFO eAgentinfo = (EntityAGENT_INFO)Session["Agent"];
            string strName = eAgentinfo.agent_name;
            string strAgentID = eAgentinfo.agent_id;            //* 業務員ID

            bool blnOld = false;
            if (strRptOld.Trim().IndexOf("Old") > -1)
            {
                blnOld = true;
            }

            bool blnHadRecord = false;                          //* 是否有資料

            if (!IsPostBack)
            {
                //* 第一次進入頁面需要查詢
                if (BRReport.Report01010402_BySP(strRptBeforeDate, strRptEndDate, strRptSEQ, strMsgID, strName, blnOld, strAgentID, ref blnHadRecord))
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
            }

            // JaJa暫時移除
            /*rptResult = new ReportDocument();
                string strRPTPathFile = AppDomain.CurrentDomain.BaseDirectory + ConfigurationManager.AppSettings["ReportTemplate"] + "member1.rpt";
                rptResult.Load(@strRPTPathFile);

                rptResult.DataDefinition.FormulaFields["Title"].Text = "'自扣ID與卡人ID不同者' ";
                rptResult.DataDefinition.FormulaFields["Con1"].Text = "' 維護欄位 : DIRECT DEBIT ID' ";
                rptResult.DataDefinition.FormulaFields["Con2"].Text = "' 維護日期 : " + strRptBeforeDate + " ~ " + strRptEndDate + "' ";
                rptResult.DataDefinition.FormulaFields["Nam"].Text = "' " + strName + "' ";
                rptResult.DataDefinition.FormulaFields["Title1"].Text = "'卡人ID ' ";
                rptResult.RecordSelectionFormula = "{CPMAST.CSIPAgentID} = '" + strAgentID + "' ";

                if (!BRReport.SetDBLogonForReport(rptResult))
                {
                    BaseHelper.GetScriptForWindowErrorClose(this.Page);
                    return;
                }


                this.crvReport.ReportSource = rptResult;*/
        }
        catch (Exception exp)
        {
            BRReport.SaveLog(exp);
            MessageHelper.ShowMessage(this.Page, "00_00000000_000");
            return;
        }

    }

    protected void Page_UnLoad(object sender, EventArgs e)
    {
        //建立完页面时，释放报表文档资源
        // JaJa暫時移除
        /*if (rptResult != null)
        {
            rptResult.Close();
            rptResult.Dispose();
        }*/
        this.Dispose();
        this.ClearChildState();
        System.GC.Collect(0);
    }

}
