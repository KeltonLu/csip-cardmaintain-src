//******************************************************************
//*  作    者：偉林
//*  功能說明：卡人-調整固定額度查詢
//*  創建日期：2009/10/07
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

public partial class Page_P030101030102 : PageBase
{
    // JaJa暫時移除
    // private ReportDocument rptResult;

    protected void Page_Load(object sender, EventArgs e)
    {

        string strMsgID = "";
        try
        {
            Page.Title = BaseHelper.GetShowText("03_01010000_012");
            string strRptPeople = RedirectHelper.GetDecryptString(this.Page, "People");
            string strRptBeforeDate = RedirectHelper.GetDecryptString(this.Page, "BeforeDate");
            string strRptEndDate = RedirectHelper.GetDecryptString(this.Page, "EndDate");
            string strRptSEQ = RedirectHelper.GetDecryptString(this.Page, "SEQ");
            string strRptOld = RedirectHelper.GetDecryptString(this.Page, "Old");
            string strRptBeforeAmount = RedirectHelper.GetDecryptString(this.Page, "BeforeAmount");
            string strRptEndAmount = RedirectHelper.GetDecryptString(this.Page, "EndAmount");

            CSIPCommonModel.EntityLayer.EntityAGENT_INFO eAgentinfo = (CSIPCommonModel.EntityLayer.EntityAGENT_INFO)Session["Agent"];
            string strName = eAgentinfo.agent_name;
            string strAgentID = eAgentinfo.agent_id;            //* 業務員ID




            bool blnOld = false;
            if (strRptOld.IndexOf("Old") > -1)
            {
                blnOld = true;
            }

            bool blnHadRecord = false;                          //* 是否有資料

            if (!IsPostBack)
            {
                //* 第一次進入頁面需要查詢
                if (BRReport.Report01010301_BySP(strRptPeople, strRptBeforeAmount, strRptEndAmount, strRptBeforeDate, strRptEndDate, strRptSEQ, strMsgID, strName, blnOld, strAgentID, ref blnHadRecord))
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

                rptResult.DataDefinition.FormulaFields["Title1"].Text = "'卡人ID'";
                rptResult.DataDefinition.FormulaFields["Nam"].Text = "'" + strName + "'";
                rptResult.DataDefinition.FormulaFields["Title"].Text = "'調整固定額度'";
                rptResult.DataDefinition.FormulaFields["Con"].Text = "'維護員 : " + strRptPeople + "'";
                rptResult.DataDefinition.FormulaFields["Con1"].Text = "'額    度 : " + strRptBeforeAmount + " ~ " + strRptEndAmount + "'";
                rptResult.DataDefinition.FormulaFields["Con2"].Text = "'維護日期 : " + strRptBeforeDate + " ~ " + strRptEndDate + "'";
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
