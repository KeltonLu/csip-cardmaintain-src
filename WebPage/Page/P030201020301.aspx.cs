//******************************************************************
//*  作    者：yangyu(rosicky)
//*  功能說明：卡片-維護欄位統計表查詢
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
using Framework.Common.Logging;
using CSIPCommonModel.BaseItem;

public partial class Page_P030201020301 : PageBase
{
    //*排序欄位
    string strSEQ = "";
    //*查詢一年半以前資料
    string strOld = "";
    protected void Page_Load(object sender, EventArgs e)
    {
        Page.Title = BaseHelper.GetShowText("03_02010000_008");
        if (!Page.IsPostBack)
        {
            jsBuilder.RegScript(this.UpdatePanel1, BaseHelper.ClientMsgShow(""));

            txtFld.Focus();

            //* 設定Tittle
            ShowControlsText();
            this.gpList.Visible = false;
            this.gpList.RecordCount = 0;
            this.grvUserView.Visible = false;
        }
    }

    /// <summary>
    /// 列印
    /// 修改紀錄：報表改由NPOI產出 by Ares Stanley 20211130
    /// </summary>
    protected void btnOK_Click(object sender, EventArgs e)
    {
        try
        {
            if (dpBeforeDate.Text != "" && dpEndDate.Text != "")
            {
                DateTime dtmBeforeData = Convert.ToDateTime(dpBeforeDate.Text.Trim());
                DateTime dtmEndData = Convert.ToDateTime(dpEndDate.Text.Trim());
                if (!BRReport.CheckDataTime(dtmBeforeData, dtmEndData))
                {
                    jsBuilder.RegScript(this.UpdatePanel1, string.Format("AlertConfirm({{title:'{0}'}}); if (document.getElementById('dpEndDate_foo')) document.getElementById('dpEndDate_foo').focus();", MessageHelper.GetMessage("03_02010000_000")));
                    return;
                }
            }


            if (chkFld.Checked == true)
            {
                strSEQ = strSEQ + "FLD_NAME,";
            }

            if (chkDate.Checked == true)
            {
                strSEQ = strSEQ + "MAINT_D";
            }
            else
            {
                if (strSEQ != "")
                {
                    strSEQ = strSEQ.Substring(0, strSEQ.Length - 1);
                }
            }


            if (chkOld.Checked == true)
            {
                strOld = "Old";
            }

            string strMsgID = "";
            string strRptFld = this.txtFld.Text;
            string strRptBeforeDate = this.dpBeforeDate.Text.Replace("/", "");
            string strRptEndDate = this.dpEndDate.Text.Replace("/", "");
            string strRptSEQ = strSEQ;
            string strRptOld = strOld;
            CSIPCommonModel.EntityLayer.EntityAGENT_INFO eAgentinfo = (CSIPCommonModel.EntityLayer.EntityAGENT_INFO)Session["Agent"];
            string strName = eAgentinfo.agent_name;
            string strAgentID = eAgentinfo.agent_id;//* 業務員ID

            bool blnOld = false;
            if (strRptOld.IndexOf("Old") > -1)
            {
                blnOld = true;
            }

            bool blnHadRecord = false;//* 是否有資料

            //* 第一次進入頁面需要查詢
            if (BRReport.Report02010203_BySP(strRptFld, strRptBeforeDate, strRptEndDate, strRptSEQ, strMsgID, strName, blnOld, strAgentID, ref blnHadRecord, "P"))
            {
                if (!blnHadRecord)
                {
                    //* 沒有資料
                    jsBuilder.RegScript(this.UpdatePanel1, string.Format("AlertConfirm({{title:'{0}'}});", MessageHelper.GetMessage("00_00000000_037")));
                    return;
                }
            }
            else
            {
                //* 出錯關閉
                jsBuilder.RegScript(this.UpdatePanel1, string.Format("AlertConfirm({{title:'{0}'}});", MessageHelper.GetMessage("00_00000000_000")));
                return;
            }

            string strServerPathFile = this.Server.MapPath(UtilHelper.GetAppSettings("ExportExcelFilePath").ToString());
            if (!BR_Excel_File.CreateExcelFile_Report02010203(strName, strRptFld, strRptBeforeDate, strRptEndDate, strAgentID, ref strServerPathFile, ref strMsgID))
            {
                jsBuilder.RegScript(this.UpdatePanel1, string.Format("AlertConfirm({{title:'{0}'}});", MessageHelper.GetMessage("00_00000000_038")));
 				return;
            }
            //* 將服務器端生成的文檔，下載到本地。
            string strYYYYMMDD = "000" + Function.MinGuoDate7length(DateTime.Now.ToString("yyyyMMdd"));
            strYYYYMMDD = strYYYYMMDD.Substring(strYYYYMMDD.Length - 8, 8);
            string strFileName = "信用卡卡片資料查詢列印統計表或關係表查詢維護欄位統計表" + strYYYYMMDD + ".xls";

            //* 顯示提示訊息：匯出到Excel文檔資料成功
            this.Session["ServerFile"] = strServerPathFile;
            this.Session["ClientFile"] = strFileName;
            string urlString = @"window.parent.postMessage({ func: 'ClientMsgShow', data: '" + MessageHelper.GetMessage("00_00000000_039") + "' }, '*');";
            urlString += @"location.href='DownLoadFile.aspx';";
            jsBuilder.RegScript(this.Page, urlString);
        }
        catch(Exception ex)
        {
            Logging.Log(ex);
            MessageHelper.ShowMessage(this.UpdatePanel1, "03_02010000_000");
            return;
        }

    }

    /// <summary>
    /// 專案代號:20210058-CSIP 作業服務平台現代化II
    /// 功能說明:業務新增查詢功能
    /// 作    者:Ares Stanley
    /// 修改時間:2021/11/30
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    protected void btnSearch_Click(object sender, EventArgs e)
    {
        this.gpList.CurrentPageIndex = 1;
        BindGridView();
    }

    /// <summary>
    /// 專案代號:20210058-CSIP作業服務平台現代化II
    /// 功能說明:綁定畫面資料
    /// 作    者:Ares Stanley
    /// 修改時間:2021/11/30
    /// </summary>
    private void BindGridView()
    {
        try
        {
            if (dpBeforeDate.Text != "" && dpEndDate.Text != "")
            {
                DateTime dtmBeforeData = Convert.ToDateTime(dpBeforeDate.Text.Trim());
                DateTime dtmEndData = Convert.ToDateTime(dpEndDate.Text.Trim());
                if (!BRReport.CheckDataTime(dtmBeforeData, dtmEndData))
                {
                    jsBuilder.RegScript(this.UpdatePanel1, string.Format("AlertConfirm({{title:'{0}'}}); if (document.getElementById('dpEndDate_foo')) document.getElementById('dpEndDate_foo').focus();", MessageHelper.GetMessage("03_02010000_000")));
                    return;
                }
            }


            if (chkFld.Checked == true)
            {
                strSEQ = strSEQ + "FLD_NAME,";
            }

            if (chkDate.Checked == true)
            {
                strSEQ = strSEQ + "MAINT_D";
            }
            else
            {
                if (strSEQ != "")
                {
                    strSEQ = strSEQ.Substring(0, strSEQ.Length - 1);
                }
            }


            if (chkOld.Checked == true)
            {
                strOld = "Old";
            }

            string strMsgID = "";
            string strRptFld = this.txtFld.Text;
            string strRptBeforeDate = this.dpBeforeDate.Text.Replace("/", "");
            string strRptEndDate = this.dpEndDate.Text.Replace("/", "");
            string strRptSEQ = strSEQ;
            string strRptOld = strOld;
            CSIPCommonModel.EntityLayer.EntityAGENT_INFO eAgentinfo = (CSIPCommonModel.EntityLayer.EntityAGENT_INFO)Session["Agent"];
            string strName = eAgentinfo.agent_name;
            string strAgentID = eAgentinfo.agent_id;//* 業務員ID

            bool blnOld = false;
            if (strRptOld.IndexOf("Old") > -1)
            {
                blnOld = true;
            }

            bool blnHadRecord = false;//* 是否有資料

            //* 第一次進入頁面需要查詢
            if (BRReport.Report02010203_BySP(strRptFld, strRptBeforeDate, strRptEndDate, strRptSEQ, strMsgID, strName, blnOld, strAgentID, ref blnHadRecord))
            {
                if (!blnHadRecord)
                {
                    //* 沒有資料
                    jsBuilder.RegScript(this.UpdatePanel1, string.Format("AlertConfirm({{title:'{0}'}});", MessageHelper.GetMessage("00_00000000_037")));
                    return;
                }
            }
            else
            {
                //* 出錯關閉
                jsBuilder.RegScript(this.UpdatePanel1, string.Format("AlertConfirm({{title:'{0}'}});", MessageHelper.GetMessage("00_00000000_000")));
                return;
            }

            int totalCount = 0;
            DataTable dt = BR_Excel_File.getData_Comm(strAgentID, string.Format(BR_Excel_File.sqlComm_01010203, UtilHelper.GetAppSettings("DB_CP_DBF")), ref totalCount, "02010203", this.gpList.CurrentPageIndex);
            if (dt.Rows.Count > 0)
            {
                BR_Excel_File.removeBlank(ref dt);
                this.gpList.Visible = true;
                this.gpList.RecordCount = totalCount;
                this.grvUserView.Visible = true;
                this.grvUserView.DataSource = dt;
                this.grvUserView.DataBind();
                jsBuilder.RegScript(this.UpdatePanel1, BaseHelper.ClientMsgShow("03_01010000_006"));
            }
            else
            {
                this.gpList.RecordCount = 0;
                this.grvUserView.DataSource = null;
                this.grvUserView.DataBind();
                this.gpList.Visible = false;
                this.grvUserView.Visible = false;
                jsBuilder.RegScript(this.UpdatePanel1, BaseHelper.ClientMsgShow("03_01010000_007"));
                jsBuilder.RegScript(this.Page, string.Format("AlertConfirm({{title:'{0}'}});", MessageHelper.GetMessage("00_00000000_037")));
            }
        }
        catch(Exception ex)
        {
            Logging.Log(ex);
            MessageHelper.ShowMessage(this.UpdatePanel1, "03_02010000_000");
            return;
        }
    }

    /// <summary>
    /// 專案代號:20210058-CSIP 作業服務平台現代化II
    /// 功能說明:業務新增查詢切換頁需求功能
    /// 作    者:Ares Stanley
    /// 修改時間:2021/11/30
    /// </summary>
    protected void gpList_PageChanged(object src, Framework.WebControls.PageChangedEventArgs e)
    {
        gpList.CurrentPageIndex = e.NewPageIndex;
        BindGridView();
    }

    /// <summary>
    /// 專案代號:20210058-CSIP 作業服務平台現代化II
    /// 功能說明:業務新增查詢標頭需求功能
    /// 作    者:Ares Stanley
    /// 修改時間:2021/11/30
    /// </summary>
    protected void ShowControlsText()
    {
        {
            //* 設置查詢結果GridView的列頭標題
            this.grvUserView.Columns[0].HeaderText = BaseHelper.GetShowText("03_02010000_009");//維護欄位
            this.grvUserView.Columns[1].HeaderText = BaseHelper.GetShowText("03_02010000_029");//維護員
            this.grvUserView.Columns[2].HeaderText = BaseHelper.GetShowText("03_02010000_031");//小計

            //* 設置一頁顯示最大筆數
            this.gpList.PageSize = int.Parse(UtilHelper.GetAppSettings("PageSize"));
            this.grvUserView.PageSize = int.Parse(UtilHelper.GetAppSettings("PageSize"));
        }
    }


}
