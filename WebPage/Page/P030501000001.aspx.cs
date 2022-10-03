//******************************************************************
//*  作    者：yangyu(rosicky)
//*  功能說明：匯入記錄查詢
//*  創建日期：2009/09/21
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
using CSIPCommonModel.EntityLayer;

public partial class Page_P030501000001 : PageBase
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Page.Title = BaseHelper.GetShowText("03_05010000_000");    
        
        if (!Page.IsPostBack)
        {
            jsBuilder.RegScript(this.Page, BaseHelper.ClientMsgShow(""));
            dpEndData.Text = DateTime.Now.ToString("yyyy/MM/dd");
            dpBeforeData.Text = DateTime.Now.AddDays(-30).ToString("yyyy/MM/dd");
            Show();

            ViewState["BeforeDate"] = DateTime.Now.AddDays(-30).ToString("yyyy/MM/dd");
            ViewState["EndDate"] = DateTime.Now.ToString("yyyy/MM/dd");

            if (RedirectHelper.GetDecryptString(this.Page, "BeforeDate")!=null && RedirectHelper.GetDecryptString(this.Page, "EndDate")!=null )
            {
                if (RedirectHelper.GetDecryptString(this.Page, "BeforeDate") != "" && RedirectHelper.GetDecryptString(this.Page, "BeforeDate") != "")
                {
                    dpBeforeData.Text = Convert.ToDateTime(RedirectHelper.GetDecryptString(this.Page, "BeforeDate")).ToString("yyyy/MM/dd");
                    dpEndData.Text = Convert.ToDateTime(RedirectHelper.GetDecryptString(this.Page, "EndDate")).ToString("yyyy/MM/dd");
                }
                else 
                {
                    dpBeforeData.Text ="";
                    dpEndData.Text = "";
                }               
                try
                {
                    ViewState["BeforeDate"] = RedirectHelper.GetDecryptString(this.Page, "BeforeDate");
                    ViewState["EndDate"] = RedirectHelper.GetDecryptString(this.Page, "EndDate");

                    BindGridView(ViewState["BeforeDate"].ToString(), ViewState["EndDate"].ToString());
                }
                catch
                {
                    //修改彈跳視窗與末端訊息一樣 by Ares Neal 2022/06/15
                    jsBuilder.RegScript(this.UpdatePanel1, BaseHelper.ClientMsgShow("03_05010000_001"));
                    jsBuilder.RegScript(this.Page, string.Format("AlertConfirm({{title:'{0}'}});", MessageHelper.GetMessage("03_05010000_001")));
                    return;
                }
            }
            else 
            {
                BindGridView(ViewState["BeforeDate"].ToString(),ViewState["EndDate"].ToString());
            }
        }
    }


    /// <summary>
    /// 顯示窗體文字
    /// </summary>
    private void Show()
    {
        grvInpotLog.Columns[0].HeaderText = BaseHelper.GetShowText("03_05010000_003");
        grvInpotLog.Columns[1].HeaderText = BaseHelper.GetShowText("03_05010000_001");
        grvInpotLog.Columns[2].HeaderText = BaseHelper.GetShowText("03_05010000_004");
        grvInpotLog.Columns[3].HeaderText = BaseHelper.GetShowText("03_05010000_005");
        grvInpotLog.Columns[4].HeaderText = BaseHelper.GetShowText("03_05010000_006");
        grvInpotLog.Columns[5].HeaderText = BaseHelper.GetShowText("03_05010000_007");
        grvInpotLog.PageSize = int.Parse(UtilHelper.GetAppSettings("PageSize"));
        gpList.PageSize = int.Parse(UtilHelper.GetAppSettings("PageSize"));
    }

    /// <summary>
    /// 綁定GridView數據源

    /// </summary>
    private void BindGridView(string strBeforeData, string strEndData)
    {
        EntitySet<EntityImport_Log> esInpotLog = BRImprot_Log.Search(GetFilterCondition( strBeforeData, strEndData), this.gpList.CurrentPageIndex, this.gpList.PageSize);
        string strMsgID = "";
        try
        {
            this.gpList.RecordCount = esInpotLog.TotalCount;
            this.grvInpotLog.DataSource = esInpotLog;
            this.grvInpotLog.DataBind();
            //20220616_Ares_Jack_區分訊息敘述
            if (esInpotLog.TotalCount > 0)
                strMsgID = "03_05010000_006";//有資料
            else
                strMsgID = "00_00000000_037";//無資料

            //修改彈跳視窗與末端訊息一樣 by Ares Neal 2022/06/15
            jsBuilder.RegScript(this.UpdatePanel1, BaseHelper.ClientMsgShow(strMsgID));
            jsBuilder.RegScript(this.Page, string.Format("AlertConfirm({{title:'{0}'}});", MessageHelper.GetMessage(strMsgID)));

        }
        catch
        {
            strMsgID = "03_05010000_007";
            //修改彈跳視窗與末端訊息一樣 by Ares Neal 2022/06/15
            jsBuilder.RegScript(this.UpdatePanel1, BaseHelper.ClientMsgShow(strMsgID));
            jsBuilder.RegScript(this.Page, string.Format("AlertConfirm({{title:'{0}'}});", MessageHelper.GetMessage(strMsgID)));
        }
    }

    /// <summary>
    ///分頁顯示
    /// </summary>
    protected void gpList_PageChanged(object src, PageChangedEventArgs e)
    {
        this.gpList.CurrentPageIndex = e.NewPageIndex;
        this.BindGridView(ViewState["BeforeDate"].ToString(), ViewState["EndDate"].ToString());
    }

    /// <summary>
    /// 得到查询的SQL语句
    /// </summary>
    /// <returns>SQL语句</returns>
    private string GetFilterCondition(string strBeforeData,string strEndData)
    {
        SqlHelper Sql = new SqlHelper();      

        //* 點擊“資料查詢”時，按角色ID進行查詢
        if (strBeforeData!= "" && strEndData != "")
        {
            Sql.AddCondition(EntityImport_Log.M_INDate, Operator.GreaterThanEqual, DataTypeUtils.String, strBeforeData.Replace("/",""));
            Sql.AddCondition(EntityImport_Log.M_INDate, Operator.LessThanEqual, DataTypeUtils.String, strEndData.Replace("/", "")); 
        }
        Sql.AddORCondition("FILENAME LIKE 'os06%'", "FILENAME LIKE 'ts06%'");
        Sql.AddOrderCondition(EntityImport_Log.M_INDate, ESortType.DESC);
        return Sql.GetFilterORCondition();

    }


    /// <summary>
    /// 查詢
    /// 修改紀錄：調整日期格式避免離開錯誤明細時錯誤 by Ares Stanley 20220325
    /// </summary>
    protected void btnOK_Click(object sender, EventArgs e)
    {
        //*檢核屬性描述必須輸入
        if (dpBeforeData.Text == "" && dpEndData.Text == "")
        {
            BindGridView("", "");
            ViewState["BeforeDate"] = "";
            ViewState["EndDate"] ="";
        }
        else if (dpBeforeData.Text != "" && dpEndData.Text != "")
        {
            try
            {
                DateTime dtmBeforeData = Convert.ToDateTime(dpBeforeData.Text.Trim());
                DateTime dtmEndData = Convert.ToDateTime(dpEndData.Text.Trim());
                if (BRReport.CheckDataTime(dtmBeforeData, dtmEndData))
                {
                    ViewState["BeforeDate"] = Convert.ToDateTime(dpBeforeData.Text.Trim()).ToString("yyyy/MM/dd");
                    ViewState["EndDate"] = Convert.ToDateTime(dpEndData.Text.Trim()).ToString("yyyy/MM/dd");
                    BindGridView(ViewState["BeforeDate"].ToString(), ViewState["EndDate"].ToString());
                }
                else
                {
                    //修改彈跳視窗與末端訊息一樣 by Ares Neal 2022/06/15
                    jsBuilder.RegScript(this.UpdatePanel1, BaseHelper.ClientMsgShow("03_05010000_002"));
                    jsBuilder.RegScript(this.Page, string.Format("AlertConfirm({{title:'{0}'}}); if (document.getElementById('dpEndData_foo')) document.getElementById('dpEndData_foo').focus();", MessageHelper.GetMessage("03_05010000_002")));
                    return;
                }
            }
            catch
            {
                //修改彈跳視窗與末端訊息一樣 by Ares Neal 2022/06/15
                jsBuilder.RegScript(this.UpdatePanel1, BaseHelper.ClientMsgShow("03_05010000_001"));
                jsBuilder.RegScript(this.Page, string.Format("AlertConfirm({{title:'{0}'}});", MessageHelper.GetMessage("03_05010000_001")));
                return;
            }
        }
        else
        {
            //修改彈跳視窗與末端訊息一樣 by Ares Neal 2022/06/15
            jsBuilder.RegScript(this.UpdatePanel1, BaseHelper.ClientMsgShow("03_05010000_000"));
            jsBuilder.RegScript(this.Page, string.Format("AlertConfirm({{title:'{0}'}});", MessageHelper.GetMessage("03_05010000_000")));
            return;
        }
    }

    /// <summary>
    /// 點選某列,查看明細
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    protected void grvInpotLog_SelectedIndexChanging(object sender, GridViewSelectEventArgs e)
    {
        string strFileName = this.grvInpotLog.Rows[e.NewSelectedIndex].Cells[2].Text;
        string strDate = this.grvInpotLog.Rows[e.NewSelectedIndex].Cells[1].Text;
        string strErrNum = this.grvInpotLog.Rows[e.NewSelectedIndex].Cells[5].Text;

        if (strErrNum != "0" && strErrNum != "")
        {
            Response.Redirect("~/Page/P030501000002.aspx?FileName=" + RedirectHelper.GetEncryptParam(this.grvInpotLog.Rows[e.NewSelectedIndex].Cells[2].Text) + "&Date=" + RedirectHelper.GetEncryptParam(this.grvInpotLog.Rows[e.NewSelectedIndex].Cells[1].Text) + "&BeforeDate=" + RedirectHelper.GetEncryptParam(ViewState["BeforeDate"].ToString()) + "&EndDate=" + RedirectHelper.GetEncryptParam(ViewState["EndDate"].ToString()));
        }
        else
        {
            //修改彈跳視窗與末端訊息一樣 by Ares Neal 2022/06/15
            jsBuilder.RegScript(this.UpdatePanel1, BaseHelper.ClientMsgShow("03_05010000_003"));
            jsBuilder.RegScript(this.Page, string.Format("AlertConfirm({{title:'{0}'}});", MessageHelper.GetMessage("03_05010000_003")));
            return;        
        }
    }

    /// <summary>
    /// 綁定時增加點選事件
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    protected void grvInpotLog_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        Label olabel;
        if (e.Row.RowType == DataControlRowType.DataRow)
        {

            e.Row.Attributes["style"] = "Cursor:pointer";
            e.Row.Attributes["onclick"] = ClientScript.GetPostBackClientHyperlink(this.grvInpotLog, "Select$" + e.Row.RowIndex);

            olabel = (Label)e.Row.Cells[0].FindControl("lblNo");
            olabel.Text = Convert.ToString((this.gpList.CurrentPageIndex - 1) * this.gpList.PageSize + this.grvInpotLog.Rows.Count + 1);
        }
    }

    protected void grvInpotLog_SelectedIndexChanged(object sender, EventArgs e)
    {
        //修改彈跳視窗與末端訊息一樣 by Ares Neal 2022/06/15
        jsBuilder.RegScript(this.UpdatePanel1, BaseHelper.ClientMsgShow("03_05010000_003"));
        jsBuilder.RegScript(this.Page, string.Format("AlertConfirm({{title:'{0}'}});", MessageHelper.GetMessage("03_05010000_003")));
    }

    /// <summary>
    /// 列印匯入紀錄
    /// 修改紀錄：調整日期格式 by Ares Stanley 20220401
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    protected void btnPrint_Click(object sender, EventArgs e)
    {
        try
        {
            string strBeforeData = ViewState["BeforeDate"].ToString().Replace("/", "");
            string strEndData = ViewState["EndDate"].ToString().Replace("/", "");
            string sqlCondition = GetFilterCondition(strBeforeData, strEndData);
            string strMsgID = string.Empty;
            EntityAGENT_INFO eAgentinfo = (EntityAGENT_INFO)Session["Agent"];
            string strAgentName = eAgentinfo.agent_name;//* 業務員名字

            string strServerPathFile = this.Server.MapPath(UtilHelper.GetAppSettings("ExportExcelFilePath").ToString());
            if (!BR_Excel_File.CreateExcelFile_05010000(strBeforeData, strEndData, sqlCondition, strAgentName, ref strServerPathFile, ref strMsgID))
            {
                //修改彈跳視窗與末端訊息一樣 by Ares Neal 2022/06/15
                jsBuilder.RegScript(this.UpdatePanel1, BaseHelper.ClientMsgShow(strMsgID));
                jsBuilder.RegScript(this.Page, string.Format("AlertConfirm({{title:'{0}'}});", MessageHelper.GetMessage(strMsgID)));
                return;
            }
            //* 將服務器端生成的文檔，下載到本地。
            string strYYYYMMDD = "000" + Function.MinGuoDate7length(DateTime.Now.ToString("yyyyMMdd"));
            strYYYYMMDD = strYYYYMMDD.Substring(strYYYYMMDD.Length - 8, 8);
            string strFileName = "匯入紀錄查詢" + strYYYYMMDD + ".xls";

            //* 顯示提示訊息：匯出到Excel文檔資料成功
            this.Session["ServerFile"] = strServerPathFile;
            this.Session["ClientFile"] = strFileName;
            string urlString = @"window.parent.postMessage({ func: 'ClientMsgShow', data: '" + MessageHelper.GetMessage("00_00000000_039") + "' }, '*');";
            urlString += @"location.href='DownLoadFile.aspx';";
            //修改彈跳視窗與末端訊息一樣 by Ares Neal 2022/06/15
            jsBuilder.RegScript(this.UpdatePanel1, urlString);
            jsBuilder.RegScript(this.Page, string.Format("AlertConfirm({{title:'{0}'}});", MessageHelper.GetMessage("00_00000000_039")));
        }
        catch(Exception ex)
        {
            Logging.Log(ex);
            //修改彈跳視窗與末端訊息一樣 by Ares Neal 2022/06/15
            jsBuilder.RegScript(this.UpdatePanel1, BaseHelper.ClientMsgShow("00_00000000_038"));
            jsBuilder.RegScript(this.Page, string.Format("AlertConfirm({{title:'{0}'}});", MessageHelper.GetMessage("00_00000000_038")));
        }
    }
}
