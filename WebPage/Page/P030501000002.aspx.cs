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

public partial class Page_P030501000002 : PageBase
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!Page.IsPostBack)
        {
            jsBuilder.RegScript(this.UpdatePanel1, BaseHelper.ClientMsgShow(""));
            if (!string.IsNullOrEmpty(RedirectHelper.GetDecryptString(this.Page, "FileName")) && !string.IsNullOrEmpty(RedirectHelper.GetDecryptString(this.Page, "Date")))
            {
                ViewState["FileName"] = RedirectHelper.GetDecryptString(this.Page, "FileName");
                ViewState["Date"] = RedirectHelper.GetDecryptString(this.Page, "Date");
                ViewState["BeforeDate"] = RedirectHelper.GetDecryptString(this.Page, "BeforeDate");
                ViewState["EndDate"] = RedirectHelper.GetDecryptString(this.Page, "EndDate");

                Show();

                txtData.Text = ViewState["Date"].ToString();
                txtFileName.Text = ViewState["FileName"].ToString();
                txtData.Enabled = false;
                txtFileName.Enabled = false;

                if (ViewState["FileName"].ToString().Substring(0, 4).ToUpper() == "OS06")
                {
                    BindGridView(ViewState["FileName"].ToString());
                }
                else if (ViewState["FileName"].ToString().Substring(0, 4).ToUpper() == "TS06")
                {
                    BindGridView4Err(ViewState["FileName"].ToString());
                }
            }
        }
    }

    /// <summary>
    /// 顯示窗體文字
    /// </summary>
    private void Show()
    {
        grvCPMASTErr.Columns[0].HeaderText = BaseHelper.GetShowText("03_05010000_003");
        grvCPMASTErr.Columns[1].HeaderText = BaseHelper.GetShowText("03_05010000_011");
        grvCPMASTErr.Columns[2].HeaderText = BaseHelper.GetShowText("03_05010000_012");
        grvCPMASTErr.Columns[3].HeaderText = BaseHelper.GetShowText("03_05010000_013");
        grvCPMASTErr.Columns[4].HeaderText = BaseHelper.GetShowText("03_05010000_014");
        grvCPMASTErr.Columns[5].HeaderText = BaseHelper.GetShowText("03_05010000_015");
        grvCPMASTErr.Columns[6].HeaderText = BaseHelper.GetShowText("03_05010000_016");
        grvCPMASTErr.Columns[7].HeaderText = BaseHelper.GetShowText("03_05010000_017");

        grvCPMASTErr.PageSize = int.Parse(UtilHelper.GetAppSettings("PageSize"));
        gpList.PageSize = int.Parse(UtilHelper.GetAppSettings("PageSize"));
    }

    /// <summary>
    /// 綁定GridView數據源
    /// 修改紀錄：調整查詢條件避免查無資料 by Ares Stanley 20220325
    /// </summary>
    private void BindGridView(string strFileName)
    {
        try
        {
            strFileName = strFileName.Substring(0, 12);
            EntitySet<EntityCPMAST_Err> esCPMAST_Err = BRCPMAST_Err.Search(GetFilterCondition(strFileName), this.gpList.CurrentPageIndex, this.gpList.PageSize);
            this.gpList.RecordCount = esCPMAST_Err.TotalCount;
            this.grvCPMASTErr.DataSource = esCPMAST_Err;
            this.grvCPMASTErr.DataBind();
            string strMsgID = "03_05010000_004";
            jsBuilder.RegScript(UpdatePanel1, BaseHelper.ClientMsgShow(strMsgID));
        }
        catch
        {
            string strMsgID = "03_05010000_005";
            jsBuilder.RegScript(UpdatePanel1, BaseHelper.ClientMsgShow(strMsgID));        
        }
    }

    /// <summary>
    ///分頁顯示
    /// </summary>
    protected void gpList_PageChanged(object src, PageChangedEventArgs e)
    {
        this.gpList.CurrentPageIndex = e.NewPageIndex;
        if (ViewState["FileName"].ToString().Substring(0, 4).ToUpper() == "OS06")
        {
            BindGridView(ViewState["FileName"].ToString());
        }
        else if (ViewState["FileName"].ToString().Substring(0, 4).ToUpper() == "TS06")
        {
            BindGridView4Err(ViewState["FileName"].ToString());
        }
        else
        {
            Response.Redirect("~/Page/P030501000001.aspx?BeforeDate=" + RedirectHelper.GetEncryptParam(ViewState["BeforeDate"].ToString()) + "&EndDate=" + RedirectHelper.GetEncryptParam(ViewState["EndDate"].ToString()), false);
        }
    }

    /// <summary>
    /// 得到查询的SQL语句
    /// </summary>
    /// <returns>SQL语句</returns>
    private string GetFilterCondition(string strFileName)
    {
        SqlHelper Sql = new SqlHelper();
        Sql.AddCondition(EntityCPMAST_Err.M_EXE_Name, Operator.Equal, DataTypeUtils.String, strFileName);
        return Sql.GetFilterCondition();
    }

    /// <summary>
    /// 綁定GridView數據源
    /// 修改紀錄：調整查詢條件避免查無資料 by Ares Stanley 20220325
    /// </summary>
    private void BindGridView4Err(string strFileName)
    {
        try
        {
            strFileName = strFileName.Substring(0, 12);
            EntitySet<EntityCPMAST4_Err> esCPMAST_Err = BRCPMAST4_Err.Search(GetFilterCondition4Err(strFileName), this.gpList.CurrentPageIndex, this.gpList.PageSize);
            this.gpList.RecordCount = esCPMAST_Err.TotalCount;
            this.grvCPMASTErr.DataSource = esCPMAST_Err;
            this.grvCPMASTErr.DataBind();
            string strMsgID = "03_05010000_004";
            jsBuilder.RegScript(UpdatePanel1, BaseHelper.ClientMsgShow(strMsgID));
        }
        catch
        {
            string strMsgID = "03_05010000_005";
            jsBuilder.RegScript(UpdatePanel1, BaseHelper.ClientMsgShow(strMsgID));        
        }
    }

    /// <summary>
    /// 得到查询的SQL语句
    /// </summary>
    /// <returns>SQL语句</returns>
    private string GetFilterCondition4Err(string strFileName)
    {
        SqlHelper Sql = new SqlHelper();
        Sql.AddCondition(EntityCPMAST4_Err.M_EXE_Name, Operator.Equal, DataTypeUtils.String, strFileName);
        return Sql.GetFilterCondition();
    }

    protected void btnOK_Click(object sender, EventArgs e)
    {
        Response.Redirect("~/Page/P030501000001.aspx?BeforeDate=" + RedirectHelper.GetEncryptParam(ViewState["BeforeDate"].ToString()) + "&EndDate=" + RedirectHelper.GetEncryptParam(ViewState["EndDate"].ToString()), false);
    }

    protected void grvCPMASTErr_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        Label lblNo;
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            lblNo = (Label)e.Row.Cells[0].FindControl("lblNo");
            lblNo.Text = Convert.ToString((this.gpList.CurrentPageIndex - 1) * this.gpList.PageSize + this.grvCPMASTErr.Rows.Count + 1);
        }
    }

    protected void btnPrint_Click(object sender, EventArgs e)
    {
        try
        {
            string importDate = this.txtData.Text.Trim();
            string fileName = ViewState["FileName"].ToString().Substring(0, 4).ToUpper();
            string tableName = string.Empty;
            string strMsgID = string.Empty;
            EntityAGENT_INFO eAgentinfo = (EntityAGENT_INFO)Session["Agent"];
            string strAgentName = eAgentinfo.agent_name;//* 業務員名字
            string strServerPathFile = this.Server.MapPath(UtilHelper.GetAppSettings("ExportExcelFilePath").ToString());

            if (fileName == "OS06")
            {
                tableName = "CPMAST_Err";
            }
            else if (fileName == "TS06")
            {
                tableName = "CPMAST4_Err";
            }
            else
            {
                return;
            }

            if (!BR_Excel_File.CreateExcelFile_05010000Detail(fileName, tableName, importDate, strAgentName, ref strServerPathFile, ref strMsgID))
            {
                jsBuilder.RegScript(this.UpdatePanel1, string.Format("AlertConfirm({{title:'{0}'}});", MessageHelper.GetMessage(strMsgID)));
                return;
            }
            //* 將服務器端生成的文檔，下載到本地。
            string strYYYYMMDD = "000" + Function.MinGuoDate7length(DateTime.Now.ToString("yyyyMMdd"));
            strYYYYMMDD = strYYYYMMDD.Substring(strYYYYMMDD.Length - 8, 8);
            string strFileName = "匯入紀錄明細查詢" + strYYYYMMDD + ".xls";

            //* 顯示提示訊息：匯出到Excel文檔資料成功
            this.Session["ServerFile"] = strServerPathFile;
            this.Session["ClientFile"] = strFileName;
            string urlString = @"window.parent.postMessage({ func: 'ClientMsgShow', data: '" + MessageHelper.GetMessage("00_00000000_039") + "' }, '*');";
            urlString += @"location.href='DownLoadFile.aspx';";
            jsBuilder.RegScript(this.Page, urlString);
        }
        catch (Exception ex)
        {
            Logging.Log(ex);
            jsBuilder.RegScript(this.UpdatePanel1, string.Format("AlertConfirm({{title:'{0}'}});", MessageHelper.GetMessage("00_00000000_038")));
        }
    }
}
