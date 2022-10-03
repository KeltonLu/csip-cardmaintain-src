//******************************************************************
//*  作    者：余洋
//*  功能說明：VD卡片-維護記錄查詢
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
using Framework.Common.Logging;
using CSIPCommonModel.BaseItem;

public partial class Page_P030401010001 : PageBase
{


    /// <summary>
    /// Session變數集合
    /// </summary>
    private CSIPCommonModel.EntityLayer.EntityAGENT_INFO eAgentInfo;
    private structPageInfo sPageInfo;//*記錄網頁訊息
    //*排序欄位
    string strSEQ = "";
    protected void Page_Load(object sender, EventArgs e)
    {
        Page.Title = BaseHelper.GetShowText("03_04010000_000");

        if (!Page.IsPostBack)
        {
            jsBuilder.RegScript(this.UpdatePanel1, BaseHelper.ClientMsgShow(""));

            txtID.Focus();

            //* 設定Tittle
            ShowControlsText();
            this.gpList.Visible = false;
            this.gpList.RecordCount = 0;
            this.grvUserView.Visible = false;
        }
        eAgentInfo = (CSIPCommonModel.EntityLayer.EntityAGENT_INFO)this.Session["Agent"]; //*Session變數集合
        sPageInfo = (structPageInfo)this.Session["PageInfo"];

    }

    /// <summary>
    /// 列印
    /// 修改紀錄：報表改由NPOI產出 by Ares Stanley 20211108
    /// </summary>
    protected void btnOK_Click(object sender, EventArgs e)
    {
        //------------------------------------------------------
        //AuditLog to SOC
        CSIPCommonModel.EntityLayer_new.EntityL_AP_LOG log;
        //------------------------------------------------------
        try
        {
            if (dpBeforeDate.Text != "" && dpEndDate.Text != "")
            {


                DateTime dtmBeforeData = Convert.ToDateTime(dpBeforeDate.Text.Trim());

                DateTime dtmEndData = Convert.ToDateTime(dpEndDate.Text.Trim());

                if (!BRReport.CheckDataTime(dtmBeforeData, dtmEndData))
                {
                    //修改彈跳視窗與末端訊息一樣 by Ares Neal 2022/06/15
                    jsBuilder.RegScript(this.UpdatePanel1, BaseHelper.ClientMsgShow("03_03010000_000"));
                    jsBuilder.RegScript(this.Page, string.Format("AlertConfirm({{title:'{0}'}}); if (document.getElementById('dpEndDate_foo')) document.getElementById('dpEndDate_foo').focus();", MessageHelper.GetMessage("03_03010000_000")));
                    return;
                }
            }

            if (chkID.Checked == true)
            {
                strSEQ = strSEQ + "CUST_ID,";
                //------------------------------------------------------
                //AuditLog to SOC 
                //20220411_Ares_Jack_以下註解改成有輸入ID就寫L_AP_LOG 
                //log = BRL_AP_LOG.getDefaultValue(eAgentInfo, sPageInfo.strPageCode);
                //log.Account_Nbr = this.txtID.Text;
                //BRL_AP_LOG.Add(log);
                //------------------------------------------------------
            }
            //20220411_Ares_Jack_ EOS2 AuditLog to SOC
            if (this.txtID.Text.Trim() != "")
            {
                log = BRL_AP_LOG.getDefaultValue(eAgentInfo, sPageInfo.strPageCode);
                log.Account_Nbr = this.txtID.Text;
                BRL_AP_LOG.Add(log);
            }

            if (chkPeople.Checked == true)
            {
                strSEQ = strSEQ + "USER_ID,";
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

            string strMsgID = "";
            try
            {
                string strRptID = this.txtID.Text;
                string strRptPeople = this.txtPeople.Text;
                string strRptBeforeDate = this.dpBeforeDate.Text.Replace("/", "");
                string strRptEndDate = this.dpEndDate.Text.Replace("/", "");
                string strRptSEQ = strSEQ;
                CSIPCommonModel.EntityLayer.EntityAGENT_INFO eAgentinfo = (CSIPCommonModel.EntityLayer.EntityAGENT_INFO)Session["Agent"];
                string strName = eAgentinfo.agent_name;
                string strAgentID = eAgentinfo.agent_id;//* 業務員ID

                bool blnHadRecord = false;//* 是否有資料

                //* 查詢
                if (BRReport.Report04010100_BySP(strRptID, strRptPeople, strRptBeforeDate, strRptEndDate, strRptSEQ, strMsgID, strName, strAgentID, ref blnHadRecord, "P"))
                {
                    if (!blnHadRecord)
                    {
                        //* 沒有資料
                        //修改彈跳視窗與末端訊息一樣 by Ares Neal 2022/06/15
                        jsBuilder.RegScript(this.UpdatePanel1, BaseHelper.ClientMsgShow("00_00000000_037"));
                        jsBuilder.RegScript(this.Page, string.Format("AlertConfirm({{title:'{0}'}});", MessageHelper.GetMessage("00_00000000_037")));
                        return;
                    }
                }
                else
                {
                    //* 出錯
                    //修改彈跳視窗與末端訊息一樣 by Ares Neal 2022/06/15
                    jsBuilder.RegScript(this.UpdatePanel1, BaseHelper.ClientMsgShow("00_00000000_000"));
                    jsBuilder.RegScript(this.Page, string.Format("AlertConfirm({{title:'{0}'}});", MessageHelper.GetMessage("00_00000000_000")));
                    return;
                }

                string strServerPathFile = this.Server.MapPath(UtilHelper.GetAppSettings("ExportExcelFilePath").ToString());
                bool isCSV = false;
                if (!BR_Excel_File.CreateExcelFile_Report04010100(strRptID, strRptBeforeDate, strRptEndDate, strName, strAgentID, ref strServerPathFile, ref strMsgID, ref isCSV))
                {
                    //修改彈跳視窗與末端訊息一樣 by Ares Neal 2022/06/15
                    jsBuilder.RegScript(this.UpdatePanel1, BaseHelper.ClientMsgShow("00_00000000_038"));
                    jsBuilder.RegScript(this.Page, string.Format("AlertConfirm({{title:'{0}'}});", MessageHelper.GetMessage("00_00000000_038")));
                    return;
                }
                //* 將服務器端生成的文檔，下載到本地。
                string strYYYYMMDD = "000" + Function.MinGuoDate7length(DateTime.Now.ToString("yyyyMMdd"));
                strYYYYMMDD = strYYYYMMDD.Substring(strYYYYMMDD.Length - 8, 8);
                string strFileName;
                if (!isCSV)
                {
                    strFileName = "VD卡片資料查詢列印維護記錄查詢" + strYYYYMMDD + ".xls";
                }
                else
                {
                    strFileName = "VD卡片資料查詢列印維護記錄查詢" + strYYYYMMDD + ".csv";
                }

                //* 顯示提示訊息：匯出到Excel文檔資料成功
                this.Session["ServerFile"] = strServerPathFile;
                this.Session["ClientFile"] = strFileName;
                string urlString = @"window.parent.postMessage({ func: 'ClientMsgShow', data: '" + MessageHelper.GetMessage("00_00000000_039") + "' }, '*');";
                urlString += @"location.href='DownLoadFile.aspx';";
                //修改彈跳視窗與末端訊息一樣 by Ares Neal 2022/06/15
                jsBuilder.RegScript(this.UpdatePanel1, urlString);
                jsBuilder.RegScript(this.Page, string.Format("AlertConfirm({{title:'{0}'}});", MessageHelper.GetMessage("00_00000000_039")));
            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                //修改彈跳視窗與末端訊息一樣 by Ares Neal 2022/06/15
                jsBuilder.RegScript(this.UpdatePanel1, BaseHelper.ClientMsgShow("00_00000000_000"));
                jsBuilder.RegScript(this.Page, string.Format("AlertConfirm({{title:'{0}'}});", MessageHelper.GetMessage("00_00000000_000")));
                return;
            }
        }
        catch (Exception ex)
        {
            Logging.Log(ex);
            //修改彈跳視窗與末端訊息一樣 by Ares Neal 2022/06/15
            jsBuilder.RegScript(this.UpdatePanel1, BaseHelper.ClientMsgShow("03_04010000_000"));
            jsBuilder.RegScript(this.Page, string.Format("AlertConfirm({{title:'{0}'}});", MessageHelper.GetMessage("03_04010000_000")));
            return;
        }
    }


    /// <summary>
    /// 專案代號:20210058-CSIP 作業服務平台現代化II
    /// 功能說明:業務新增查詢功能
    /// 作    者:Ares Stanley
    /// 修改時間:2021/11/08
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
    /// 修改時間:2021/11/08
    /// </summary>
    private void BindGridView()
    {
        //------------------------------------------------------
        //AuditLog to SOC
        CSIPCommonModel.EntityLayer_new.EntityL_AP_LOG log;
        //------------------------------------------------------
        try
        {
            if (dpBeforeDate.Text != "" && dpEndDate.Text != "")
            {


                DateTime dtmBeforeData = Convert.ToDateTime(dpBeforeDate.Text.Trim());

                DateTime dtmEndData = Convert.ToDateTime(dpEndDate.Text.Trim());

                if (!BRReport.CheckDataTime(dtmBeforeData, dtmEndData))
                {
                    //修改彈跳視窗與末端訊息一樣 by Ares Neal 2022/06/15
                    jsBuilder.RegScript(this.UpdatePanel1, BaseHelper.ClientMsgShow("03_03010000_000"));
                    jsBuilder.RegScript(this.Page, string.Format("AlertConfirm({{title:'{0}'}}); if (document.getElementById('dpEndDate_foo')) document.getElementById('dpEndDate_foo').focus();", MessageHelper.GetMessage("03_03010000_000")));
                    return;
                }
            }

            if (chkID.Checked == true)
            {
                strSEQ = strSEQ + "CUST_ID,";
                //------------------------------------------------------
                //AuditLog to SOC
                //20220411_Ares_Jack_以下註解改成有輸入ID就寫L_AP_LOG  
                //log = BRL_AP_LOG.getDefaultValue(eAgentInfo, sPageInfo.strPageCode);
                //log.Account_Nbr = this.txtID.Text;
                //BRL_AP_LOG.Add(log);
                //------------------------------------------------------
            }
            //20220411_Ares_Jack_ EOS2 AuditLog to SOC
            if (this.txtID.Text.Trim() != "")
            {
                log = BRL_AP_LOG.getDefaultValue(eAgentInfo, sPageInfo.strPageCode);
                log.Account_Nbr = this.txtID.Text;
                BRL_AP_LOG.Add(log);
            }

            if (chkPeople.Checked == true)
            {
                strSEQ = strSEQ + "USER_ID,";
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

            string strMsgID = "";
            try
            {
                string strRptID = this.txtID.Text;
                string strRptPeople = this.txtPeople.Text;
                string strRptBeforeDate = this.dpBeforeDate.Text.Replace("/", "");
                string strRptEndDate = this.dpEndDate.Text.Replace("/", "");
                string strRptSEQ = strSEQ;
                CSIPCommonModel.EntityLayer.EntityAGENT_INFO eAgentinfo = (CSIPCommonModel.EntityLayer.EntityAGENT_INFO)Session["Agent"];
                string strName = eAgentinfo.agent_name;
                string strAgentID = eAgentinfo.agent_id;//* 業務員ID

                bool blnHadRecord = false;//* 是否有資料

                //* 查詢
                if (BRReport.Report04010100_BySP(strRptID, strRptPeople, strRptBeforeDate, strRptEndDate, strRptSEQ, strMsgID, strName, strAgentID, ref blnHadRecord))
                {
                    if (!blnHadRecord)
                    {
                        //* 沒有資料
                        //修改彈跳視窗與末端訊息一樣 by Ares Neal 2022/06/15
                        jsBuilder.RegScript(this.UpdatePanel1, BaseHelper.ClientMsgShow("00_00000000_037"));
                        jsBuilder.RegScript(this.Page, string.Format("AlertConfirm({{title:'{0}'}});", MessageHelper.GetMessage("00_00000000_037")));
                        return;
                    }
                }
                else
                {
                    //* 出錯
                    //修改彈跳視窗與末端訊息一樣 by Ares Neal 2022/06/15
                    jsBuilder.RegScript(this.UpdatePanel1, BaseHelper.ClientMsgShow("00_00000000_000"));
                    jsBuilder.RegScript(this.Page, string.Format("AlertConfirm({{title:'{0}'}});", MessageHelper.GetMessage("00_00000000_000")));
                    return;
                }

                int totalCount = 0;
                //根據USER需求若有填寫卡人ID或卡號就以日期新到舊排序 by Ares Stanley 20220728
                bool isOrderBy = !string.IsNullOrEmpty(this.txtID.Text.Trim());
                DataTable dt = BR_Excel_File.getData_member1(strAgentID, ref totalCount, "04010100", this.gpList.CurrentPageIndex, "S", isOrderBy);
                if (dt.Rows.Count > 0)
                {
                    this.gpList.Visible = true;
                    this.gpList.RecordCount = totalCount;
                    this.grvUserView.Visible = true;
                    this.grvUserView.DataSource = dt;
                    this.grvUserView.DataBind();
                    //修改彈跳視窗與末端訊息一樣 by Ares Neal 2022/06/15
                    jsBuilder.RegScript(this.UpdatePanel1, BaseHelper.ClientMsgShow("03_01010000_006"));
                    jsBuilder.RegScript(this.Page, string.Format("AlertConfirm({{title:'{0}'}});", MessageHelper.GetMessage("03_01010000_006")));
                }
                else
                {
                    this.gpList.RecordCount = 0;
                    this.grvUserView.DataSource = null;
                    this.grvUserView.DataBind();
                    this.gpList.Visible = false;
                    this.grvUserView.Visible = false;
                    jsBuilder.RegScript(this.UpdatePanel1, BaseHelper.ClientMsgShow("00_00000000_037"));
                    jsBuilder.RegScript(this.Page, string.Format("AlertConfirm({{title:'{0}'}});", MessageHelper.GetMessage("00_00000000_037")));
                }
            }
            catch (Exception exp)
            {
                BRReport.SaveLog(exp);
                //修改彈跳視窗與末端訊息一樣 by Ares Neal 2022/06/15
                jsBuilder.RegScript(this.UpdatePanel1, BaseHelper.ClientMsgShow("00_00000000_000"));
                jsBuilder.RegScript(this.Page, string.Format("AlertConfirm({{title:'{0}'}});", MessageHelper.GetMessage("00_00000000_000")));
                return;
            }
        }
        catch (Exception ex)
        {
            Logging.Log(ex);
            //修改彈跳視窗與末端訊息一樣 by Ares Neal 2022/06/15
            jsBuilder.RegScript(this.UpdatePanel1, BaseHelper.ClientMsgShow("03_04010000_000"));
            jsBuilder.RegScript(this.Page, string.Format("AlertConfirm({{title:'{0}'}});", MessageHelper.GetMessage("03_04010000_000")));
            return;
        }
    }

    /// <summary>
    /// 專案代號:20210058-CSIP 作業服務平台現代化II
    /// 功能說明:業務新增查詢切換頁需求功能
    /// 作    者:Ares Stanley
    /// 修改時間:2021/11/08
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
    /// 修改時間:2021/11/08
    /// </summary>
    protected void ShowControlsText()
    {
        {
            //* 設置查詢結果GridView的列頭標題
            this.grvUserView.Columns[0].HeaderText = BaseHelper.GetShowText("03_04010000_024");//卡人ID
            this.grvUserView.Columns[1].HeaderText = BaseHelper.GetShowText("03_04010000_025");//欄位名稱
            this.grvUserView.Columns[2].HeaderText = BaseHelper.GetShowText("03_04010000_026");//調整前內容
            this.grvUserView.Columns[3].HeaderText = BaseHelper.GetShowText("03_04010000_027");//調整後內容
            this.grvUserView.Columns[4].HeaderText = BaseHelper.GetShowText("03_04010000_028");//維護日期
            this.grvUserView.Columns[5].HeaderText = BaseHelper.GetShowText("03_04010000_029");//維護時間
            this.grvUserView.Columns[6].HeaderText = BaseHelper.GetShowText("03_04010000_002");//維護員

            //* 設置一頁顯示最大筆數
            this.gpList.PageSize = int.Parse(UtilHelper.GetAppSettings("PageSize"));
            this.grvUserView.PageSize = int.Parse(UtilHelper.GetAppSettings("PageSize"));
        }
    }

}
