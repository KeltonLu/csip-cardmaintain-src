//******************************************************************
//*  作    者：占偉林(James)
//*  功能說明：系統角色維護

//*  創建日期：2009/07/10
//*  修改記錄：

//*<author>            <time>            <TaskID>                <desc>
//*******************************************************************

using System;
using System.Data;
using System.Globalization;
using System.Threading;
using System.Configuration;
using System.Collections.Generic;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.ComponentModel;
using System.Text;
using System.Reflection;
using CSIPCommonModel.BusinessRules;
using CSIPCommonModel.EntityLayer;
using CSIPCommonModel.BaseItem;
using Framework.Common.Utility;
using Framework.Common.Message;
using Framework.Data.OM;
using Framework.Data.OM.Collections;
using Framework.WebControls;
using Framework.Common.JavaScript;
using Framework.Common;
using Framework.Common.Logging;

/// <summary>
/// 要做權限判斷的頁面基礎類別

/// </summary>
public class PageBase : System.Web.UI.Page
{
    protected long ProgramBeginRunTime;
    protected long programRunTime;
    protected string strMsg = "";

    public PageBase()
    {
        this.ProgramBeginRunTime = System.Environment.TickCount; //程序开始运行时间
    }
    /*增加記錄網頁訊息的struct Add by 陳靜嫻2009-09-21 Start */
    public struct structPageInfo
    {
        public string strPageCode;//*網頁FunctionID
        public string strPageName;//*網頁名稱
    }
    /*增加記錄網頁訊息的struct Add by 陳靜嫻2009-09-21 End */
    /// <summary>
    /// 填充页面上显示程序运行时间的文本控件
    /// </summary>
    /// <param name="literal">显示程序运行时间的文本控件</param>
    private void ProgramRunTime()
    {
        long ProgramEndRunTime = System.Environment.TickCount;
        programRunTime = ProgramEndRunTime - this.ProgramBeginRunTime;
        // jsBuilder.RegScript(this.Page, "var local = window.parent.location!=window.location?window.parent:window.opener?window.opener.parent:window;local.document.all.runtime.innerText='" + programRunTime.ToString() + " 毫秒';");
        //jsBuilder.RegScript(this.Page, "var local = window.parent.location!=window.location?window.parent:window.opener?window.opener.parent:window;local.document.getElementById('runtime').innerText='" + programRunTime.ToString() + " 毫秒';");
        jsBuilder.RegScript(this.Page, "window.parent.postMessage({ func: 'ProgramRunTime', data: '" + programRunTime.ToString() + " 毫秒' }, '*');");
    }

    protected override void Render(System.Web.UI.HtmlTextWriter writer)
    {
        this.ProgramRunTime();
        base.Render(writer);
    }


    /// <summary>
    /// 頁面的Function_ID
    /// </summary>
    private String _Function_ID;

    /// <summary>
    /// 黨頁面加載時
    /// </summary>
    /// <param name="e">事件參數</param>
    protected override void OnLoad(EventArgs e)
    {
        //檢核操作瀏覽器功能
        CensorPage();

        /*增加記錄網頁訊息的struct Add by 陳靜嫻2009-09-21 Start */
        structPageInfo sPageInfo = new structPageInfo();
        /*增加記錄網頁訊息的struct Add by 陳靜嫻2009-09-21 End */

        string strMsg = "";
        string strUrlError = UtilHelper.GetAppSettings("Error").ToString();

        //*判斷Session是否存在
        if (this.Session["Agent"] == null)
        {
            #region 判斷Session是否存在及重新取Session值
            //*Session不存在時，判斷TicketID是否存在
            if (string.IsNullOrEmpty(RedirectHelper.GetDecryptString(this.Page, "TicketID")))
            {
                strMsg = "00_00000000_035";
                //*TicketID不存在，顯示重新登入訊息，轉向重新登入畫面
                Response.Redirect(strUrlError + "?MsgID=" + RedirectHelper.GetEncryptParam(strMsg), false);
            }
            else
            {
                //*TicketID存在時，
                //*取TicketID
                string strTicketID = RedirectHelper.GetDecryptString(this.Page, "TicketID");
                //*以TicketID到DB中取Session資料。

                if (!getSessionFromDB(strTicketID, ref strMsg))
                {
                    Response.Redirect(strUrlError + "?MsgID=" + RedirectHelper.GetEncryptParam(strMsg), false);
                }
            }
            #endregion 判斷Session是否存在及重新取Session值
        }
        else
        {
            #region 判斷用戶是否有使用該頁面的權限

            //*取頁面的功能ID號(Function_ID)
            this._Function_ID = "88888888";
            string strPath = this.Server.MapPath(this.Request.Url.AbsolutePath).ToUpper();
            if (strPath.IndexOf("DEFAULT") == -1)
            {
                PageAction pgaNow = PopedomManager.MainPopedomManager.PageSettings[strPath];
                this._Function_ID = pgaNow.FunctionID;   // 頁面的功能ID號
                /*Session中增加記錄網頁訊息的struct Add by 陳靜嫻2009-09-21 Start */
                sPageInfo.strPageCode = pgaNow.FunctionID;
                this.Session["PageInfo"] = sPageInfo;
                //20220616_Ares_Jack_ Log新增FunctionID
                Logging.UpdateLogAgentFunctionId(sPageInfo.strPageCode);
                /*Session中增加記錄網頁訊息的struct Add by 陳靜嫻2009-09-21 End */
                bool blCanUseAction = false;
                //*檢查用戶的權限列表中是否存在當前頁面的Funcion_ID;
                for (int intLoop = 0; intLoop < ((DataTable)((EntityAGENT_INFO)this.Session["Agent"]).dtfunction).Rows.Count; intLoop++)
                {
                    if (((DataTable)((EntityAGENT_INFO)this.Session["Agent"]).dtfunction).Rows[intLoop]["Function_ID"].ToString() == this._Function_ID)
                    {
                        blCanUseAction = true;
                        break;
                    }
                }

                //*沒有權限使用該功能ID
                if (!blCanUseAction)
                {
                    Logging.Log("CardMaintain_沒有權限使用該功能");
                    strMsg = "00_00000000_025";
                    Response.Redirect(strUrlError + "?MsgID=" + RedirectHelper.GetEncryptParam(strMsg), false);
                    return;
                }
            }
            #endregion
        }



        base.OnLoad(e);
    }

    /// <summary>
    /// 以TicketID到DB中取Session資料。
    /// </summary>
    /// <param name="strTicketID"></param>
    private bool getSessionFromDB(String strTicketID, ref string strMsg)
    {
        EntityAGENT_INFO eAgentInfo = new EntityAGENT_INFO();

        EntitySESSION_INFO eSessionInfo = new EntitySESSION_INFO();

        eSessionInfo.TICKET_ID = strTicketID;

        //* 取Session訊息
        if (!BRSESSION_INFO.Search(eSessionInfo, ref eAgentInfo, ref strMsg))
        {
            return false;
        }

        //* 重新回覆當前Session的訊息
        this.Session["Agent"] = eAgentInfo;
        //20220616_Ares_Jack_ Log新增UserID
        Logging.NewLogAgent(eAgentInfo.agent_id);

        //* 刪除DB中的TicketID對應的Session訊息
        if (!BRSESSION_INFO.Delete(eSessionInfo, ref strMsg))
        {
            return false;
        }
        return true;
    }

    /// <summary>
    /// 檢核操作瀏覽器
    /// (避免分頁多開導致Session異常)
    /// </summary>
    protected void CensorPage()
    {
        string strUrlErrorIframe = UtilHelper.GetAppSettings("ERROR_IFRAME").ToString();

        try
        {
            if (!IsPostBack)
            {
                String usrBrowser = Request.Browser.Browser;
                String GUID = Guid.NewGuid().ToString();

                HttpContext.Current.Session["usrBrowser"] = usrBrowser;
                this.ViewState["usrBrowser"] = usrBrowser;

                HttpContext.Current.Session["usrGUID"] = GUID;
                this.ViewState["usrGUID"] = GUID;
            }
            else
            {
                //檢核操作瀏覽器是否相符
                if (this.ViewState["usrBrowser"].Equals(Session["usrBrowser"]))
                {
                    //檢核KEY是否相符
                    if (!Session["usrGUID"].Equals(this.ViewState["usrGUID"]))
                    {
                        Logging.Log("CardMaintain_檢核KEY不相符");
                        //20210412_Ares_Stanley-調整轉向方法
                        Response.Redirect(strUrlErrorIframe, false);
                        //jsBuilder.RegScript(this.Page, "alert('" + MessageHelper.GetMessage(strMsg) + "');var local = window.parent.location!=window.location?window.parent:window.opener?window.opener.parent:window;local.location.href='" + strUrlError2 + "';");
                        return;
                    }
                }
            }
        }
        catch (ThreadAbortException taex)
        {
            Logging.Log(taex, LogLayer.UI);
            //20210412_Ares_Stanley-調整轉向方法
            Response.Redirect(strUrlErrorIframe, false);
        }
        catch (Exception exp)
        {
            Logging.Log(exp, LogLayer.UI);
            //20210412_Ares_Stanley-調整轉向方法
            Response.Redirect(strUrlErrorIframe, false);
        }
    }
}