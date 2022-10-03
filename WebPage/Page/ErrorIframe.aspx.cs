//******************************************************************
//*  作    者：Luke
//*  功能說明：錯誤畫面，重定向到登陸畫面
//*  創建日期：2020/04/26
//*  修改紀錄：
//*  <author>            <time>            <TaskID>                <desc>
//*  Ares Luke          2020/11/19         20200031-CSIP EOS       調整取web.config加解密參數
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
using Framework.Common.JavaScript;
using Framework.Common.Message;
using Framework.Common.Utility;

public partial class Page_ErrorIframe : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        //不清除Session，並導入登入頁面
        string strMsg = "00_00000000_000";
        string strUrlLogon = UtilHelper.GetAppSettings("ERROR2").ToString();

        jsBuilder.RegScript(this.Page, "alert('" + MessageHelper.GetMessage(strMsg) + "');" + "var local = window.parent.location!=window.location?window.parent:window.opener?window.opener.parent:window;local.location.href='" + strUrlLogon + "';");
    }
}
