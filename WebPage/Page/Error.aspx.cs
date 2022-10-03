//******************************************************************
//*  作    者：yangyu(rosicky)
//*  功能說明：卡人維護記錄查詢
//*  創建日期：2009/10/01
//*  修改記錄：
//*<author>            <time>            <TaskID>                <desc>
//* Ares Stanley    2022/02/14    20210058-CSIP作業服務平台現代化II    調整webconfig取參數方式
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
using Framework.Common.Utility;
using Framework.Common.JavaScript;
using Framework.Common.Message;

public partial class Page_Error : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        string strUrlLogon = UtilHelper.GetAppSettings("LOGOUT").ToString();

        if (string.IsNullOrEmpty(RedirectHelper.GetDecryptString(this.Page, "MsgID")))
        {
            jsBuilder.RegScript(this.Page, string.Format("AlertConfirm({{title:'{0}', url:'{1}', local:window.parent.location!=window.location?window.parent:window.opener?window.opener.parent:window}});", MessageHelper.GetMessage("00_00000000_035"), strUrlLogon));
        }
        else
        {
            string strMsgID = RedirectHelper.GetDecryptString(this.Page, "MsgID");

            jsBuilder.RegScript(this.Page, string.Format("AlertConfirm({{title:'{0}', url:'{1}', local:window.parent.location!=window.location?window.parent:window.opener?window.opener.parent:window}});", MessageHelper.GetMessage(strMsgID), strUrlLogon));
        }
    }
}
