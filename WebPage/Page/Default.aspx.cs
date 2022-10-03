//******************************************************************
//*  作    者：宋戈
//*  功能說明：預設畫面.重新載入Session
//*  創建日期：2009/08/10
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
using CSIPCommonModel.EntityLayer;
using CSIPCommonModel.BusinessRules;
using Framework.Common.Utility;
using Framework.Common.Logging;

public partial class Page_Default : PageBase
{
    protected void Page_Load(object sender, EventArgs e)
    {
        try
        {
            string strTicketID = RedirectHelper.GetDecryptString(this.Page, "TicketID");
            getSessionFromDB(strTicketID, ref strMsg);
        }
        catch(Exception ex)
        {
            Logging.Log(ex);
        }
    }

    /// <summary>
    /// 以TicketID到DB中取Session資料。
    /// </summary>
    /// <param name="strTicketID"></param>
    private bool getSessionFromDB(String strTicketID, ref string strMsg)
    {
        try
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
            return true;
        }
        catch (Exception ex)
        {
            Logging.Log(ex);
            return false;

        }
        
    }
}
