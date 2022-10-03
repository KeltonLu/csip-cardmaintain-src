using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Data.SqlClient;
using CSIPCardMaintain.BusinessRules;
using CSIPCommonModel.BusinessRules;


/// <summary>
/// Com_AutoJob 的摘要描述
/// </summary>
public class Com_AutoJob
{
    /// <summary>
    /// 功能說明:查詢job工作狀態
    /// 作    者:Simba Liu
    /// 創建時間:2010/05/18
    /// 修改記錄:
    /// </summary>
    /// <param name="dtAutoJob"></param>
    /// <param name="strJobId"></param>
    /// <param name="strMsgID"></param>
    /// <returns></returns>
    public static bool SearchJobStatus(string strFunctionKey, ref string strJobStatus, string strJobId, ref string strMsgID)
    {
        try
        {
            string sql = "SELECT STATUS FROM M_AUTOJOB WITH(NOLOCK) WHERE JOB_ID= @jobID AND FUNCTION_KEY = @functionKey";

            SqlCommand sqlcmd = new SqlCommand();
            sqlcmd.CommandType = CommandType.Text;
            sqlcmd.CommandText = sql;
            sqlcmd.Parameters.Add(new SqlParameter("@jobID", strJobId));
            sqlcmd.Parameters.Add(new SqlParameter("@functionKey", strFunctionKey));

            DataSet ds = BRFORM_COLUMN.SearchOnDataSet(sqlcmd, "Connection_CSIP");
            if (ds != null)
            {
                DataTable dtAutoJob = ds.Tables[0];
                if (dtAutoJob != null && dtAutoJob.Rows.Count > 0)
                {
                    strJobStatus = dtAutoJob.Rows[0]["STATUS"].ToString();
                    strMsgID = "02_00000000_007";
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                strMsgID = "02_00000000_007";
                return false;
            }
        }
        catch (Exception exp)
        {
            BRM_AUTOJOB.SaveLog(exp.Message);
            strMsgID = "02_00000000_008";
            return false;
        }
    }

    public string GetEmailMembers(string functionKey, string propertyKey)
    {
        string emailMember = "";

        try
        {
            string sql = "SELECT PROPERTY_NAME FROM dbo.M_PROPERTY_CODE WITH(NOLOCK) WHERE FUNCTION_KEY = @functionKey AND PROPERTY_KEY = @propertyKey";

            SqlCommand sqlcmd = new SqlCommand();
            sqlcmd.CommandType = CommandType.Text;
            sqlcmd.CommandText = sql;
            sqlcmd.Parameters.Add(new SqlParameter("@functionKey", functionKey));
            sqlcmd.Parameters.Add(new SqlParameter("@propertyKey", propertyKey));

            DataSet ds = BRFORM_COLUMN.SearchOnDataSet(sqlcmd, "Connection_CSIP");
            if (ds != null)
            {
                string tmpMail = "";
                DataTable dtAutoJob = ds.Tables[0];
                if (dtAutoJob != null && dtAutoJob.Rows.Count > 0)
                {
                    for (int i = 0; i < dtAutoJob.Rows.Count; i++)
                    {
                        tmpMail = dtAutoJob.Rows[i]["PROPERTY_NAME"].ToString().Split(':')[1];
                        emailMember += tmpMail + ";";
                    }
                }
            }
        }
        catch (Exception exp)
        {
            BRM_AUTOJOB.SaveLog(exp.Message);
        }

        return emailMember;
    }
}
