//******************************************************************
//*  功能說明：Audit Log System: SOC (Security Operation Center) 資料備查
//*  作    者：James
//*  創建日期：2019/07/02
//*  修改記錄：
//*<author>            <time>            <TaskID>            <desc>
//*******************************************************************

using CSIPCommonModel.EntityLayer;
using CSIPCommonModel.EntityLayer_new;
using Framework.Common.Logging;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

/// <summary>
/// BRL_AP_LOG 的摘要描述
/// </summary>
public class BRL_AP_LOG : CSIPCommonModel.BusinessRules.BRBase<EntityL_AP_LOG>
{
    private const string _tableName = "[L_AP_LOG]";
    /// <summary>
    /// 系統別，因sPageInfo 只帶入FunctionID八碼，故需增加此代號
    /// </summary>
    private const string _SystemCode = "03";

    public BRL_AP_LOG()
    {
    }

    /// <summary>
    /// 取得預設(必要值)
    /// </summary>
    /// <param name="agentInfo"></param>
    /// <param name="FunctionID">網頁FunctionID</param>
    /// <returns></returns>
    public static EntityL_AP_LOG getDefaultValue(EntityAGENT_INFO agentInfo, string FunctionID)
    {
        EntityL_AP_LOG result = new EntityL_AP_LOG();

        try
        {
            DateTime CurrentDate = DateTime.Now;
            
            result.System_Code = "CSIP";
            result.Login_Account_Nbr = agentInfo.agent_id;
            result.Query_Datetime = CurrentDate.ToString("yyyy-MM-dd HH:mm:ss.fff");
            //Talas 20191002 調整為  _SystemCode + FunctionID
            //result.AP_Txn_Code = FunctionID;
            result.AP_Txn_Code = _SystemCode + FunctionID;
            result.Server_Name = "";
            result.User_Terminal = "";
            result.AP_Account_Nbr = "";
            result.Txn_Type_Code = "";
            result.Statement_Text = "";
            result.Object_Name = "";
            result.Txn_Status_Code = "";
            result.Customer_Id = "";
            result.Account_Nbr = "";
            result.Branch_Nbr = "";
            result.Role_Id = "";
            result.Import_Source = "CSIP_APLOG";
            result.As_Of_Date = CurrentDate.ToString("yyyyMMdd");
        }
        catch (Exception)
        {
        }

        return result;
    }

    public static List<EntityL_AP_LOG> Query(EntityL_AP_LOG paramObj,string SearchStart,string SearchEnd)
    {
        string sql = @"SELECT  [ID],[BasicTaxID],[keyin_day],[BasicRegistyNameEN],[BasicRegistyNameCH]
      ,[user_1key],[user_2key] FROM  " + _tableName;
        SqlCommand sqlcmd = new SqlCommand();
        sqlcmd.CommandType = CommandType.Text;

        #region 查詢條件
        List<string> _par = new List<string>();
        if (!string.IsNullOrEmpty(paramObj.Customer_Id))
        {
            _par.Add("Customer_Id = @Customer_Id ");
            sqlcmd.Parameters.Add(new SqlParameter("@Customer_Id", paramObj.Customer_Id));
        }
        
        if (!string.IsNullOrEmpty(SearchStart) && !string.IsNullOrEmpty(SearchEnd))
        {
            _par.Add("[As_Of_Date] BETWEEN @SearchStart AND @SearchEnd");
            sqlcmd.Parameters.Add(new SqlParameter("@SearchStart", SearchStart));
            sqlcmd.Parameters.Add(new SqlParameter("@SearchEnd", SearchEnd));
        }
        if (_par.Count > 0)
        {
            sql = string.Format("{0} where {1}", sql, string.Join(" and ", _par.ToArray()));
        }
        #endregion

        sqlcmd.CommandText = sql;
        try
        {
            DataTable dt = new DataTable();
            DataSet DS = SearchOnDataSet(sqlcmd);
            if (DS != null && DS.Tables.Count > 0)
            {
                dt = DS.Tables[0];
            }
            List<EntityL_AP_LOG> rtn = DataTableConvertor.ConvertCollToObj<EntityL_AP_LOG>(dt);
            return rtn;
        }
        catch (Exception ex)
        {
            Logging.Log("查詢 SOC 資料失敗：" + ex, LogState.Error, LogLayer.BusinessRule);
            return null;
        }


    }

    public static bool Add(EntityL_AP_LOG paramObj)
    { 

        bool result = false;
        //Talas 調整組成字串至ADD方法，頁面只提供值即可
        paramObj.Statement_Text = string.Format("CUSTOMER_ID={0}|AC_NO={1}|BRANCH_ID={2}|ROLE_ID={3}",
         paramObj.Customer_Id, paramObj.Account_Nbr, paramObj.Branch_Nbr, paramObj.Role_Id); //查詢條件內容: 用 | 區隔,要看文件確認        

        //執行參數: 用 | 區隔,若無值仍需帶屬性名,切記不要內含 逗點, 所以要置換為空白
        paramObj.Statement_Text = paramObj.Statement_Text.Replace(",", "");

        string sSQL = string.Format("Insert into {0} ", _tableName) +
"([System_Code],[Login_Account_Nbr],[Query_Datetime],[AP_Txn_Code],[Server_Name],[User_Terminal]" +
",[AP_Account_Nbr],[Txn_Type_Code],[Statement_Text],[Object_Name],[Txn_Status_Code],[Customer_Id]" +
",[Account_Nbr],[Branch_Nbr],[Role_Id],[Import_Source],[As_Of_Date])" +
"VALUES(" +
"@System_Code,@Login_Account_Nbr,@Query_Datetime,@AP_Txn_Code,@Server_Name,@User_Terminal" +
",@AP_Account_Nbr,@Txn_Type_Code,@Statement_Text,@Object_Name,@Txn_Status_Code,@Customer_Id" +
",@Account_Nbr,@Branch_Nbr,@Role_Id,@Import_Source,@As_Of_Date)";

        SqlCommand sqlcmd = new SqlCommand();
        sqlcmd.CommandType = CommandType.Text;
        sqlcmd.CommandText = sSQL;

        sqlcmd.Parameters.Add(new SqlParameter("@System_Code", paramObj.System_Code));
        sqlcmd.Parameters.Add(new SqlParameter("@Login_Account_Nbr", paramObj.Login_Account_Nbr));
        sqlcmd.Parameters.Add(new SqlParameter("@Query_Datetime", paramObj.Query_Datetime));
        sqlcmd.Parameters.Add(new SqlParameter("@AP_Txn_Code", paramObj.AP_Txn_Code));
        sqlcmd.Parameters.Add(new SqlParameter("@Server_Name", paramObj.Server_Name));
        sqlcmd.Parameters.Add(new SqlParameter("@User_Terminal", paramObj.User_Terminal));
        sqlcmd.Parameters.Add(new SqlParameter("@AP_Account_Nbr", paramObj.AP_Account_Nbr));
        sqlcmd.Parameters.Add(new SqlParameter("@Txn_Type_Code", paramObj.Txn_Type_Code));
        sqlcmd.Parameters.Add(new SqlParameter("@Statement_Text", paramObj.Statement_Text));
        sqlcmd.Parameters.Add(new SqlParameter("@Object_Name", paramObj.Object_Name));
        sqlcmd.Parameters.Add(new SqlParameter("@Txn_Status_Code", paramObj.Txn_Status_Code));
        sqlcmd.Parameters.Add(new SqlParameter("@Customer_Id", paramObj.Customer_Id));
        sqlcmd.Parameters.Add(new SqlParameter("@Account_Nbr", paramObj.Account_Nbr));
        sqlcmd.Parameters.Add(new SqlParameter("@Branch_Nbr", paramObj.Branch_Nbr));
        sqlcmd.Parameters.Add(new SqlParameter("@Role_Id", paramObj.Role_Id));
        sqlcmd.Parameters.Add(new SqlParameter("@Import_Source", paramObj.Import_Source));
        sqlcmd.Parameters.Add(new SqlParameter("@As_Of_Date", paramObj.As_Of_Date));
        
        result = Add(sqlcmd, "Connection_CSIP");
        //result = AddNewEntity(paramObj, "Connection_CSIP");
        return result;

    }   
   
}