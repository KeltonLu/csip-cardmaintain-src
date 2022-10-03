//******************************************************************
//*  作    者：蘇洺葳
//*  功能說明：表單欄位
//*  創建日期：2018/02/08
//*  修改記錄：
//*<author>            <time>            <TaskID>                <desc>
//*Ares Luke          2021/02/22         20200031-CSIP EOS       調整connectionStr方式
//*Ares Stanley    2022/02/17    20210058-CSIP作業服務平台現代化II    白箱調整
//*******************************************************************
using System;
using System.Collections.Generic;
using System.Text;
using Framework.Data.OM;
using CSIPCommonModel.BaseItem;
using CSIPCommonModel.EntityLayer;
using CSIPCommonModel.BusinessRules;
using System.Data;
using Framework.Data.OM.Collections;
using Framework.Data.OM.Transaction;
using System.Configuration;
using System.Data.SqlClient;
using Framework.Common;
using Framework.Common.Message;
using Framework.Common.Logging;
using Framework.Common.Utility;
using CSIPCardMaintain.EntityLayer;



namespace CSIPCardMaintain.BusinessRules
{
    public class BRFORM_COLUMN : CSIPCommonModel.BusinessRules.BRBase<EntityFORM_COLUMN>
    {
        /// <summary>
        /// 大量資料匯入
        /// </summary>
        /// <param name="tableName"></param>
        /// <param name="dtInvData"></param>
        /// <returns></returns>
        public static bool InsertToInvData(string tableName, DataTable dataTable)
        {
            bool result = false;
            // string connnection = System.Configuration.ConfigurationManager.ConnectionStrings["Connection_System"].ConnectionString;
            string connnection = UtilHelper.GetConnectionStrings("Connection_System");
            SqlConnection conn = new SqlConnection(connnection);
            SqlBulkCopy sbc = new SqlBulkCopy(connnection);
            sbc.DestinationTableName = tableName;

            try
            {
                conn.Open();
                sbc.WriteToServer(dataTable);
                dataTable.Clear();

                result = true;
                return result;
            }
            catch (Exception ex)
            {
                Logging.Log(ex, LogLayer.DB);
                return result;
            }
            finally
            {
                dataTable.Clear();
                dataTable.Dispose();
                sbc.Close();
                conn.Close();
                conn.Dispose();
            }
        }
    }
}