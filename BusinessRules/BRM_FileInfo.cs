//******************************************************************
//*  功能說明：交換檔案業務邏輯層
//*  作    者：Ares Rick
//*  創建日期：2021/12/27
//*  修改記錄：
//*<author>            <time>            <TaskID>            <desc>
//*******************************************************************
using System;
using System.Collections.Generic;
using System.Text;
using Framework.Data.OM;
using Framework.Data.OM.Collections;
using Framework.Data.OM.Transaction;
using CSIPCardMaintain.EntityLayer;
using System.Data.SqlClient;
using System.Data;

namespace CSIPCardMaintain.BusinessRules
{
    public class BRM_FileInfo : CSIPCommonModel.BusinessRules.BRBase<Entity_FileInfo>
    {

        /// <summary>
        /// 功能說明:新增一筆Job資料
        /// 作    者:Simba Liu
        /// 創建時間:2010/04/26
        /// 修改記錄:
        /// </summary>
        /// <param name="FileInfo"></param>
        /// <returns></returns>
        public static bool insert(Entity_FileInfo FileInfo)
        {
            try
            {
                using (OMTransactionScope ts = new OMTransactionScope())
                {
                    if (BRM_FileInfo.AddNewEntity(FileInfo))
                    {
                        ts.Complete();
                        return true;
                    }
                    else
                    {
                        return false;
                    }

                }
            }
            catch (Exception exp)
            {
                BRM_FileInfo.SaveLog(exp.Message);
                return false;
            }
        }

        /// <summary>
        /// 功能說明:刪除一筆Job資料
        /// 作    者:Simba Liu
        /// 創建時間:2010/04/26
        /// 修改記錄:
        /// </summary>
        /// <param name="FileInfo"></param>
        /// <param name="strCondition"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool delete(Entity_FileInfo FileInfo, string strCondition, ref string strMsgID)
        {
            try
            {

                using (OMTransactionScope ts = new OMTransactionScope())
                {
                    if (BRM_FileInfo.DeleteEntityByCondition(FileInfo, strCondition))
                    {
                        ts.Complete();
                        strMsgID = "06_06040100_005";
                        return true;
                    }
                    else
                    {
                        strMsgID = "06_06040100_006";
                        return false;
                    }
                }
            }
            catch (Exception exp)
            {
                BRM_FileInfo.SaveLog(exp.Message);
                strMsgID = "06_06040100_006";
                return false;
            }
        }

        /// <summary>
        /// 刪除JOBID資料
        /// </summary>
        /// <param name="eAutoJob"></param>
        /// <param name="strMsgID">ID</param>
        /// <returns>DataTable</returns>
        public static bool Delete(Entity_FileInfo eFileInfo, ref string strMsgID)
        {

            SqlHelper Sql = new SqlHelper();

            Sql.AddCondition(Entity_FileInfo.M_FileId, Operator.Equal, DataTypeUtils.Integer, eFileInfo.FileId.ToString());

            try
            {
                BRM_FileInfo.DeleteEntityByCondition(eFileInfo, Sql.GetFilterCondition());
                strMsgID = "06_06040000_019";
                return true;
            }
            catch
            {
                strMsgID = "06_06040000_020";
                return false;
            }
        }

        /// <summary>
        /// 功能說明:更新一筆Job資料
        /// 作    者:Simba Liu
        /// 創建時間:2010/04/26
        /// 修改記錄:
        /// </summary>
        /// <param name="FileInfo"></param>
        /// <param name="strCondition"></param>
        /// <param name="strMsgID"></param>
        /// <returns></returns>
        public static bool update(Entity_FileInfo FileInfo, string strCondition, ref string strMsgID, params  string[] FiledSpit)
        {
            try
            {
                using (OMTransactionScope ts = new OMTransactionScope())
                {
                    if (BRM_FileInfo.UpdateEntityByCondition(FileInfo, strCondition, FiledSpit))
                    {
                        ts.Complete();
                        strMsgID = "06_06040100_003";
                        return true;
                    }
                    else
                    {
                        strMsgID = "06_06040100_004";
                        return false;
                    }
                }
            }
            catch (Exception exp)
            {
                BRM_FileInfo.SaveLog(exp.Message);
                strMsgID = "06_06040100_004";
                return false;
            }
        }


        /// <summary>
        /// 功能說明:根據JOB ID和FUNCTION_KEY查詢Job狀態
        /// 作    者:Simba Liu
        /// 創建時間:2010/04/26
        /// 修改記錄: 2020/12/03 Ares Luke 處理白箱報告SQL Injection
        /// </summary>
        /// <param name="dtFileInfo"></param>
        /// <param name="strJobId"></param>
        /// <returns></returns>
        public static bool selectFileInfo(ref  DataTable dtFileInfo, string strJobId)
        {
            try
            {
                string sql = @"SELECT [Job_ID]
                                      ,[FtpFileName]
                                      ,[MerchCode]
                                      ,[MerchName]
                                      ,[AMPMFlg]
                                      ,[CardType]
                                      ,[FtpPath]
                                      ,[ZipPwd]
                                      ,[CancelTime]
                                      ,[BLKCode]
                                      ,[MEMO]
                                      ,[ReasonCode]
                                      ,[ActionCode]
                                      ,[CWBRegions]
                                      ,[FunctionFlg]
                                      ,[PExpFlg]
                                      ,[BExpFlg]
                                      ,[FtpIP]
                                      ,[FtpUserName]
                                      ,[FtpPwd]
                                      ,[Status]
                                      ,[ImportDate]
                                  FROM [dbo].[tbl_FileInfo]
                                where Status ='U'and Job_ID = @strJobId order by MerchCode,CardType";
              
                SqlCommand sqlcmd = new SqlCommand();
                sqlcmd.CommandType = CommandType.Text;
                sqlcmd.CommandText = sql;
                sqlcmd.Parameters.Add(new SqlParameter("@strJobId", strJobId));


                DataSet ds = BRM_FileInfo.SearchOnDataSet(sqlcmd);
                if (ds != null)
                {
                    dtFileInfo = ds.Tables[0];
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception exp)
            {
                BRM_FileInfo.SaveLog(exp.Message);
                return false;
            }

        }

        /// <summary>
        /// 功能說明:查詢所有交換檔資料
        /// 作    者:HAO CHEN
        /// 創建時間:2010/07/21
        /// 修改記錄:
        /// </summary>
        /// <param name="dtFileInfo"></param>
        /// <param name="strJobId"></param>
        /// <returns></returns>
        public static DataTable selectFileAll(ref string strMsg)
        {
            StringBuilder sbSql = new StringBuilder();
            sbSql.Append("SELECT FileId,[Job_ID],[FtpFileName],[Status] ");
            sbSql.Append(" FROM [dbo].[tbl_FileInfo] order by FileId ");
            DataTable dtFileAll = null;
            try
            {
                dtFileAll = BRM_FileInfo.SearchOnDataSet(sbSql.ToString()).Tables[0];
            }                            
            catch (Exception exp)
            {
                strMsg = "00_00000000_000";
                throw exp;
            }
            return dtFileAll;
        }

        /// <summary>
        /// 功能說明:查詢解壓縮密碼
        /// 作    者:Simba Liu
        /// 創建時間:2010/04/26
        /// 修改記錄:
        /// </summary>
        /// <param name="dtFileInfo"></param>
        /// <param name="strJobId"></param>
        /// <returns></returns>
        public static bool selectZipPwd(ref  DataTable dtZipPwd)
        {
            try
            {
                string sql = @"SELECT 
                                    [item] 
                                    FROM [dbo].[is_commoncombo]
                                    WHERE app = 'mm' and idn1 = '1' and idn2='1' and status ='u'";

                SqlCommand sqlcmd = new SqlCommand();
                sqlcmd.CommandType = CommandType.Text;
                sqlcmd.CommandText = sql;
                DataSet ds = BRM_FileInfo.SearchOnDataSet(sqlcmd);
                if (ds != null)
                {
                    dtZipPwd = ds.Tables[0];
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception exp)
            {
                BRM_FileInfo.SaveLog(exp.Message);
                return false;
            }

        }


        /// <summary>
        /// 功能說明: 清除 FileInfo Parameter 參數 
        /// 作    者: Luke Ares
        /// 創建時間: 2021/03/12
        /// 修改記錄:
        /// </summary>
        /// <param name="jobID"></param>
        /// <returns></returns>
        public static bool UpdateParam(string jobID, string param)
        {
            string strSql = @"update dbo.tbl_FileInfo set Parameter = @Param where Job_ID = @jobID ";

            SqlCommand sqlComm = new SqlCommand();

            try
            {
                sqlComm.CommandType = CommandType.Text;
                sqlComm.CommandText = strSql;

                sqlComm.Parameters.Add(new SqlParameter("jobID", jobID));
                sqlComm.Parameters.Add(new SqlParameter("Param", param));
            }
            catch (Exception exp)
            {
                BRM_FileInfo.SaveLog(exp.Message);
                return false;
            }

            return Update(sqlComm);
        }

        /// <summary>
        /// 取得Job相關資料
        /// </summary>
        /// <param name="jobID"></param>
        /// <param name="fileInfo"></param>
        /// <returns></returns>
        public static bool GetFTPFileInfo(string jobID, ref DataTable fileInfo)
        {
            const string sqlText = @"SELECT * FROM [dbo].[tbl_FileInfo] WITH(NOLOCK) WHERE Job_ID = @Job_ID";
            try
            {
                SqlCommand sqlcmd = new SqlCommand();
                sqlcmd.CommandText = sqlText;
                sqlcmd.Parameters.Add(new SqlParameter("@Job_ID", jobID));

                DataSet ds = BRM_FileInfo.SearchOnDataSet(sqlcmd);
                if (ds != null)
                {
                    fileInfo = ds.Tables[0];
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch(Exception ex)
            {
                BRM_FileInfo.SaveLog(ex.ToString());
                return false;
            }
        }
    }
}
