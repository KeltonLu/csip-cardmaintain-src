//******************************************************************
//*  作    者：yangyu(rosicky)
//*  功能說明：VD卡人表
//*  創建日期：2009/09/21
//*  修改記錄：
//*<author>            <time>            <TaskID>                <desc>
//*******************************************************************
using System;
using System.Collections.Generic;
using System.Text;
using Framework.Data.OM;
using CSIPCardMaintain.EntityLayer;
using System.Data;
using Framework.Data.OM.Collections;
using Framework.Data.OM.Transaction;
using System.Data.SqlClient;

namespace CSIPCardMaintain.BusinessRules
{
    public class BRImprot_Log : CSIPCommonModel.BusinessRules.BRBase<EntityImport_Log>
    {

        /// <summary>
        /// Insert Data
        /// </summary>
        /// <param name="strInD">寫入日期</param>
        /// <param name="strFName">檔案名稱</param>
        /// <returns>是否成功</returns>
        public static bool Insert(string strInD, string strFName)
        {
            EntityImport_Log eImportLog = new EntityImport_Log();

            string strCon = " INDate = '" + strInD + "' and FileName = '" + strFName + "'";

            try
            {
                if (!eImportLog.FillSelf(strCon))
                {
                    eImportLog.INDate = strInD;
                    eImportLog.FileName = strFName;

                    return BRImprot_Log.AddNewEntity(eImportLog);
                }

                return true;
            }
            catch (Exception ex)
            {
                BRImprot_Log.SaveLog(ex);
                return false;
            }
        }

        /// <summary>
        /// Update Data
        /// </summary>
        /// <param name="strInD">寫入日期</param>
        /// <param name="strFName">檔案名稱</param>
        /// <param name="iRNum">成功筆數</param>
        /// <param name="strStatus">狀態</param>
        /// <param name="iENum">失敗筆數</param>
        /// <returns>是否成功</returns>
        public static bool Update(string strInD, string strFName, int iRNum, string strStatus, int iENum)
        {
            EntityImport_Log eImportLog = new EntityImport_Log();

            eImportLog.RecordNums = iRNum;
            eImportLog.Active_Status = strStatus;
            eImportLog.ErrorNums = iENum;

            string strCon = " INDate = '" + strInD + "' and FileName = '" + strFName + "'";

            try
            {
                return BRImprot_Log.UpdateEntityByCondition(eImportLog, strCon, EntityImport_Log.M_RecordNums, EntityImport_Log.M_Active_Status, EntityImport_Log.M_ErrorNums);
            }
            catch (Exception ex)
            {
                BRImprot_Log.SaveLog(ex);
                return false;
            }
        }

    }
}
