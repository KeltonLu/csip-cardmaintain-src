//******************************************************************
//*  �@    �̡Gyangyu(rosicky)
//*  �\�໡���GVD�d�H��
//*  �Ыؤ���G2009/09/21
//*  �ק�O���G
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
        /// <param name="strInD">�g�J���</param>
        /// <param name="strFName">�ɮצW��</param>
        /// <returns>�O�_���\</returns>
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
        /// <param name="strInD">�g�J���</param>
        /// <param name="strFName">�ɮצW��</param>
        /// <param name="iRNum">���\����</param>
        /// <param name="strStatus">���A</param>
        /// <param name="iENum">���ѵ���</param>
        /// <returns>�O�_���\</returns>
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
