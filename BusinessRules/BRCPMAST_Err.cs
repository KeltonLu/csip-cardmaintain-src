//******************************************************************
//*  作    者：yangyu(rosicky)
//*  功能說明：卡人匯入錯誤表
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
    public class BRCPMAST_Err : CSIPCommonModel.BusinessRules.BRBase<EntityCPMAST_Err>
    {
        /// <summary>
        /// 批次匯入CPMAST_Err
        /// </summary>
        /// <param name="esetCPMASTErr">CPMAST_Err Entity集合</param>
        /// <returns>是否成功</returns>
        public static bool Insert(EntitySet<EntityCPMAST_Err> esetCPMASTErr)
        {
            try
            {
                using (OMTransactionScope ts = new OMTransactionScope())
                {
                    if (esetCPMASTErr.Count > 0)
                    {
                        for (int i = 0; i < esetCPMASTErr.Count; i++)
                        {
                            EntityCPMAST_Err eCPMASTErr = esetCPMASTErr.GetEntity(i);

                            if (!BRCPMAST_Err.AddNewEntity(eCPMASTErr))
                            {
                                return false;
                            }
                        }
                    }
                    ts.Complete();
                }
            }
            catch (Exception ex)
            {
                BRCPMAST_Err.SaveLog(ex);
                return false;
            }
            return true;
        }
    }
}
