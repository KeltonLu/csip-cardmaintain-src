//******************************************************************
//*  作    者：yangyu(rosicky)
//*  功能說明：VD卡人匯入錯誤表
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
    public class BRCPMAST4_Err : CSIPCommonModel.BusinessRules.BRBase<EntityCPMAST4_Err>
    {
        /// <summary>
        /// 批次匯入CPMAST4_Err
        /// </summary>
        /// <param name="esetCPMAST4Err">CPMAST4_Err Entity集合</param>
        /// <returns>是否成功</returns>
        public static bool Insert(EntitySet<EntityCPMAST4_Err> esetCPMAST4Err)
        {
            try
            {
                using (OMTransactionScope ts = new OMTransactionScope())
                {
                    if (esetCPMAST4Err.Count > 0)
                    {
                        for (int i = 0; i < esetCPMAST4Err.Count; i++)
                        {
                            EntityCPMAST4_Err eCPMAST4Err = esetCPMAST4Err.GetEntity(i);

                            if (!BRCPMAST4_Err.AddNewEntity(eCPMAST4Err))
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
                BRCPMAST4_Err.SaveLog(ex);
                return false;
            }
            return true;
        }
    }
}
