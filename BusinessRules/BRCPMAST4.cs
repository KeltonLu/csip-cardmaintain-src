//******************************************************************
//*  作    者：yangyu(rosicky)
//*  功能說明：卡人表
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
    public class BRCPMAST4 : CSIPCommonModel.BusinessRules.BRBase<EntityCPMAST4>
    {

        /// <summary>
        /// 批次匯入CPMAST4檔
        /// </summary>
        /// <param name="esetCPMAST4">CPMAST4檔資料</param>
        /// <returns>是否成功</returns>
        public static bool Insert(EntitySet<EntityCPMAST4> esetCPMAST4)
        {
            try
            {
                using (OMTransactionScope ts = new OMTransactionScope())
                {
                    if (esetCPMAST4.Count > 0)
                    {
                        for (int i = 0; i < esetCPMAST4.Count; i++)
                        {
                            EntityCPMAST4 eCPMAST4 = esetCPMAST4.GetEntity(i);

                            if (!BRCPMAST4.AddNewEntity(eCPMAST4))
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
                BRCPMAST4.SaveLog(ex);
                return false;
            }
            return true;
        }


    }
}
