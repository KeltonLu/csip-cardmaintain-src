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
    public class BRCPMAST : CSIPCommonModel.BusinessRules.BRBase<EntityCPMAST>
    {
        /// <summary>
        /// 批次匯入CPMAST檔
        /// </summary>
        /// <param name="esetCPMAST">CPMAST檔資料</param>
        /// <returns>是否成功</returns>
        public static bool Insert(EntitySet<EntityCPMAST> esetCPMAST)
        {
            try
            {
                using (OMTransactionScope ts = new OMTransactionScope())
                {
                    if (esetCPMAST.Count > 0)
                    {
                        for (int i = 0; i < esetCPMAST.Count; i++)
                        {
                            EntityCPMAST eCPMAST = esetCPMAST.GetEntity(i);

                            if (!BRCPMAST.AddNewEntity(eCPMAST))
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
                BRCPMAST.SaveLog(ex);
                return false;
            }
            return true;
        }
    }
}
