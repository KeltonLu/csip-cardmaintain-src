//******************************************************************
//*  �@    �̡Gyangyu(rosicky)
//*  �\�໡���G�d�H�פJ���~��
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
    public class BRCPMAST_Err : CSIPCommonModel.BusinessRules.BRBase<EntityCPMAST_Err>
    {
        /// <summary>
        /// �妸�פJCPMAST_Err
        /// </summary>
        /// <param name="esetCPMASTErr">CPMAST_Err Entity���X</param>
        /// <returns>�O�_���\</returns>
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
