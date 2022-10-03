//******************************************************************
//*  �@    �̡Gyangyu(rosicky)
//*  �\�໡���G�d�H��
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
    public class BRCPMAST4 : CSIPCommonModel.BusinessRules.BRBase<EntityCPMAST4>
    {

        /// <summary>
        /// �妸�פJCPMAST4��
        /// </summary>
        /// <param name="esetCPMAST4">CPMAST4�ɸ��</param>
        /// <returns>�O�_���\</returns>
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
