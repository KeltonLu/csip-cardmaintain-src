using System;
using System.Collections.Generic;
using System.Text;
using CSIPCardMaintain.EntityLayer;

namespace CSIPCardMaintain.BusinessRules
{
    public class BRJOBSTEPLOG : CSIPCommonModel.BusinessRules.BRBase<EntityJOBSTEPLOG>
    {
        /// <summary>
        /// Insert JOBSTEPLOG
        /// </summary>
        /// <param name="strJOBNAME">JOBNAME</param>
        /// <param name="strSTEP">step</param>
        /// <param name="strCMD">CMD</param>
        /// <returns>是否成功</returns>
        public static bool Insert(string strJOBNAME, string strSTEP, string strCMD)
        {
            EntityJOBSTEPLOG eJOBSTEPLOG = new EntityJOBSTEPLOG();

            eJOBSTEPLOG.JOBNAME = strJOBNAME;
            eJOBSTEPLOG.DT = DateTime.Now.ToString();
            eJOBSTEPLOG.STEP = strSTEP;
            eJOBSTEPLOG.CMD = strCMD;

            try
            {
                return BRJOBSTEPLOG.AddNewEntity(eJOBSTEPLOG);
            }
            catch (Exception ex)
            {
                BRJOBSTEPLOG.SaveLog(ex);
                return false;
            }
        }
    }
}
