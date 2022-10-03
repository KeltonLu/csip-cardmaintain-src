using System;
using System.Collections.Generic;
using System.Text;
using CSIPCardMaintain.EntityLayer;

namespace CSIPCardMaintain.BusinessRules
{
    public class BRJOBLOG : CSIPCommonModel.BusinessRules.BRBase<EntityJOBLOG>
    {
        /// <summary>
        /// 如果今天還在執行中 或 執行成功了 或 因為其他問題結束了.今天就不再執行
        /// </summary>
        /// <param name="strJOBNAME">strJOBNAME</param>
        /// <param name="strJOBDATE">strJOBDATE</param>
        /// <returns>是否在執行中 或 執行成功了 或 因為其他問題結束了</returns>
        public static bool Select(string strJOBNAME, string strJOBDATE)
        {
            EntityJOBLOG eJOBLOG = new EntityJOBLOG();

            string strCon = @"(EXECUTESTATUS = '執行中' or EXECUTESTATUS = '成功' or EXECUTESTATUS = '結束') and 
                              JOBNAME= '" + strJOBNAME + @"' and EXECBDATE='" + strJOBDATE + @"'";

            try
            {
                return eJOBLOG.FillSelf(strCon);
            }
            catch (Exception ex)
            {
                BRJOBLOG.SaveLog(ex);
                return false;
            }
        }

        /// <summary>
        /// Insert JOBLOG
        /// </summary>
        /// <param name="JOBNAME">JOBNAME</param>
        /// <returns>是否成功</returns>
        public static bool Insert(string strJOBNAME)
        {
            EntityJOBLOG eJOBLOG = new EntityJOBLOG();

            eJOBLOG.JOBNAME = strJOBNAME;
            eJOBLOG.EXECBDATE = DateTime.Now.ToString("yyyyMMdd");
            eJOBLOG.EXECBTIME = DateTime.Now.ToString("HHmmss");
            eJOBLOG.EXECUTESTATUS = "執行中";
            eJOBLOG.EXECUSER = "system";

            try
            {
                return BRJOBLOG.AddNewEntity(eJOBLOG);
            }
            catch (Exception ex)
            {
                BRJOBLOG.SaveLog(ex);
                return false;
            }
        }

        /// <summary>
        /// Update JOBLOG
        /// </summary>
        /// <param name="strJOBNAME">JOBNAME</param>
        /// <param name="strCSTATUS">狀態</param>
        /// <param name="strMEMO">memo</param>
        /// <returns>是否成功</returns>
        public static bool Update(string strJOBNAME, string strCSTATUS, string strMEMO)
        {
            EntityJOBLOG eJOBLOG = new EntityJOBLOG();

            eJOBLOG.EXECMEMO = strMEMO;
            eJOBLOG.EXECEDATE = DateTime.Now.ToString("yyyyMMdd");
            eJOBLOG.EXECETIME = DateTime.Now.ToString("HHmmss");
            eJOBLOG.EXECUTESTATUS = strCSTATUS;

            string strCon = @"JOBNAME = '" + strJOBNAME + @"' and EXECUTESTATUS = '執行中' and 
                             ExecBDate = '" + DateTime.Now.ToString("yyyyMMdd") + @"'";

            try
            {
                return BRJOBLOG.UpdateEntityByCondition(eJOBLOG, strCon, EntityJOBLOG.M_EXECMEMO, EntityJOBLOG.M_EXECEDATE, EntityJOBLOG.M_EXECETIME, EntityJOBLOG.M_EXECUTESTATUS);
            }
            catch (Exception ex)
            {
                BRJOBLOG.SaveLog(ex);
                return false;
            }
        }


    }
}
