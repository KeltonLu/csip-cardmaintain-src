using System;
using System.Text;
using System.Text.RegularExpressions;
using System.Data;
using System.Collections;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.IO;
using CSIPCommonModel.BaseItem;
using Framework.Common.Utility;
using Framework.Common.Message;
using Framework.Common.JavaScript;
using System.Data.SqlClient;
using CSIPCommonModel.BusinessRules;
using CSIPCommonModel.EntityLayer;

/// <summary>
/// Summary description for BaseHelper
/// </summary>
public sealed class BaseHelper
{
    #region GetScript
    /// <summary>
    /// ÓÃ»§Session¶ªÊ§Ö®ºóÌø×ªµ½HomePageÒ³Ãæ½Å±¾
    /// </summary>
    /// <param name="page"></param>
    /// <returns></returns>
    public static string GetScriptForUserSessionOut(Page page)
    {
        StringBuilder sbScript = new StringBuilder();
        sbScript.Append("alert('")
        .Append(MessageHelper.GetMessage("0040"))
        .Append("');")
        .Append("window.location.href = '")
        .Append(page.ResolveUrl("~/Default.aspx")).Append("';");
        return sbScript.ToString();
    }

    /// <summary>
    /// ¹Ø±Õ×Ô¼º²¢Ë¢ÐÂ¸¸Ò³Ãæ
    /// </summary>
    /// <param name="page"></param>
    /// <returns></returns>
    public static string GetScriptForUserSessionOut_CloseMe(Page page)
    {
        StringBuilder sbScript = new StringBuilder();
        sbScript.Append("if(window.opener != null && window.opener != undefined)")
        .Append("window.opener.location.reload();")
        .Append("window.close();");
        return sbScript.ToString();
    }

    public static string GetScriptForCloseMeAndGotoURL(Page page, string URL)
    {
        StringBuilder sbScript = new StringBuilder();
        sbScript.Append("if(window.opener != null && window.opener != undefined)")
        .Append("window.opener.location.replace('" + URL + "');")
        .Append("window.close();");
        return sbScript.ToString();
    }


    #endregion

    #region Set Control
    /// <summary>
    /// ÉèÖÃCancel°´Å¥ÌáÊ¾
    /// </summary>
    /// <param name="btnCancel"></param>
    public static void SetCancelBtn(Framework.WebControls.CustButton btnCancel)
    {
        btnCancel.ConfirmMsg = MessageHelper.GetMessage("0028");
    }


    /// <summary>
    /// Àò¨ú±±¥óÅã¥Ü­È
    /// </summary>
    /// <param name="ShowID"></param>
    public static string GetShowText(string ShowID)
    {
        return WebHelper.GetShowText(ShowID);
    }

    /// <summary>
    /// Åã¥ÜºÝ¥½«H®§
    /// ­×§ï°O¿ý¡G½Õ¾ãClientMsgShow»yªk by Ares Stanley 20211119
    /// </summary>
    /// <param name="ShowID"></param>
    public static string ClientMsgShow(string strMsgID)
    {
        //return "ClientMsgShow('" + MessageHelper.GetMessage(strMsgID) + "');";
        return "window.parent.postMessage({ func: 'ClientMsgShow', data: '" + MessageHelper.GetMessage(strMsgID) + "' }, '*');";
    }

    public static string GetScriptForWindowOpenURL(Page page, string URL)
    {
        StringBuilder sbScript = new StringBuilder();
        sbScript.Append("window.open('" + URL + "','','width='+(screen.availWidth-7)+',height='+(screen.availHeight-38)+',top=0,left=0,toolbar=no,menubar=no,scrollbars=yes,resizable=yes,location=no,status=no');");
        return sbScript.ToString();
    }

    public static void GetScriptForWindowClose(Control Page)
    {
        StringBuilder sbScript = new StringBuilder();
        sbScript.Append("alert('" + MessageHelper.GetMessage("00_00000000_037") + "');");
        sbScript.Append("window.close();");
        jsBuilder.RegScript(Page, sbScript.ToString());
    }

    public static void GetScriptForWindowErrorClose(Control Page)
    {
        StringBuilder sbScript = new StringBuilder();
        sbScript.Append("alert('" + MessageHelper.GetMessage("00_00000000_000") + "');");
        sbScript.Append("window.close();");
        jsBuilder.RegScript(Page, sbScript.ToString());
    }

    /// <summary>
    /// ¨ú±oJOB¸ê®Æ
    /// </summary>
    /// <param name="strFunctionKey">ÄÝ©ÊKEY</param>
    /// <param name="strMsgID">¿ù»~«H®§</param>
    /// <returns>DataTable</returns>
    public static DataTable GetJobData(string strFunctionKey, ref string strMsgID)
    {
        SqlCommand sqlcmd = new SqlCommand();
        sqlcmd.CommandText = @"SELECT RUN_SECONDS,RUN_MINUTES,RUN_HOURS,RUN_DAY_OF_MONTH,RUN_MONTH,RUN_DAY_OF_WEEK,EXEC_PROG, STATUS, RUN_USER_LDAPID,RUN_USER_LDAPPWD, RUN_USER_RACFID,RUN_USER_RACFPWD,MAIL_TO,DESCRIPTION,CHANGED_USER,CONVERT(varchar, CHANGED_TIME, 120 ) as  CHANGED_TIME,JOB_ID 
                                FROM  CSIP.dbo.M_AUTOJOB WHERE   FUNCTION_KEY= @FUNCTION_KEY ";

        sqlcmd.CommandType = CommandType.Text;
        SqlParameter parmKey = new SqlParameter("@FUNCTION_KEY", strFunctionKey);

        sqlcmd.Parameters.Add(parmKey);

        DataSet dstProperty = BRM_AUTOJOB.SearchOnDataSet(sqlcmd);
        if (dstProperty != null)
        {
            return dstProperty.Tables[0];
        }
        else
        {
            strMsgID = "00_00000000_000";
            return null;
        }
    }

    #endregion

    #region ¶×¤J¤åÀÉ®æ¦¡ÀË®Ö

    /// <summary>
    /// °O¿ý¿ù»~¦æ¼Æ¿é¥X
    /// </summary>
    /// <param name="intTemp">¦æ¯Á¤Þ</param>
    /// <param name="blnIsNote"> ¬O§_¤w°O¿ý</param>
    /// <param name="arrlErrorRow">¿ù»~¦æ¼Æ¼Æ²Õ</param>
    public static void AddErrorMsg(int intTemp, int intColumn, string strMsgID, ref ArrayList arrListMsg)
    {
        arrListMsg.Add(MessageHelper.GetMessage("00_01060000_006") + intTemp.ToString() + MessageHelper.GetMessage("00_01060000_008") + MessageHelper.GetMessage("00_01060000_006") + Convert.ToString(intColumn + 1) + MessageHelper.GetMessage(strMsgID));
    }

    /// <summary>
    /// ƒe®Ú¾Ú¦r²Å¦êªº¦r¸`ªø«×ºI¨ú¦r²Å¦ê
    /// </summary>
    /// <param name="strReadLine">ƒe¦r²Å¦ê¹ï¶H</param>
    /// <param name="begin">ƒe¶}©l¦ì¸m</param>
    /// <param name="length">ƒeºI¨úªø«×</param>
    /// <param name="nextBegin">ƒe¦r²Å¦ê¹ï¶H</param>
    /// <returns>int</returns>
    public static string GetSubstringByByte(string strReadLine, int begin, int length, out int nextBegin)
    {

        //string strTemp1 = strReadLine.Substring(begin, length + length - GetByteLength(strReadLine.Substring(begin, length)));
        string strTemp1 = SubStr(strReadLine,begin, length + length - GetByteLength(SubStr(strReadLine, begin, length)));

        //nextBegin = begin + strTemp1.Length;
        nextBegin = begin + GetByteLength(strTemp1);

        return strTemp1;

    }

    /// <summary>
    /// ƒe®Ú¾ÚÀq»{ªº½s½X¨ú±o¦r²Å¦êªº¦r¸`ªø«×
    /// </summary>
    /// <param name="text">ƒe¦r²Å¦ê¹ï¶H</param>
    /// <returns>int</returns>
    public static int GetByteLength(string text)
    {
        return System.Text.Encoding.Default.GetBytes(text).Length;
    }

    /// <summary>
    /// «Dunicode¤è¦¡substring
    /// </summary>
    /// <param name="strStr"></param>
    /// <param name="iStartIndex"></param>
    /// <param name="iLength"></param>
    /// <returns></returns>
    public static string SubStr(string strStr, int iStartIndex, int iLength)
    {
        Encoding l_Encoding = Encoding.GetEncoding("big5", new EncoderExceptionFallback(), new DecoderReplacementFallback(""));
        byte[] l_byte = l_Encoding.GetBytes(strStr);
        if (iLength <= 0)
            return "";
        //¨Ò­Yªø«×10
        //­Ya_StartIndex¶Ç¤J9 -> ok, 10 ->¤£¦æ
        if (iStartIndex + 1 > l_byte.Length)
            return "";
        else
        {
            //­Ya_StartIndex¶Ç¤J9 , a_Cnt ¶Ç¤J2 -> ¤£¦æ -> §ï¦¨ 9,1
            if (iStartIndex + iLength > l_byte.Length)
                iLength = l_byte.Length - iStartIndex;
        }
        return l_Encoding.GetString(l_byte, iStartIndex, iLength);
    }


    /// <summary>
    /// ƒe°O¿ý¶×¤J¤é§Ó
    /// </summary>
    /// <param name="eLUpload">ƒe¶×¤J¤é§Ó</param>
    /// <param name="eLUploadDetail">ƒe¶×¤J¿ù»~¤é§Ó</param>
    /// <param name="strMsgID">ƒe¿ù»~ID</param>
    /// <returns>int</returns>
    public static void LogUpload(EntityL_UPLOAD eLUpload, EntityL_UPLOAD_DETAIL eLUploadDetail, string strMsgID)
    {
        eLUploadDetail.FAIL_REASON = MessageHelper.GetMessage(strMsgID);
        BRL_UPLOAD.Add(eLUpload, eLUploadDetail, ref strMsgID);

    }

    /// <summary>
    /// ƒe°O¿ý¶×¤J¤é§Ó
    /// </summary>
    /// <param name="eLUploadDetail">ƒe¶×¤J¿ù»~¤é§Ó</param>
    /// <param name="intRow">ƒe¿ù»~¦æ¸¹</param>
    /// <param name="strMsg">ƒe¿ù»~«H®§</param>
    /// <returns>int</returns>
    public static void LogUpload(EntityL_UPLOAD_DETAIL eLUploadDetail, int intRow, string strMsg)
    {
        eLUploadDetail.FAIL_REC_NO = intRow.ToString();
        eLUploadDetail.FAIL_REASON = strMsg;

        BRL_UPLOAD_DETAIL.Add(eLUploadDetail, ref  strMsg);

    }

    /// <summary>
    /// ¶×¤JÀË®Ö
    /// </summary>
    /// <param name="strUserID"> ¥Î¤áID</param>
    /// <param name="strFunctionKey">¨t²ÎÅv­­</param>
    /// <param name="strUploadID"> ¶×¤J§@·~½s¸¹</param>
    /// <param name="dtmThisDate"> ¶×¤J§@·~®É¶¡</param>
    /// <param name="strUploadName"> ¶×¤J§@·~¦WºÙ</param>
    /// <param name="strFilePath">¤W¶Ç¤å¥ó¦a§}</param>
    /// <param name="intMax">³Ì¤jµ§¼Æ</param>
    /// <param name="arrListMsg">ÀË®Ö¦^¶Ç«H®§</param>
    /// <param name="strMsgID">¿ù»~«H®§ID</param>
    /// <param name="dtblBegin">ÀYµ§¼Æ¼Æ¾Ú</param>
    /// <param name="dtblEnd">§Àµ§¼Æ¼Æ¾Ú</param>
    /// <returns>DataTable</returns>
    public static DataTable UploadCheck(string strUserID, string strFunctionKey, string strUploadID, DateTime dtmThisDate, string strUploadName, string strFilePath, int intMax, ArrayList arrListMsg, ref string strMsgID, DataTable dtblBegin, DataTable dtblEnd)
    {
        EntityL_UPLOAD eLUpload = new EntityL_UPLOAD();

        //* ¶×¤J¤é§ÓÄæ¦ì½á­È
        eLUpload.CHANGED_USER = strUserID;
        eLUpload.FUNCTION_KEY = strFunctionKey;
        eLUpload.UPLOAD_ID = strUploadID;
        eLUpload.UPLOAD_NAME = strUploadName;
        eLUpload.UPLOAD_DATE = dtmThisDate;
        eLUpload.UPLOAD_STATUS = "N";
        eLUpload.FILE_NAME = "";

        EntityL_UPLOAD_DETAIL eLUploadDetail = new EntityL_UPLOAD_DETAIL();

        //* ¶×¤J¥¢±Ñ¤é§ÓÄæ¦ì½á­È
        eLUploadDetail.FUNCTION_KEY = strFunctionKey;
        eLUploadDetail.UPLOAD_ID = strUploadID;
        eLUploadDetail.UPLOAD_DATE = dtmThisDate;
        eLUploadDetail.FAIL_REC_NO = "";



        DataTable dtblUpload = new DataTable();

        #region  ÀÉ®×¦WºÙÀË®Ö

        if (Regex.Match(strFilePath, "[\u4E00-\u9FA5]+").Length > 0)
        {
            strMsgID = "00_01060000_000";

            LogUpload(eLUpload, eLUploadDetail, strMsgID);

            return dtblUpload;
        }
        #endregion

        #region  ÀÉ®×Ãþ«¬ÀË®Ö

        FileInfo file = new FileInfo(strFilePath);

        eLUpload.FILE_NAME = file.Name;

        DataTable dtblUploadCheck = null;

        //* §PÂ_ÀÉ®×¬O§_¦s¦b
        if (!file.Exists)
        {
            strMsgID = "00_01060000_002";

            LogUpload(eLUpload, eLUploadDetail, strMsgID);

            return dtblUpload;
        }
        else
        {
            try
            {
                dtblUploadCheck = BRM_UPLOAD_CHECK.Search(strFunctionKey, strUploadID);
            }
            catch
            {
                strMsgID = "00_00000000_000";

                LogUpload(eLUpload, eLUploadDetail, strMsgID);

                return dtblUpload;
            }

            //* §PÂ_¸Ó¶×¤JÀË®Ö¦³µLÃþ«¬§PÂ_¼Æ¾Ú
            if (dtblUploadCheck.Rows.Count > 0)
            {
                //* §PÂ_ÀÉ®×Ãþ«¬
                if (file.Extension.ToUpper() != dtblUploadCheck.Rows[0]["EXTEND_NAME"].ToString())
                {
                    strMsgID = "00_01060000_001";

                    LogUpload(eLUpload, eLUploadDetail, strMsgID);

                    return dtblUpload;
                }
            }
            else
            {
                strMsgID = "00_01060000_003";

                LogUpload(eLUpload, eLUploadDetail, strMsgID);

                return dtblUpload;
            }
        }
        #endregion

        int intBeginCount = int.Parse(dtblUploadCheck.Rows[0]["BEGIN_COUNT"].ToString());
        int intEndCount = int.Parse(dtblUploadCheck.Rows[0]["END_COUNT"].ToString());

        int intBeginColumn = int.Parse(dtblUploadCheck.Rows[0]["BEGIN_COLUMN"].ToString());
        int intEndColumn = int.Parse(dtblUploadCheck.Rows[0]["END_COLUMN"].ToString());

        #region  ¸ê®Æ®wÄæ¦ìÃþ«¬©w¸qÀË®Ö

        DataTable dtblUploadType = null;

        try
        {
            dtblUploadType = BRM_UPLOAD_TYPE.Search(strFunctionKey, strUploadID);
        }
        catch
        {
            strMsgID = "00_00000000_000";

            LogUpload(eLUpload, eLUploadDetail, strMsgID);

            return dtblUpload;
        }


        if (dtblUploadType.Rows.Count > 0)
        {
            //* ¥Í¦¨¿é¥XªíªºÄæ¦ì
            for (int i = 0; i < dtblUploadType.Rows.Count; i++)
            {
                DataColumn dcolUpload = new DataColumn(dtblUploadType.Rows[i]["FIELD_NAME"].ToString());

                dtblUpload.Columns.Add(dcolUpload);
            }
        }
        else
        {
            strMsgID = "00_01060000_003";

            LogUpload(eLUpload, eLUploadDetail, strMsgID);

            return dtblUpload;
        }


        #endregion

        int intTemp = 0;

        int intOut = 0;

        string strTemp = "";

        decimal decOut = 0;
        string strUpload = "";
        string strField = "";

        int intFieldLength = 0;

        int intDecimalDigits = 0;


        int intUploadTotalCount = 0;

        //*.CSV¤å¥ó¥ÎODBC°µÀË®Ö
        if (dtblUploadCheck.Rows[0]["EXTEND_NAME"].ToString().ToUpper() == ".CSV")
        {
            #region  ÀÉ®×µ§¼ÆÀË®Ö

            DataTable dtblCsv = null;

            try
            {
                dtblCsv = Function.CsvToDtbl(file.DirectoryName, file.Name, dtblUploadType.Rows.Count);
            }
            catch
            {
                strMsgID = "00_01060000_004";

                LogUpload(eLUpload, eLUploadDetail, strMsgID);

                return dtblUpload;
            }

            intUploadTotalCount = dtblCsv.Rows.Count;

            eLUpload.UPLOAD_TOTAL_COUNT = intUploadTotalCount - intBeginCount - intEndCount;

            //* ¸ê®Æ¦æ¼Æ¤j¤_15000,´£¥Ü¿ù»~
            if (intUploadTotalCount - intBeginCount - intEndCount > intMax)
            {
                strMsgID = "00_01060000_005";

                LogUpload(eLUpload, eLUploadDetail, strMsgID);

                return dtblUpload;
            }


            #endregion

            #region  ÀÉ®×Äæ¦ìÀË®Ö

            string strMessage = "";
            BRL_UPLOAD.Add(eLUpload, ref strMessage);

            //* ÀYµ§¼Æ¼Æ¾Ú
            for (int i = 0; i < intBeginColumn; i++)
            {
                dtblBegin.Columns.Add("begin" + i.ToString());
            }

            //* §Àµ§¼Æ¼Æ¾Ú
            for (int i = 0; i < intEndColumn; i++)
            {
                dtblEnd.Columns.Add("end" + i.ToString());
            }

            for (int j = 0; j < dtblCsv.Rows.Count; j++)
            {
                intTemp++;

                if (intTemp > intBeginCount && intTemp <= intUploadTotalCount - intEndCount)
                {
                    DataRow drowUpload = dtblUpload.NewRow();

                    //* ¸ê®Æ®w¤¤Äæ¦ìÀË®Ö­Ó¼Æ»P¤å¥ó¤¤ªº­Ó¼Æ¤£µ¥
                    if (dtblUploadType.Rows.Count > dtblCsv.Columns.Count)
                    {
                        dtblUpload.Rows.Add(drowUpload);

                        arrListMsg.Add(MessageHelper.GetMessage("00_01060000_006") + intTemp.ToString() + MessageHelper.GetMessage("00_01060000_007"));

                        //* ¸ê®Æ®w¤¤Äæ¦ìÀË®Ö­Ó¼Æ»P¤å¥ó¤¤ªº­Ó¼Æ¤£µ¥,°O¿ý¶iÀË®Ö¤é§Ó
                        LogUpload(eLUploadDetail, intTemp, MessageHelper.GetMessage("00_01060000_006") + intTemp.ToString() + MessageHelper.GetMessage("00_01060000_007"));
                    }
                    else
                    {
                        for (int i = 0; i < dtblUploadType.Rows.Count; i++)
                        {
                            strUpload = dtblCsv.Rows[j][i].ToString().Trim();

                            intFieldLength = int.Parse(dtblUploadType.Rows[i]["FIELD_LENGTH"].ToString());

                            intDecimalDigits = int.Parse(dtblUploadType.Rows[i]["DECIMAL_DIGITS"].ToString());

                            switch (dtblUploadType.Rows[i]["FIELD_TYPE"].ToString().ToUpper())
                            {
                                //* ¦r²ÅÃþ«¬
                                case "STRING":
                                    if (GetByteLength(strUpload) > intFieldLength)
                                    {
                                        AddErrorMsg(intTemp, i, "00_01060000_010", ref arrListMsg);

                                        //* Äæ¦ìªø«×¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                        LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                    }
                                    break;

                                //* ¾ã¼ÆÃþ«¬
                                case "INT":
                                    if (!int.TryParse(strUpload == "" ? "0" : strUpload, out intOut))
                                    {
                                        AddErrorMsg(intTemp, i, "00_01060000_009", ref arrListMsg);

                                        //* Äæ¦ìÃþ«¬¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                        LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                    }
                                    else
                                    {
                                        if (strUpload.Length > intFieldLength)
                                        {
                                            AddErrorMsg(intTemp, i, "00_01060000_010", ref arrListMsg);

                                            //* Äæ¦ìªø«×¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                            LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                        }
                                    }
                                    break;

                                //* ®É¶¡¤é´ÁÃþ«¬
                                case "DATETIME":
                                    strField = strUpload.Replace(" ", "").Replace("-", "").Replace("/", "").Replace(":", "");
                                    if (!int.TryParse(strField == "" ? "0" : strField, out intOut))
                                    {
                                        AddErrorMsg(intTemp, i, "00_01060000_009", ref arrListMsg);

                                        //* Äæ¦ìÃþ«¬¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                        LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                    }
                                    break;

                                //* ¼Æ¦rÃþ«¬
                                case "DECIMAL":
                                    if (!decimal.TryParse(strUpload == "" ? "0" : strUpload, out decOut))
                                    {
                                        AddErrorMsg(intTemp, i, "00_01060000_009", ref arrListMsg);

                                        //* Äæ¦ìÃþ«¬¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                        LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                    }
                                    else
                                    {
                                        if (strUpload.Split('.').Length > 1)
                                        {
                                            strField = strUpload.Split('.')[0];
                                            if (strField.Length > intFieldLength - intDecimalDigits - 1)
                                            {
                                                AddErrorMsg(intTemp, i, "00_01060000_011", ref arrListMsg);
                                                //* Äæ¦ì¾ã¼Æ¦ì¼Æ¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                                LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                            }
                                            else
                                            {
                                                strField = strUpload.Split('.')[1];

                                                if (strField.Length > intDecimalDigits)
                                                {
                                                    AddErrorMsg(intTemp, i, "00_01060000_012", ref arrListMsg);
                                                    //* Äæ¦ì¤p¼Æ¦ì¼Æ¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                                    LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (strUpload.Length > intFieldLength - intDecimalDigits - 1)
                                            {
                                                AddErrorMsg(intTemp, i, "00_01060000_011", ref arrListMsg);
                                                //* Äæ¦ì¾ã¼Æ¦ì¼Æ¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                                LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                            }
                                        }
                                    }

                                    break;


                                //* ¦Ê¤À¤ñÃþ«¬
                                case "PERCENT":
                                    strField = strUpload.Replace("%", "");

                                    if (!decimal.TryParse(strField == "" ? "0" : strField, out decOut))
                                    {
                                        AddErrorMsg(intTemp, i, "00_01060000_009", ref arrListMsg);

                                        //* Äæ¦ìÃþ«¬¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                        LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                    }
                                    else
                                    {
                                        if (strField.Split('.').Length > 1)
                                        {
                                            strTemp = strField.Split('.')[0];
                                            if (strTemp.Length > intFieldLength - intDecimalDigits - 2)
                                            {
                                                AddErrorMsg(intTemp, i, "00_01060000_011", ref arrListMsg);
                                                //* Äæ¦ì¾ã¼Æ¦ì¼Æ¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                                LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                            }
                                            else
                                            {
                                                strTemp = strField.Split('.')[1];

                                                if (strTemp.Length > intDecimalDigits)
                                                {
                                                    AddErrorMsg(intTemp, i, "00_01060000_012", ref arrListMsg);
                                                    //* Äæ¦ì¤p¼Æ¦ì¼Æ¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                                    LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (strField.Length > intFieldLength - intDecimalDigits - 2)
                                            {
                                                AddErrorMsg(intTemp, i, "00_01060000_011", ref arrListMsg);
                                                //* Äæ¦ì¾ã¼Æ¦ì¼Æ¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                                LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                            }
                                        }
                                    }

                                    break;

                            }
                            drowUpload[i] = strUpload;


                        }

                        dtblUpload.Rows.Add(drowUpload);
                    }
                }
                else if (intTemp <= intBeginCount)
                {
                    DataRow drowBegin = dtblBegin.NewRow();


                    for (int i = 0; i < dtblBegin.Columns.Count; i++)
                    {
                        if (dtblCsv.Rows[j][i] != null)
                            drowBegin[i] = dtblCsv.Rows[j][i].ToString().Trim();
                    }

                    dtblBegin.Rows.Add(drowBegin);
                }
                else
                {
                    DataRow drowEnd = dtblEnd.NewRow();

                    for (int i = 0; i < dtblEnd.Columns.Count; i++)
                    {
                        if (dtblCsv.Rows[j][i] != null)
                            drowEnd[i] = dtblCsv.Rows[j][i].ToString().Trim();
                    }

                    dtblEnd.Rows.Add(drowEnd);
                }
            }
            #endregion

            return dtblUpload;
        }
        else
        {

            #region  ÀÉ®×µ§¼ÆÀË®Ö
            StreamReader objStreamReader = null;
            //* Åª¨ú¤å¥ó,°O¿ý¦æ¼Æ
            try
            {
                objStreamReader = file.OpenText();

                while (objStreamReader.Peek() != -1)
                {
                    objStreamReader.ReadLine();
                    intUploadTotalCount++;
                }

                eLUpload.UPLOAD_TOTAL_COUNT = intUploadTotalCount - intBeginCount - intEndCount;
            }
            catch
            {
                strMsgID = "00_01060000_004";

                LogUpload(eLUpload, eLUploadDetail, strMsgID);

                return dtblUpload;
            }
            finally
            {
                objStreamReader.Close();
                file = null;
            }

            //* ¸ê®Æ¦æ¼Æ¤j¤_15000,´£¥Ü¿ù»~
            if (intUploadTotalCount - intBeginCount - intEndCount > intMax)
            {
                strMsgID = "00_01060000_005";

                LogUpload(eLUpload, eLUploadDetail, strMsgID);

                return dtblUpload;
            }

            #endregion

            #region  ÀÉ®×Äæ¦ìÀË®Ö

            try
            {
                string strMessage = "";
                BRL_UPLOAD.Add(eLUpload, ref strMessage);

                objStreamReader = new StreamReader(strFilePath, System.Text.Encoding.Default);

                string strString = "";


                string strSplit = dtblUploadCheck.Rows[0]["LIST_SEPARATOR"].ToString();

                #region ¦³¤À¹j²Å
                if (strSplit != "")
                {
                    //* ÀYµ§¼Æ¼Æ¾Ú
                    for (int i = 0; i < intBeginColumn; i++)
                    {
                        dtblBegin.Columns.Add("begin" + i.ToString());
                    }

                    //* §Àµ§¼Æ¼Æ¾Ú
                    for (int i = 0; i < intEndColumn; i++)
                    {
                        dtblEnd.Columns.Add("end" + i.ToString());
                    }

                    while (objStreamReader.Peek() != -1)
                    {
                        intTemp++;

                        strString = objStreamReader.ReadLine();

                        string[] strUploads = strString.Split(strSplit.ToCharArray());

                        if (intTemp > intBeginCount && intTemp <= intUploadTotalCount - intEndCount)
                        {
                            DataRow drowUpload = dtblUpload.NewRow();

                            //* ¸ê®Æ®w¤¤Äæ¦ìÀË®Ö­Ó¼Æ»P¤å¥ó¤¤ªº­Ó¼Æ¤£µ¥
                            if (dtblUploadType.Rows.Count > strUploads.Length)
                            {
                                dtblUpload.Rows.Add(drowUpload);

                                arrListMsg.Add(MessageHelper.GetMessage("00_01060000_006") + intTemp.ToString() + MessageHelper.GetMessage("00_01060000_007"));

                                //* ¸ê®Æ®w¤¤Äæ¦ìÀË®Ö­Ó¼Æ»P¤å¥ó¤¤ªº­Ó¼Æ¤£µ¥,°O¿ý¶iÀË®Ö¤é§Ó
                                LogUpload(eLUploadDetail, intTemp, MessageHelper.GetMessage("00_01060000_006") + intTemp.ToString() + MessageHelper.GetMessage("00_01060000_007"));
                            }
                            else
                            {
                                for (int i = 0; i < dtblUploadType.Rows.Count; i++)
                                {
                                    strUpload = strUploads[i].Trim();

                                    intFieldLength = int.Parse(dtblUploadType.Rows[i]["FIELD_LENGTH"].ToString());

                                    intDecimalDigits = int.Parse(dtblUploadType.Rows[i]["DECIMAL_DIGITS"].ToString());

                                    switch (dtblUploadType.Rows[i]["FIELD_TYPE"].ToString().ToUpper())
                                    {
                                        //* ¦r²ÅÃþ«¬
                                        case "STRING":
                                            if (GetByteLength(strUpload) > intFieldLength)
                                            {
                                                AddErrorMsg(intTemp, i, "00_01060000_010", ref arrListMsg);

                                                //* Äæ¦ìªø«×¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                                LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                            }
                                            break;

                                        //* ¾ã¼ÆÃþ«¬
                                        case "INT":
                                            if (!int.TryParse(strUpload == "" ? "0" : strUpload, out intOut))
                                            {
                                                AddErrorMsg(intTemp, i, "00_01060000_009", ref arrListMsg);

                                                //* Äæ¦ìÃþ«¬¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                                LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                            }
                                            else
                                            {
                                                if (strUpload.Length > intFieldLength)
                                                {
                                                    AddErrorMsg(intTemp, i, "00_01060000_010", ref arrListMsg);

                                                    //* Äæ¦ìªø«×¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                                    LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                                }
                                            }
                                            break;

                                        //* ®É¶¡¤é´ÁÃþ«¬
                                        case "DATETIME":
                                            strField = strUpload.Replace(" ", "").Replace("-", "").Replace("/", "").Replace(":", "");
                                            if (!int.TryParse(strField == "" ? "0" : strField, out intOut))
                                            {
                                                AddErrorMsg(intTemp, i, "00_01060000_009", ref arrListMsg);

                                                //* Äæ¦ìÃþ«¬¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                                LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                            }
                                            break;

                                        //* ¼Æ¦rÃþ«¬
                                        case "DECIMAL":
                                            if (!decimal.TryParse(strUpload == "" ? "0" : strUpload, out decOut))
                                            {
                                                AddErrorMsg(intTemp, i, "00_01060000_009", ref arrListMsg);

                                                //* Äæ¦ìÃþ«¬¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                                LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                            }
                                            else
                                            {
                                                if (strUpload.Split('.').Length > 1)
                                                {
                                                    strField = strUpload.Split('.')[0];
                                                    if (strField.Length > intFieldLength - intDecimalDigits - 1)
                                                    {
                                                        AddErrorMsg(intTemp, i, "00_01060000_011", ref arrListMsg);
                                                        //* Äæ¦ì¾ã¼Æ¦ì¼Æ¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                                        LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                                    }
                                                    else
                                                    {
                                                        strField = strUpload.Split('.')[1];

                                                        if (strField.Length > intDecimalDigits)
                                                        {
                                                            AddErrorMsg(intTemp, i, "00_01060000_012", ref arrListMsg);
                                                            //* Äæ¦ì¤p¼Æ¦ì¼Æ¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                                            LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    if (strUpload.Length > intFieldLength - intDecimalDigits - 1)
                                                    {
                                                        AddErrorMsg(intTemp, i, "00_01060000_011", ref arrListMsg);
                                                        //* Äæ¦ì¾ã¼Æ¦ì¼Æ¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                                        LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                                    }
                                                }
                                            }

                                            break;


                                        //* ¦Ê¤À¤ñÃþ«¬
                                        case "PERCENT":
                                            strField = strUpload.Replace("%", "");

                                            if (!decimal.TryParse(strField == "" ? "0" : strField, out decOut))
                                            {
                                                AddErrorMsg(intTemp, i, "00_01060000_009", ref arrListMsg);

                                                //* Äæ¦ìÃþ«¬¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                                LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                            }
                                            else
                                            {
                                                if (strField.Split('.').Length > 1)
                                                {
                                                    strTemp = strField.Split('.')[0];
                                                    if (strTemp.Length > intFieldLength - intDecimalDigits - 2)
                                                    {
                                                        AddErrorMsg(intTemp, i, "00_01060000_011", ref arrListMsg);
                                                        //* Äæ¦ì¾ã¼Æ¦ì¼Æ¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                                        LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                                    }
                                                    else
                                                    {
                                                        strTemp = strField.Split('.')[1];

                                                        if (strTemp.Length > intDecimalDigits)
                                                        {
                                                            AddErrorMsg(intTemp, i, "00_01060000_012", ref arrListMsg);
                                                            //* Äæ¦ì¤p¼Æ¦ì¼Æ¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                                            LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    if (strField.Length > intFieldLength - intDecimalDigits - 2)
                                                    {
                                                        AddErrorMsg(intTemp, i, "00_01060000_011", ref arrListMsg);
                                                        //* Äæ¦ì¾ã¼Æ¦ì¼Æ¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                                        LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                                    }
                                                }
                                            }

                                            break;

                                    }
                                    drowUpload[i] = strUpload;


                                }

                                dtblUpload.Rows.Add(drowUpload);
                            }
                        }
                        else if (intTemp <= intBeginCount)
                        {
                            DataRow drowBegin = dtblBegin.NewRow();


                            for (int i = 0; i < dtblBegin.Columns.Count; i++)
                            {
                                if (strUploads[i] != null)
                                    drowBegin[i] = strUploads[i];
                            }

                            dtblBegin.Rows.Add(drowBegin);
                        }
                        else
                        {
                            DataRow drowEnd = dtblEnd.NewRow();

                            for (int i = 0; i < dtblEnd.Columns.Count; i++)
                            {
                                if (strUploads[i] != null)
                                    drowEnd[i] = strUploads[i];
                            }

                            dtblEnd.Rows.Add(drowEnd);
                        }
                    }
                }
                #endregion
                #region µL¤À¹j²Å
                else
                {
                    //* ÀYµ§¼Æ¼Æ¾Ú
                    dtblBegin.Columns.Add("begin");

                    //* §Àµ§¼Æ¼Æ¾Ú
                    dtblEnd.Columns.Add("end");

                    int intRowTotal = 0;
                    //* ¨C¦æ¤¹³\ªºÁ`ªø«×
                    for (int i = 0; i < dtblUploadType.Rows.Count; i++)
                    {
                        intRowTotal = intRowTotal + Convert.ToInt32(dtblUploadType.Rows[i]["FIELD_LENGTH"].ToString());
                    }

                    while (objStreamReader.Peek() != -1)
                    {
                        intTemp++;
                        strString = objStreamReader.ReadLine();

                        if (intTemp > intBeginCount && intTemp <= intUploadTotalCount - intEndCount)
                        {

                            DataRow drowUpload = dtblUpload.NewRow();

                            if (GetByteLength(strString) < intRowTotal)
                            {
                                dtblUpload.Rows.Add(drowUpload);

                                arrListMsg.Add(MessageHelper.GetMessage("00_01060000_006") + intTemp.ToString() + MessageHelper.GetMessage("00_01060000_013"));

                                //* Äæ¦ìªø«×¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                            }
                            else
                            {
                                int intNextBegin = 0;
                                for (int i = 0; i < dtblUploadType.Rows.Count; i++)
                                {

                                    intFieldLength = int.Parse(dtblUploadType.Rows[i]["FIELD_LENGTH"].ToString());

                                    intDecimalDigits = int.Parse(dtblUploadType.Rows[i]["DECIMAL_DIGITS"].ToString());


                                    //*ºI¨ú»Ý­nÀË®ÖªºÄæ¦ì
                                    strUpload = GetSubstringByByte(strString, intNextBegin, intFieldLength, out intNextBegin).Trim();

                                    switch (dtblUploadType.Rows[i]["FIELD_TYPE"].ToString().ToUpper())
                                    {
                                        //* ¾ã¼ÆÃþ«¬
                                        case "INT":

                                            if (!int.TryParse(strUpload == "" ? "0" : strUpload, out intOut))
                                            {
                                                AddErrorMsg(intTemp, i, "00_01060000_009", ref arrListMsg);

                                                //* Äæ¦ìÃþ«¬¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                                LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                            }
                                            break;

                                        //* ®É¶¡¤é´ÁÃþ«¬
                                        case "DATETIME":
                                            if (!int.TryParse(strUpload.Replace(" ", "").Replace("-", "").Replace("/", "").Replace(":", "") == "" ? "0" : strUpload.Replace(" ", "").Replace("-", "").Replace("/", "").Replace(":", ""), out intOut))
                                            {
                                                AddErrorMsg(intTemp, i, "00_01060000_009", ref arrListMsg);

                                                //* Äæ¦ìÃþ«¬¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                                LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                            }
                                            break;

                                        //* ¼Æ¦rÃþ«¬
                                        case "DECIMAL":

                                            if (!decimal.TryParse(strUpload == "" ? "0" : strField, out decOut))
                                            {
                                                AddErrorMsg(intTemp, i, "00_01060000_009", ref arrListMsg);

                                                //* Äæ¦ìÃþ«¬¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                                LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                            }
                                            else
                                            {
                                                if (strUpload.Split('.').Length > 1)
                                                {
                                                    strField = strUpload.Split('.')[0];
                                                    if (strField.Length > intFieldLength - intDecimalDigits - 1)
                                                    {
                                                        AddErrorMsg(intTemp, i, "00_01060000_011", ref arrListMsg);
                                                        //* Äæ¦ì¾ã¼Æ¦ì¼Æ¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                                        LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                                    }
                                                    else
                                                    {
                                                        strField = strUpload.Split('.')[1];

                                                        if (strField.Length > intDecimalDigits)
                                                        {
                                                            AddErrorMsg(intTemp, i, "00_01060000_012", ref arrListMsg);
                                                            //* Äæ¦ì¤p¼Æ¦ì¼Æ¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                                            LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    if (strUpload.Length > intFieldLength - intDecimalDigits - 1)
                                                    {
                                                        AddErrorMsg(intTemp, i, "00_01060000_011", ref arrListMsg);
                                                        //* Äæ¦ì¾ã¼Æ¦ì¼Æ¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                                        LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                                    }
                                                }
                                            }

                                            break;

                                        //* ¦Ê¤À¤ñÃþ«¬
                                        case "PERCENT":
                                            strField = strUpload.Replace("%", "");

                                            if (!decimal.TryParse(strField == "" ? "0" : strField, out decOut))
                                            {
                                                AddErrorMsg(intTemp, i, "00_01060000_009", ref arrListMsg);

                                                //* Äæ¦ìÃþ«¬¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                                LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                            }
                                            else
                                            {
                                                if (strField.Split('.').Length > 1)
                                                {
                                                    strTemp = strField.Split('.')[0];
                                                    if (strTemp.Length > intFieldLength - intDecimalDigits - 2)
                                                    {
                                                        AddErrorMsg(intTemp, i, "00_01060000_011", ref arrListMsg);
                                                        //* Äæ¦ì¾ã¼Æ¦ì¼Æ¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                                        LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                                    }
                                                    else
                                                    {
                                                        strTemp = strField.Split('.')[1];

                                                        if (strTemp.Length > intDecimalDigits)
                                                        {
                                                            AddErrorMsg(intTemp, i, "00_01060000_012", ref arrListMsg);
                                                            //* Äæ¦ì¤p¼Æ¦ì¼Æ¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                                            LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    if (strField.Length > intFieldLength - intDecimalDigits - 2)
                                                    {
                                                        AddErrorMsg(intTemp, i, "00_01060000_011", ref arrListMsg);
                                                        //* Äæ¦ì¾ã¼Æ¦ì¼Æ¿ù»~,°O¿ý¶iÀË®Ö¤é§Ó
                                                        LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                                    }
                                                }
                                            }

                                            break;

                                    }
                                    drowUpload[i] = strUpload;

                                }
                                dtblUpload.Rows.Add(drowUpload);
                            }
                        }
                        else if (intTemp <= intBeginCount)
                        {
                            DataRow drowBegin = dtblBegin.NewRow();

                            drowBegin[0] = strString;

                            dtblBegin.Rows.Add(drowBegin);

                        }
                        else
                        {
                            DataRow drowEnd = dtblEnd.NewRow();

                            drowEnd[0] = strString;

                            dtblEnd.Rows.Add(drowEnd);
                        }

                    }
                }
                #endregion
            }
            catch
            {
                strMsgID = "00_01060000_004";

                LogUpload(eLUpload, eLUploadDetail, strMsgID);

                return dtblUpload;
            }
            finally
            {
                objStreamReader.Close();
            }

            #endregion

            return dtblUpload;
        }
    }

    #endregion ¶×¤J¤åÀÉ®æ¦¡ÀË®Ö
}
