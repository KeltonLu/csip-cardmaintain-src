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
    /// �û�Session��ʧ֮����ת��HomePageҳ��ű�
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
    /// �ر��Լ���ˢ�¸�ҳ��
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
    /// ����Cancel��ť��ʾ
    /// </summary>
    /// <param name="btnCancel"></param>
    public static void SetCancelBtn(Framework.WebControls.CustButton btnCancel)
    {
        btnCancel.ConfirmMsg = MessageHelper.GetMessage("0028");
    }


    /// <summary>
    /// ���������ܭ�
    /// </summary>
    /// <param name="ShowID"></param>
    public static string GetShowText(string ShowID)
    {
        return WebHelper.GetShowText(ShowID);
    }

    /// <summary>
    /// ��ܺݥ��H��
    /// �ק�O���G�վ�ClientMsgShow�y�k by Ares Stanley 20211119
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
    /// ���oJOB���
    /// </summary>
    /// <param name="strFunctionKey">�ݩ�KEY</param>
    /// <param name="strMsgID">���~�H��</param>
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

    #region �פJ���ɮ榡�ˮ�

    /// <summary>
    /// �O�����~��ƿ�X
    /// </summary>
    /// <param name="intTemp">�����</param>
    /// <param name="blnIsNote"> �O�_�w�O��</param>
    /// <param name="arrlErrorRow">���~��ƼƲ�</param>
    public static void AddErrorMsg(int intTemp, int intColumn, string strMsgID, ref ArrayList arrListMsg)
    {
        arrListMsg.Add(MessageHelper.GetMessage("00_01060000_006") + intTemp.ToString() + MessageHelper.GetMessage("00_01060000_008") + MessageHelper.GetMessage("00_01060000_006") + Convert.ToString(intColumn + 1) + MessageHelper.GetMessage(strMsgID));
    }

    /// <summary>
    /// �e�ھڦr�Ŧꪺ�r�`���׺I���r�Ŧ�
    /// </summary>
    /// <param name="strReadLine">�e�r�Ŧ��H</param>
    /// <param name="begin">�e�}�l��m</param>
    /// <param name="length">�e�I������</param>
    /// <param name="nextBegin">�e�r�Ŧ��H</param>
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
    /// �e�ھ��q�{���s�X���o�r�Ŧꪺ�r�`����
    /// </summary>
    /// <param name="text">�e�r�Ŧ��H</param>
    /// <returns>int</returns>
    public static int GetByteLength(string text)
    {
        return System.Text.Encoding.Default.GetBytes(text).Length;
    }

    /// <summary>
    /// �Dunicode�覡substring
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
        //�ҭY����10
        //�Ya_StartIndex�ǤJ9 -> ok, 10 ->����
        if (iStartIndex + 1 > l_byte.Length)
            return "";
        else
        {
            //�Ya_StartIndex�ǤJ9 , a_Cnt �ǤJ2 -> ���� -> �令 9,1
            if (iStartIndex + iLength > l_byte.Length)
                iLength = l_byte.Length - iStartIndex;
        }
        return l_Encoding.GetString(l_byte, iStartIndex, iLength);
    }


    /// <summary>
    /// �e�O���פJ���
    /// </summary>
    /// <param name="eLUpload">�e�פJ���</param>
    /// <param name="eLUploadDetail">�e�פJ���~���</param>
    /// <param name="strMsgID">�e���~ID</param>
    /// <returns>int</returns>
    public static void LogUpload(EntityL_UPLOAD eLUpload, EntityL_UPLOAD_DETAIL eLUploadDetail, string strMsgID)
    {
        eLUploadDetail.FAIL_REASON = MessageHelper.GetMessage(strMsgID);
        BRL_UPLOAD.Add(eLUpload, eLUploadDetail, ref strMsgID);

    }

    /// <summary>
    /// �e�O���פJ���
    /// </summary>
    /// <param name="eLUploadDetail">�e�פJ���~���</param>
    /// <param name="intRow">�e���~�渹</param>
    /// <param name="strMsg">�e���~�H��</param>
    /// <returns>int</returns>
    public static void LogUpload(EntityL_UPLOAD_DETAIL eLUploadDetail, int intRow, string strMsg)
    {
        eLUploadDetail.FAIL_REC_NO = intRow.ToString();
        eLUploadDetail.FAIL_REASON = strMsg;

        BRL_UPLOAD_DETAIL.Add(eLUploadDetail, ref  strMsg);

    }

    /// <summary>
    /// �פJ�ˮ�
    /// </summary>
    /// <param name="strUserID"> �Τ�ID</param>
    /// <param name="strFunctionKey">�t���v��</param>
    /// <param name="strUploadID"> �פJ�@�~�s��</param>
    /// <param name="dtmThisDate"> �פJ�@�~�ɶ�</param>
    /// <param name="strUploadName"> �פJ�@�~�W��</param>
    /// <param name="strFilePath">�W�Ǥ��a�}</param>
    /// <param name="intMax">�̤j����</param>
    /// <param name="arrListMsg">�ˮ֦^�ǫH��</param>
    /// <param name="strMsgID">���~�H��ID</param>
    /// <param name="dtblBegin">�Y���Ƽƾ�</param>
    /// <param name="dtblEnd">�����Ƽƾ�</param>
    /// <returns>DataTable</returns>
    public static DataTable UploadCheck(string strUserID, string strFunctionKey, string strUploadID, DateTime dtmThisDate, string strUploadName, string strFilePath, int intMax, ArrayList arrListMsg, ref string strMsgID, DataTable dtblBegin, DataTable dtblEnd)
    {
        EntityL_UPLOAD eLUpload = new EntityL_UPLOAD();

        //* �פJ��������
        eLUpload.CHANGED_USER = strUserID;
        eLUpload.FUNCTION_KEY = strFunctionKey;
        eLUpload.UPLOAD_ID = strUploadID;
        eLUpload.UPLOAD_NAME = strUploadName;
        eLUpload.UPLOAD_DATE = dtmThisDate;
        eLUpload.UPLOAD_STATUS = "N";
        eLUpload.FILE_NAME = "";

        EntityL_UPLOAD_DETAIL eLUploadDetail = new EntityL_UPLOAD_DETAIL();

        //* �פJ���Ѥ�������
        eLUploadDetail.FUNCTION_KEY = strFunctionKey;
        eLUploadDetail.UPLOAD_ID = strUploadID;
        eLUploadDetail.UPLOAD_DATE = dtmThisDate;
        eLUploadDetail.FAIL_REC_NO = "";



        DataTable dtblUpload = new DataTable();

        #region  �ɮצW���ˮ�

        if (Regex.Match(strFilePath, "[\u4E00-\u9FA5]+").Length > 0)
        {
            strMsgID = "00_01060000_000";

            LogUpload(eLUpload, eLUploadDetail, strMsgID);

            return dtblUpload;
        }
        #endregion

        #region  �ɮ������ˮ�

        FileInfo file = new FileInfo(strFilePath);

        eLUpload.FILE_NAME = file.Name;

        DataTable dtblUploadCheck = null;

        //* �P�_�ɮ׬O�_�s�b
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

            //* �P�_�ӶפJ�ˮ֦��L�����P�_�ƾ�
            if (dtblUploadCheck.Rows.Count > 0)
            {
                //* �P�_�ɮ�����
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

        #region  ��Ʈw��������w�q�ˮ�

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
            //* �ͦ���X�����
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

        //*.CSV����ODBC���ˮ�
        if (dtblUploadCheck.Rows[0]["EXTEND_NAME"].ToString().ToUpper() == ".CSV")
        {
            #region  �ɮ׵����ˮ�

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

            //* ��Ʀ�Ƥj�_15000,���ܿ��~
            if (intUploadTotalCount - intBeginCount - intEndCount > intMax)
            {
                strMsgID = "00_01060000_005";

                LogUpload(eLUpload, eLUploadDetail, strMsgID);

                return dtblUpload;
            }


            #endregion

            #region  �ɮ�����ˮ�

            string strMessage = "";
            BRL_UPLOAD.Add(eLUpload, ref strMessage);

            //* �Y���Ƽƾ�
            for (int i = 0; i < intBeginColumn; i++)
            {
                dtblBegin.Columns.Add("begin" + i.ToString());
            }

            //* �����Ƽƾ�
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

                    //* ��Ʈw������ˮ֭ӼƻP��󤤪��ӼƤ���
                    if (dtblUploadType.Rows.Count > dtblCsv.Columns.Count)
                    {
                        dtblUpload.Rows.Add(drowUpload);

                        arrListMsg.Add(MessageHelper.GetMessage("00_01060000_006") + intTemp.ToString() + MessageHelper.GetMessage("00_01060000_007"));

                        //* ��Ʈw������ˮ֭ӼƻP��󤤪��ӼƤ���,�O���i�ˮ֤��
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
                                //* �r������
                                case "STRING":
                                    if (GetByteLength(strUpload) > intFieldLength)
                                    {
                                        AddErrorMsg(intTemp, i, "00_01060000_010", ref arrListMsg);

                                        //* �����׿��~,�O���i�ˮ֤��
                                        LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                    }
                                    break;

                                //* �������
                                case "INT":
                                    if (!int.TryParse(strUpload == "" ? "0" : strUpload, out intOut))
                                    {
                                        AddErrorMsg(intTemp, i, "00_01060000_009", ref arrListMsg);

                                        //* ����������~,�O���i�ˮ֤��
                                        LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                    }
                                    else
                                    {
                                        if (strUpload.Length > intFieldLength)
                                        {
                                            AddErrorMsg(intTemp, i, "00_01060000_010", ref arrListMsg);

                                            //* �����׿��~,�O���i�ˮ֤��
                                            LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                        }
                                    }
                                    break;

                                //* �ɶ��������
                                case "DATETIME":
                                    strField = strUpload.Replace(" ", "").Replace("-", "").Replace("/", "").Replace(":", "");
                                    if (!int.TryParse(strField == "" ? "0" : strField, out intOut))
                                    {
                                        AddErrorMsg(intTemp, i, "00_01060000_009", ref arrListMsg);

                                        //* ����������~,�O���i�ˮ֤��
                                        LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                    }
                                    break;

                                //* �Ʀr����
                                case "DECIMAL":
                                    if (!decimal.TryParse(strUpload == "" ? "0" : strUpload, out decOut))
                                    {
                                        AddErrorMsg(intTemp, i, "00_01060000_009", ref arrListMsg);

                                        //* ����������~,�O���i�ˮ֤��
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
                                                //* ����Ʀ�ƿ��~,�O���i�ˮ֤��
                                                LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                            }
                                            else
                                            {
                                                strField = strUpload.Split('.')[1];

                                                if (strField.Length > intDecimalDigits)
                                                {
                                                    AddErrorMsg(intTemp, i, "00_01060000_012", ref arrListMsg);
                                                    //* ���p�Ʀ�ƿ��~,�O���i�ˮ֤��
                                                    LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (strUpload.Length > intFieldLength - intDecimalDigits - 1)
                                            {
                                                AddErrorMsg(intTemp, i, "00_01060000_011", ref arrListMsg);
                                                //* ����Ʀ�ƿ��~,�O���i�ˮ֤��
                                                LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                            }
                                        }
                                    }

                                    break;


                                //* �ʤ�������
                                case "PERCENT":
                                    strField = strUpload.Replace("%", "");

                                    if (!decimal.TryParse(strField == "" ? "0" : strField, out decOut))
                                    {
                                        AddErrorMsg(intTemp, i, "00_01060000_009", ref arrListMsg);

                                        //* ����������~,�O���i�ˮ֤��
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
                                                //* ����Ʀ�ƿ��~,�O���i�ˮ֤��
                                                LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                            }
                                            else
                                            {
                                                strTemp = strField.Split('.')[1];

                                                if (strTemp.Length > intDecimalDigits)
                                                {
                                                    AddErrorMsg(intTemp, i, "00_01060000_012", ref arrListMsg);
                                                    //* ���p�Ʀ�ƿ��~,�O���i�ˮ֤��
                                                    LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (strField.Length > intFieldLength - intDecimalDigits - 2)
                                            {
                                                AddErrorMsg(intTemp, i, "00_01060000_011", ref arrListMsg);
                                                //* ����Ʀ�ƿ��~,�O���i�ˮ֤��
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

            #region  �ɮ׵����ˮ�
            StreamReader objStreamReader = null;
            //* Ū�����,�O�����
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

            //* ��Ʀ�Ƥj�_15000,���ܿ��~
            if (intUploadTotalCount - intBeginCount - intEndCount > intMax)
            {
                strMsgID = "00_01060000_005";

                LogUpload(eLUpload, eLUploadDetail, strMsgID);

                return dtblUpload;
            }

            #endregion

            #region  �ɮ�����ˮ�

            try
            {
                string strMessage = "";
                BRL_UPLOAD.Add(eLUpload, ref strMessage);

                objStreamReader = new StreamReader(strFilePath, System.Text.Encoding.Default);

                string strString = "";


                string strSplit = dtblUploadCheck.Rows[0]["LIST_SEPARATOR"].ToString();

                #region �����j��
                if (strSplit != "")
                {
                    //* �Y���Ƽƾ�
                    for (int i = 0; i < intBeginColumn; i++)
                    {
                        dtblBegin.Columns.Add("begin" + i.ToString());
                    }

                    //* �����Ƽƾ�
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

                            //* ��Ʈw������ˮ֭ӼƻP��󤤪��ӼƤ���
                            if (dtblUploadType.Rows.Count > strUploads.Length)
                            {
                                dtblUpload.Rows.Add(drowUpload);

                                arrListMsg.Add(MessageHelper.GetMessage("00_01060000_006") + intTemp.ToString() + MessageHelper.GetMessage("00_01060000_007"));

                                //* ��Ʈw������ˮ֭ӼƻP��󤤪��ӼƤ���,�O���i�ˮ֤��
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
                                        //* �r������
                                        case "STRING":
                                            if (GetByteLength(strUpload) > intFieldLength)
                                            {
                                                AddErrorMsg(intTemp, i, "00_01060000_010", ref arrListMsg);

                                                //* �����׿��~,�O���i�ˮ֤��
                                                LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                            }
                                            break;

                                        //* �������
                                        case "INT":
                                            if (!int.TryParse(strUpload == "" ? "0" : strUpload, out intOut))
                                            {
                                                AddErrorMsg(intTemp, i, "00_01060000_009", ref arrListMsg);

                                                //* ����������~,�O���i�ˮ֤��
                                                LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                            }
                                            else
                                            {
                                                if (strUpload.Length > intFieldLength)
                                                {
                                                    AddErrorMsg(intTemp, i, "00_01060000_010", ref arrListMsg);

                                                    //* �����׿��~,�O���i�ˮ֤��
                                                    LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                                }
                                            }
                                            break;

                                        //* �ɶ��������
                                        case "DATETIME":
                                            strField = strUpload.Replace(" ", "").Replace("-", "").Replace("/", "").Replace(":", "");
                                            if (!int.TryParse(strField == "" ? "0" : strField, out intOut))
                                            {
                                                AddErrorMsg(intTemp, i, "00_01060000_009", ref arrListMsg);

                                                //* ����������~,�O���i�ˮ֤��
                                                LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                            }
                                            break;

                                        //* �Ʀr����
                                        case "DECIMAL":
                                            if (!decimal.TryParse(strUpload == "" ? "0" : strUpload, out decOut))
                                            {
                                                AddErrorMsg(intTemp, i, "00_01060000_009", ref arrListMsg);

                                                //* ����������~,�O���i�ˮ֤��
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
                                                        //* ����Ʀ�ƿ��~,�O���i�ˮ֤��
                                                        LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                                    }
                                                    else
                                                    {
                                                        strField = strUpload.Split('.')[1];

                                                        if (strField.Length > intDecimalDigits)
                                                        {
                                                            AddErrorMsg(intTemp, i, "00_01060000_012", ref arrListMsg);
                                                            //* ���p�Ʀ�ƿ��~,�O���i�ˮ֤��
                                                            LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    if (strUpload.Length > intFieldLength - intDecimalDigits - 1)
                                                    {
                                                        AddErrorMsg(intTemp, i, "00_01060000_011", ref arrListMsg);
                                                        //* ����Ʀ�ƿ��~,�O���i�ˮ֤��
                                                        LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                                    }
                                                }
                                            }

                                            break;


                                        //* �ʤ�������
                                        case "PERCENT":
                                            strField = strUpload.Replace("%", "");

                                            if (!decimal.TryParse(strField == "" ? "0" : strField, out decOut))
                                            {
                                                AddErrorMsg(intTemp, i, "00_01060000_009", ref arrListMsg);

                                                //* ����������~,�O���i�ˮ֤��
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
                                                        //* ����Ʀ�ƿ��~,�O���i�ˮ֤��
                                                        LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                                    }
                                                    else
                                                    {
                                                        strTemp = strField.Split('.')[1];

                                                        if (strTemp.Length > intDecimalDigits)
                                                        {
                                                            AddErrorMsg(intTemp, i, "00_01060000_012", ref arrListMsg);
                                                            //* ���p�Ʀ�ƿ��~,�O���i�ˮ֤��
                                                            LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    if (strField.Length > intFieldLength - intDecimalDigits - 2)
                                                    {
                                                        AddErrorMsg(intTemp, i, "00_01060000_011", ref arrListMsg);
                                                        //* ����Ʀ�ƿ��~,�O���i�ˮ֤��
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
                #region �L���j��
                else
                {
                    //* �Y���Ƽƾ�
                    dtblBegin.Columns.Add("begin");

                    //* �����Ƽƾ�
                    dtblEnd.Columns.Add("end");

                    int intRowTotal = 0;
                    //* �C�椹�\���`����
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

                                //* �����׿��~,�O���i�ˮ֤��
                                LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                            }
                            else
                            {
                                int intNextBegin = 0;
                                for (int i = 0; i < dtblUploadType.Rows.Count; i++)
                                {

                                    intFieldLength = int.Parse(dtblUploadType.Rows[i]["FIELD_LENGTH"].ToString());

                                    intDecimalDigits = int.Parse(dtblUploadType.Rows[i]["DECIMAL_DIGITS"].ToString());


                                    //*�I���ݭn�ˮ֪����
                                    strUpload = GetSubstringByByte(strString, intNextBegin, intFieldLength, out intNextBegin).Trim();

                                    switch (dtblUploadType.Rows[i]["FIELD_TYPE"].ToString().ToUpper())
                                    {
                                        //* �������
                                        case "INT":

                                            if (!int.TryParse(strUpload == "" ? "0" : strUpload, out intOut))
                                            {
                                                AddErrorMsg(intTemp, i, "00_01060000_009", ref arrListMsg);

                                                //* ����������~,�O���i�ˮ֤��
                                                LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                            }
                                            break;

                                        //* �ɶ��������
                                        case "DATETIME":
                                            if (!int.TryParse(strUpload.Replace(" ", "").Replace("-", "").Replace("/", "").Replace(":", "") == "" ? "0" : strUpload.Replace(" ", "").Replace("-", "").Replace("/", "").Replace(":", ""), out intOut))
                                            {
                                                AddErrorMsg(intTemp, i, "00_01060000_009", ref arrListMsg);

                                                //* ����������~,�O���i�ˮ֤��
                                                LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                            }
                                            break;

                                        //* �Ʀr����
                                        case "DECIMAL":

                                            if (!decimal.TryParse(strUpload == "" ? "0" : strField, out decOut))
                                            {
                                                AddErrorMsg(intTemp, i, "00_01060000_009", ref arrListMsg);

                                                //* ����������~,�O���i�ˮ֤��
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
                                                        //* ����Ʀ�ƿ��~,�O���i�ˮ֤��
                                                        LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                                    }
                                                    else
                                                    {
                                                        strField = strUpload.Split('.')[1];

                                                        if (strField.Length > intDecimalDigits)
                                                        {
                                                            AddErrorMsg(intTemp, i, "00_01060000_012", ref arrListMsg);
                                                            //* ���p�Ʀ�ƿ��~,�O���i�ˮ֤��
                                                            LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    if (strUpload.Length > intFieldLength - intDecimalDigits - 1)
                                                    {
                                                        AddErrorMsg(intTemp, i, "00_01060000_011", ref arrListMsg);
                                                        //* ����Ʀ�ƿ��~,�O���i�ˮ֤��
                                                        LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                                    }
                                                }
                                            }

                                            break;

                                        //* �ʤ�������
                                        case "PERCENT":
                                            strField = strUpload.Replace("%", "");

                                            if (!decimal.TryParse(strField == "" ? "0" : strField, out decOut))
                                            {
                                                AddErrorMsg(intTemp, i, "00_01060000_009", ref arrListMsg);

                                                //* ����������~,�O���i�ˮ֤��
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
                                                        //* ����Ʀ�ƿ��~,�O���i�ˮ֤��
                                                        LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                                    }
                                                    else
                                                    {
                                                        strTemp = strField.Split('.')[1];

                                                        if (strTemp.Length > intDecimalDigits)
                                                        {
                                                            AddErrorMsg(intTemp, i, "00_01060000_012", ref arrListMsg);
                                                            //* ���p�Ʀ�ƿ��~,�O���i�ˮ֤��
                                                            LogUpload(eLUploadDetail, intTemp, arrListMsg[arrListMsg.Count - 1].ToString());
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    if (strField.Length > intFieldLength - intDecimalDigits - 2)
                                                    {
                                                        AddErrorMsg(intTemp, i, "00_01060000_011", ref arrListMsg);
                                                        //* ����Ʀ�ƿ��~,�O���i�ˮ֤��
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

    #endregion �פJ���ɮ榡�ˮ�
}
