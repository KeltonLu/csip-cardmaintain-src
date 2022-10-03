//******************************************************************
//*  作    者：占偉林(James)
//*  功能說明：下載文件
//*  創建日期：2009/12/08
//*  修改記錄：

//*<author>            <time>            <TaskID>                <desc>
//*******************************************************************
using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.IO;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using Framework.Common.Logging;

public partial class DownLoadFile : Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        string strServerFileName = this.Session["ServerFile"].ToString();
        string strClientFileName = this.Session["ClientFile"].ToString();
        if (strServerFileName.Length > 4)
        {
            switch (strServerFileName.Substring(strServerFileName.Length - 4, 4).ToLower())
            { 
                case ".xls":
                    // 輸出格式為Excel
                    Response.Clear();
                    Response.Buffer = true;
                    Session.CodePage = 950;
                    Response.ContentType = "application/vnd.ms-excel";
                    Response.AppendHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode(strClientFileName, System.Text.Encoding.UTF8));
                    break;
                case ".txt":
                    Response.Clear();
                    Response.Buffer = true;
                    Session.CodePage = 950;
                    Response.ContentType = "text/plain";
                    Response.AppendHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode(strClientFileName, System.Text.Encoding.UTF8));
                    break;
                case ".csv":
                    Response.Clear();
                    Response.Buffer = true;
                    Session.CodePage = 950;
                    Response.ContentType = "application/vnd.ms-excel";
                    Response.AppendHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode(strClientFileName, System.Text.Encoding.UTF8));
                    break;
                default:
                    Response.Clear();
                    Response.Buffer = true;
                    Session.CodePage = 950;
                    Response.ContentType = "application/vnd.ms-excel";
                    Response.AppendHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode(strClientFileName, System.Text.Encoding.UTF8));
                    break;
            }
            // 下載文件
            sendFile(strServerFileName);
            Response.End();
        }
    }

    /// <summary>
    /// 下載文件
    /// </summary>
    /// <param name="strFileName">文件名(含路徑)</param>
    private void sendFile(string strFileName)
    {
        try
        {
            if (!File.Exists(strFileName))
            {
                return;
            }

            Stream oStream = File.OpenRead(strFileName);
            Int32 nBytesRead = 0;
            Int32 nBufferSize = 32768;
            Byte[] oBuffer = new Byte[nBufferSize];
            while (true)
            {
                nBytesRead = oStream.Read(oBuffer, 0, nBufferSize);
                if (nBytesRead <= 0)
                {
                    break;
                }
                if (nBytesRead != nBufferSize)
                {
                    Byte[] oBuffer2 = new Byte[nBytesRead];
                    for (int intLoop = 0; intLoop < nBytesRead; intLoop++)
                    {
                        oBuffer2[intLoop] = oBuffer[intLoop];
                    }

                    Response.BinaryWrite(oBuffer2);
                }
                else
                {
                    Response.BinaryWrite(oBuffer);
                }
            }

            oStream.Close();
        }
        catch (Exception exp)
        {
            Logging.Log(exp, LogLayer.UI);
        }
    }
}
