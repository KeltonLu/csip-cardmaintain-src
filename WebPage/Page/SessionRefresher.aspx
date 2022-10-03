<%@ Page Language="C#" AutoEventWireup="true" CodeFile="SessionRefresher.aspx.cs" Inherits="Page_SessionRefresher" EnableEventValidation="false" ValidateRequest="false" %>

<%@ Register Assembly="Framework.WebControls" Namespace="Framework.WebControls" TagPrefix="cc1" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title></title>
</head>
<body>

    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server" AsyncPostBackErrorMessage="" AsyncPostBackTimeout="0" OnAsyncPostBackError="ScriptManager1_AsyncPostBackError">
        </asp:ScriptManager>
        <script language="javascript" type="text/javascript">
        Sys.WebForms.PageRequestManager.getInstance().add_endRequest(EndRequestHandler);
        function EndRequestHandler(sender, args)
        {
            if (args.get_error() != undefined)
            {
//                if ((args.get_response().get_statusCode() == '12007') || (args.get_response().get_statusCode() == '12029') || (args.get_response().get_statusCode() == '12002'))
//                {
//                    //Show a Message like 'Please make sure you are connected to internet';
//                    document.writeln("伺服器連接失敗");
                    window.onerror = false; //JS不做任何操作

//                }
            }
        }
        </script>
      
        <asp:Timer ID="FiveSeconds" Interval="300000" runat="server" />
        <asp:UpdatePanel ID="UpdatePanel1" runat="server">
            <ContentTemplate>
                <asp:Label ID="CurrentTime" runat="server" />
            </ContentTemplate>
        <Triggers>
            <asp:AsyncPostBackTrigger ControlID="FiveSeconds" EventName="Tick" />
        </Triggers>
    </asp:UpdatePanel>
    </form>

    
</body>
</html>
