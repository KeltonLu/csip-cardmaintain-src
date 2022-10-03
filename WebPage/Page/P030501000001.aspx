<%@ Page Language="C#" AutoEventWireup="true" CodeFile="P030501000001.aspx.cs" Inherits="Page_P030501000001" %>


<%@ Register Assembly="Framework.WebControls" Namespace="Framework.WebControls" TagPrefix="cc1" %>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%--新增列印功能 by Ares Stanley 20211213--%>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1"  runat="server">
     <title></title>
    <script src="../Common/Script/JavaScript.js"></script>

    <script src="../Common/Script/JQuery/jquery-1.3.2.min.js"></script>

    <script src="../Common/Script/JQuery/jquery-ui-1.7.min.js"></script>

    <script src="../Common/Script/JQuery/WINF_JQuery.js"></script>
    <script type="text/javascript"> 
    </script>
<link href="../App_Themes/Default/global.css" type="text/css" rel="stylesheet" />
</head>
<body class="workingArea">
        <form id="form1" runat="server" >
<asp:ScriptManager ID="ScriptManager1" EnablePageMethods="True" runat="server">
</asp:ScriptManager>
        <asp:UpdatePanel ID="UpdatePanel1" runat="server" UpdateMode="Conditional" ChildrenAsTriggers="true">
        <ContentTemplate>
    <table cellpadding="0" cellspacing="1" width="100%" >
    <tr class="itemTitle" >
					<td colspan="4">
					<li>
                        <cc1:CustLabel ID="lblTitle" runat="server" CurAlign="left" CurSymbol="&#163;" FractionalDigit="2" IsColon="False" IsCurrency="False" NeedDateFormat="False" NumBreak="0" NumOmit="0" SetBreak="False" SetOmit="False" ShowID="03_05010000_000" StickHeight="False"></cc1:CustLabel></li></td>
				</tr>
    <tr class="trOdd">
        <td style="text-align: right;width:15%">
            <cc1:CustLabel ID="lblData" runat="server" CurAlign="left" CurSymbol="&#163;" FractionalDigit="2" IsColon="True" IsCurrency="False" NeedDateFormat="False" NumBreak="0" NumOmit="0" SetBreak="False" SetOmit="False" ShowID="03_05010000_018" StickHeight="False"></cc1:CustLabel></td>
        <td style="text-align: left;width:85%">
         <cc1:DatePicker ID = "dpBeforeData" runat="server" ></cc1:DatePicker>
         ~
         <cc1:DatePicker ID = "dpEndData" runat="server" ></cc1:DatePicker>&nbsp;
         </td>
         </tr>
<tr class="itemTitle" align="center" >
    <td colspan="2">
        &nbsp;<cc1:CustButton ID="btnOK" runat="server" CssClass="smallButton" OnClick="btnOK_Click" ShowID="03_05010000_002"/>
        &nbsp;<cc1:CustButton ID="btnPrint" runat="server" CssClass="smallButton" OnClick="btnPrint_Click" ShowID="03_05010000_019"/>
    </td></tr>
    <tr><td colspan="2">
        <cc1:CustGridView ID="grvInpotLog" runat="server" AllowSorting="True" 
             PagerID="gpList"
            Width="100%" BorderWidth="0px" CellPadding="0" CellSpacing="1" BorderStyle="Solid" OnRowDataBound="grvInpotLog_RowDataBound" OnSelectedIndexChanging="grvInpotLog_SelectedIndexChanging" OnSelectedIndexChanged="grvInpotLog_SelectedIndexChanged"  >
            <Columns>
                <asp:TemplateField>
                    <edititemtemplate>
&nbsp;
</edititemtemplate>
                    <itemstyle width="5%" />
                    <headerstyle width="5%" />
                    <itemtemplate>
<asp:Label id="lblNo" runat="server" ></asp:Label> 
</itemtemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="INDate" >
                    <itemstyle   width="25%" />
                    <headerstyle width="25%" />
                </asp:BoundField>
                <asp:BoundField DataField="FileName" >
                    <itemstyle   width="25%" />
                    <headerstyle width="25%" />
                </asp:BoundField>
                <asp:BoundField DataField="RecordNums" >
                    <itemstyle   width="15%" horizontalalign="Right" />
                    <headerstyle width="15%" />
                </asp:BoundField>
                <asp:BoundField DataField="Active_Status" >
                    <itemstyle   width="15%" />
                    <headerstyle width="15%" />
                </asp:BoundField>
                <asp:BoundField DataField="ErrorNums" >
                    <itemstyle   width="15%" horizontalalign="Right" />
                    <headerstyle width="15%" />
                </asp:BoundField>
            </Columns>
            <RowStyle CssClass="Grid_Item"    Wrap="False" />
            <SelectedRowStyle CssClass="Grid_SelectedItem" />
            <HeaderStyle  CssClass="Grid_Header A"  Wrap="False" />
            <AlternatingRowStyle CssClass="Grid_AlternatingItem" Wrap="False" />
            <PagerSettings Visible="False"   />
            <EmptyDataRowStyle HorizontalAlign="Center"/>
        </cc1:CustGridView>
             <cc1:GridPager ID="gpList"  runat="server"  CustomInfoTextAlign="Right"  AlwaysShow="True" OnPageChanged="gpList_PageChanged">
                            </cc1:GridPager>
</td></tr></table>
        </ContentTemplate>
        </asp:UpdatePanel>   
    </form>
</body>
</html>
