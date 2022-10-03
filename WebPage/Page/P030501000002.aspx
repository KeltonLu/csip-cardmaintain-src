<%@ Page Language="C#" AutoEventWireup="true" CodeFile="P030501000002.aspx.cs" Inherits="Page_P030501000002" %>


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
                        <cc1:CustLabel ID="lblTitle" runat="server" CurAlign="left" CurSymbol="&#163;" FractionalDigit="2" IsColon="False" IsCurrency="False" NeedDateFormat="False" NumBreak="0" NumOmit="0" SetBreak="False" SetOmit="False" ShowID="03_05010000_008" StickHeight="False"></cc1:CustLabel></li></td>
				</tr>
    <tr class="trOdd">
        <td style="text-align: right;width:15%">
            <cc1:CustLabel ID="lblData" runat="server" CurAlign="left" CurSymbol="&#163;" FractionalDigit="2" IsColon="True" IsCurrency="False" NeedDateFormat="False" NumBreak="0" NumOmit="0" SetBreak="False" SetOmit="False" ShowID="03_05010000_009" StickHeight="False"></cc1:CustLabel></td><td style="text-align: left;width:35%"><cc1:CustTextBox ID="txtData" runat="server" MaxLength="20" Width="225px"></cc1:CustTextBox></td>
        <td style="text-align: right;width:15%">
         <cc1:CustLabel ID="lblFileName" runat="server" CurAlign="left" CurSymbol="&#163;" FractionalDigit="2" IsColon="True" IsCurrency="False" NeedDateFormat="False" NumBreak="0" NumOmit="0" SetBreak="False" SetOmit="False" ShowID="03_05010000_004" StickHeight="False"></cc1:CustLabel></td><td style="text-align: left;width:35%"><cc1:CustTextBox ID="txtFileName" runat="server" MaxLength="20" Width="225px"></cc1:CustTextBox>
         </td>
    </tr>
<tr class="itemTitle" align="center" >
    <td colspan="4">
        <cc1:CustButton ID="btnOK" runat="server" CssClass="smallButton" OnClick="btnOK_Click" ShowID="03_05010000_010"/>
        &nbsp;<cc1:CustButton ID="btnPrint" runat="server" CssClass="smallButton" OnClick="btnPrint_Click" ShowID="03_05010000_019"/>
    </td></tr>
    <tr><td colspan="4">
        <cc1:CustGridView ID="grvCPMASTErr" runat="server" AllowSorting="True" 
             PagerID="gpList"
            Width="100%" BorderWidth="0px" CellPadding="0" CellSpacing="1" BorderStyle="Solid" OnRowDataBound="grvCPMASTErr_RowDataBound"  >
            <Columns>
                <asp:TemplateField>
                    <edititemtemplate>

</edititemtemplate>
                    <itemstyle width="5%" />
                    <headerstyle width="5%" />
                    <itemtemplate>
                  
                   <asp:Label ID="lblNo" runat="server" ></asp:Label>
                  
</itemtemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="CUST_ID" >
                    <itemstyle   width="5%" />
                    <headerstyle width="5%" />
                </asp:BoundField>
                <asp:BoundField DataField="FLD_NAME" >
                    <itemstyle   width="10%" />
                    <headerstyle width="10%" />
                </asp:BoundField>
                <asp:BoundField DataField="BEFOR_UPD" >
                    <itemstyle   width="20%" />
                    <headerstyle width="20%" />
                </asp:BoundField>
                <asp:BoundField DataField="AFTER_UPD" >
                    <itemstyle   width="20%" />
                    <headerstyle width="20%" />
                </asp:BoundField>
                <asp:BoundField DataField="MAINT_D" >
                    <itemstyle   width="15%" />
                    <headerstyle width="15%" />
                </asp:BoundField>
                <asp:BoundField DataField="MAINT_T" >
                    <itemstyle   width="15%" />
                    <headerstyle width="15%" />
                </asp:BoundField>
                <asp:BoundField DataField="USER_ID" >
                    <itemstyle   width="15%" />
                    <headerstyle width="15%" />
                </asp:BoundField>
            </Columns>
            <RowStyle CssClass="Grid_Item"    />
            <SelectedRowStyle CssClass="Grid_SelectedItem" />
            <HeaderStyle  CssClass="Grid_Header"  Wrap="False" />
            <AlternatingRowStyle CssClass="Grid_AlternatingItem" />
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
