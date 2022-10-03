<%@ Page Language="C#" AutoEventWireup="true" CodeFile="P030401010001.aspx.cs" Inherits="Page_P030401010001" %>


<%@ Register Assembly="Framework.WebControls" Namespace="Framework.WebControls" TagPrefix="cc1" %>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%--新增查詢功能 by Ares Stanley 20211108--%>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1"  runat="server">
     <title></title>
    <script src="../Common/Script/JavaScript.js"></script>

    <script src="../Common/Script/JQuery/jquery-1.3.2.min.js"></script>

    <script src="../Common/Script/JQuery/jquery-ui-1.7.min.js"></script>

    <script src="../Common/Script/JQuery/WINF_JQuery.js"></script>
     <script type="text/javascript"> 

    </script>
<link href="../App_Themes/Default/global.css" type="text/css" rel="stylesheet"/>
</head>
<body class="workingArea">
        <form id="form1" runat="server" >
<asp:ScriptManager ID="ScriptManager1" EnablePageMethods="True" runat="server" AsyncPostBackTimeout="120">
</asp:ScriptManager>
        <asp:UpdatePanel ID="UpdatePanel1" runat="server" UpdateMode="Conditional" ChildrenAsTriggers="true">
        <ContentTemplate>
    <table cellpadding="0" cellspacing="1" width="100%" >
    <tr class="itemTitle" >
					<td colspan="4">
					<li>
                        <cc1:CustLabel ID="lblTitle" runat="server" CurAlign="left" CurSymbol="&#163;" FractionalDigit="2" IsColon="False" IsCurrency="False" NeedDateFormat="False" NumBreak="0" NumOmit="0" SetBreak="False" SetOmit="False" ShowID="03_04010000_000" StickHeight="False"></cc1:CustLabel></li></td>
				</tr>
    <tr class="trOdd">
          <td style="text-align: right;width:25%">
            <cc1:CustLabel ID="lblID" runat="server" CurAlign="left" CurSymbol="&#163;" FractionalDigit="2" IsColon="True" IsCurrency="False" NeedDateFormat="False" NumBreak="0" NumOmit="0" SetBreak="False" SetOmit="False" ShowID="03_04010000_001" StickHeight="False"></cc1:CustLabel></td>
          <td style="text-align: left;width:75%">
          <cc1:CustCheckBox ID = "chkID" runat="server" />
          <cc1:CustTextBox ID="txtID" runat="server" MaxLength="16" Width="225px" InputType="Int"></cc1:CustTextBox></td>
    </tr>
         
   <tr class="trEven">
         <td style="text-align: right;width:25%">
            <cc1:CustLabel ID="lblPeople" runat="server" CurAlign="left" CurSymbol="&#163;" FractionalDigit="2" IsColon="True" IsCurrency="False" NeedDateFormat="False" NumBreak="0" NumOmit="0" SetBreak="False" SetOmit="False" ShowID="03_04010000_002" StickHeight="False"></cc1:CustLabel></td>
          <td style="text-align: left;width:75%">
          <cc1:CustCheckBox ID = "chkPeople" runat="server"  />
          <cc1:CustTextBox ID="txtPeople" runat="server" MaxLength="50" Width="225px"></cc1:CustTextBox></td>
   </tr>
         
   <tr class="trOdd">
        <td style="text-align: right;width:15%">
            <cc1:CustLabel ID="lblDate" runat="server" CurAlign="left" CurSymbol="&#163;" FractionalDigit="2" IsColon="True" IsCurrency="False" NeedDateFormat="False" NumBreak="0" NumOmit="0" SetBreak="False" SetOmit="False" ShowID="03_04010000_003" StickHeight="False"></cc1:CustLabel></td>
        <td style="text-align: left;width:85%">
         <cc1:CustCheckBox ID = "chkDate" runat="server"  />
         <cc1:DatePicker ID = "dpBeforeDate" runat="server" ></cc1:DatePicker>
         ~
         <cc1:DatePicker ID = "dpEndDate" runat="server" ></cc1:DatePicker>&nbsp;
         </td>
  </tr>
<tr class="itemTitle" align="center" >
    <td colspan="2">
		&nbsp;<cc1:CustButton ID="btnSearch" runat="server" CssClass="smallButton" onclick="btnSearch_Click" ShowID="03_04010000_004"/>
        &nbsp;<cc1:CustButton ID="btnOK" runat="server" CssClass="smallButton" onclick="btnOK_Click" ShowID="03_04010000_023"/></td></tr>
    </table>
            <table width="100%" border="0" cellpadding="0" cellspacing="0" id="Table1">
					<tr>
						<td colspan="20">
							<cc1:CustGridView ID="grvUserView" runat="server" AllowSorting="True" AllowPaging="True"
								PagerID="gpList" Width="100%" BorderWidth="0px" CellPadding="0" CellSpacing="1"
								BorderStyle="Solid">
								<RowStyle CssClass="Grid_Item" Wrap="True" />
								<SelectedRowStyle CssClass="Grid_SelectedItem" />
								<HeaderStyle CssClass="Grid_Header" Wrap="False" />
								<AlternatingRowStyle CssClass="Grid_AlternatingItem" Wrap="True" />
								<PagerSettings Visible="False" />
								<EmptyDataRowStyle HorizontalAlign="Center" />
								<Columns>
									<asp:BoundField DataField="CUST_ID">
										<ItemStyle Width="7%" HorizontalAlign="Center" />
									</asp:BoundField>
									<asp:BoundField DataField="FLD_NAME">
										<ItemStyle Width="14%" HorizontalAlign="Center" />
									</asp:BoundField>
									<asp:BoundField DataField="BEFOR_UPD">
										<ItemStyle Width="21%" HorizontalAlign="Center" />
									</asp:BoundField>
									<asp:BoundField DataField="AFTER_UPD">
										<ItemStyle Width="21%" HorizontalAlign="Center" />
									</asp:BoundField>
									<asp:BoundField DataField="MAINT_D">
										<ItemStyle Width="7%" HorizontalAlign="Center" />
									</asp:BoundField>
									<asp:BoundField DataField="MAINT_T">
										<ItemStyle Width="7%" HorizontalAlign="Center" />
									</asp:BoundField>
									<asp:BoundField DataField="USER_ID">
										<ItemStyle Width="7%" HorizontalAlign="Center" />
									</asp:BoundField>
								</Columns>
							</cc1:CustGridView>
						</td>
					</tr>
					<tr>
						<td>
							<cc1:GridPager ID="gpList" runat="server" AlwaysShow="True" CustomInfoTextAlign="Right"
								InputBoxStyle="height:15px" OnPageChanged="gpList_PageChanged">
							</cc1:GridPager>
						</td>
					</tr>
				</table>
        </ContentTemplate>
        </asp:UpdatePanel>   
    </form>
</body>
</html>
