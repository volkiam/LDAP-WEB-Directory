<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="LogViewer.aspx.cs" Inherits="directory.LogViewer" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="cc1" %>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">

<head id="Head1" runat="server">
    <title></title>
    <link rel="stylesheet" href="Styles.css">
</head>
  <style>
   h1 {
    background: #006699; /* Цвет фона */
    border: 1px solid #0404B4; /* Белая рамка */
    width: 98%;
    text-align: center;
      }
      #block
      {
          width: 1265px;
          text-align: left;
      }
      .style1
      {
          width: 91px;
          height: 45px;
      }
      .style2
      {
          width: 135px;
          height: 45px;
      }
      .style3
      {
          width: 15%;
          height: 45px;
      }
      .style4
      {
          height: 45px;
      }
      </style>
<body bgcolor="Beige">
    <form id="form1" runat="server">
       <div style="text-align: center; color: #FFFFFF; font-size: xx-large; background-color: #006699">

            <asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="~/Main.aspx" 
                style="color:white;" ToolTip="Перейти на главную страницу">Справочник сотрудников</asp:HyperLink>
          
        </div>
        <br />
        <table style="width: 90%">
        <tr>
        <td class="style3">
            <asp:Label ID="Label1" runat="server" Text="Выберите лог файл:   ">
            </asp:Label>
        </td>
        <td class="style1">
        <asp:DropDownList  ID="DropDownList1" runat="server" 
                onselectedindexchanged="DropDownList1_SelectedIndexChanged" 
                AutoPostBack="True" ViewStateMode="Enabled">
                </asp:DropDownList>
        </td>
        <td class="style4">
        
            <asp:Button ID="Button1" runat="server" onclick="Button1_Click" 
                Text="Обновить" />
        
        </td>
        <td class="style2">
            <asp:Label ID="Label2" runat="server" Text="Пользователь:"></asp:Label>
            <asp:TextBox ID="TextBox1" runat="server"></asp:TextBox>
        </td>
        <td class="style4" valign="bottom">
            <asp:Button ID="Button2" runat="server" Text="Фильтр" onclick="Button2_Click" 
                style="margin-left: 0px;" />
        </td>
        </tr>
        </table>
        <div>______________________________________________________________________________________________________________________________________________________</div>
        <asp:GridView ID="GridView1" runat="server"  BackColor="Beige" 
           BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" CellPadding="3" 
        width="98%" 
        style="text-align: center; margin-left: 10px" Font-Names="Times New Roman" 
        Font-Size="Medium" AllowPaging="True" AllowSorting="True" 
           onpageindexchanging="GridView1_PageIndexChanging" PageSize="20" 
           onsorting="GridView1_Sorting">
        <FooterStyle BackColor="Beige" ForeColor="#000066" />
        <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
            <PagerSettings FirstPageText="First" LastPageText="Last" 
                Mode="NumericFirstLast" />
        <PagerStyle BackColor="Beige" ForeColor="#000066" HorizontalAlign="Left" />
        <RowStyle ForeColor="#000066" />

       </asp:GridView>
        <br />

   
    
            <asp:HyperLink ID="HyperLink2" runat="server" width="98%" 
            style="text-align: center">©&quot;Информационные технологии&quot;</asp:HyperLink>
    </form>
</body>
</html>

