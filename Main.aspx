<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Main.aspx.cs" Inherits="directory.WebForm1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">

<head runat="server">
    <title>Справочник Name of oraganization</title>
    <style type="text/css">

        .style1
        {
            color: #FF0000;
            font-size: small;
        }
        .style2
        {
            font-size: small;
        }
        .style3
        {
            width: 97%;
            height: 24px;
        }
        .style4
        {
            width: 102px;
        }
        .style5
        {
            width: 963px;
        }
        .style6
        {
            height: 10%;
            HorizontalAlign: "Left"
        }
    </style>
</head>

<body bgcolor="Beige">
    <form id="form1" runat="server">
        <div style="text-align: center; color: #FFFFFF; font-size: xx-large; background-color: #006699">

            <asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="~/Main.aspx" 
                style="color:white;" ToolTip="Перейти на главную страницу">Справочник сотрудников Name of oraganization</asp:HyperLink>
          
        </div>
        <table width = "99%" align="center">
        <tr>
        <td style="text-align:center">
            <asp:LinkButton ID="LinkButton1" runat="server" onclick="LinkButton1_Click" 
                Visible="False">Загрузить внешний телефонный справочник</asp:LinkButton>
        </td>
        </tr>
       </table>
<div align=right>
        <asp:Label ID="Label2" runat="server" Text="Label" Visible="False"></asp:Label>
        </div>
        <table width="98%">
        <td class="style3">
            <asp:Label ID="Label1" runat="server" 
            Text="Введите данные для поиска (ФИО, телефон, отдел, организация, адрес)" 
            Font-Names="Times New Roman" Font-Size="Large" ForeColor="Black"  
                style="margin-left: 26px" 
            ></asp:Label>
         </td>
        <td>
            <asp:LinkButton ID="LinkButton4" runat="server" onclick="LinkButton4_Click" 
                Visible="False" HorizontalAlign="Right">Logs</asp:LinkButton>
        </td>    
       
        </table>
           

        <asp:Panel ID="panSearch" runat="server" DefaultButton="Button3" Width="99%">
            <asp:TextBox ID="TextBox1" runat="server" Width="70%" style="margin-left: 10px"></asp:TextBox>
       <asp:Button ID="Button3" runat="server" onclick="Button3_Click" Text="Искать"  style="margin-left: 26px"/>
        <asp:Button ID="Button1" runat="server" onclick="Button1_Click" Text="Сброс" style="margin-left: 26px"/>
        
            
        
            <asp:LinkButton ID="LinkButton2" runat="server" onclick="LinkButton2_Click" 
                Visible="False" style="margin-left: 26px">Поиск по структуре</asp:LinkButton>
        
            
        
        </asp:Panel>   
       
        <table style="width: 95%">
        <tr>
        <td style="width: 90%">
        <span class="style1">&nbsp;&nbsp;&nbsp;(О всех ошибках, обнаруженных в данном справочнике,&nbsp; просьба сообщить&nbsp; по телефону вн. 
        2002 ( внеш. 222-222-22) и на электронную почту <a href="mailto:itdepart@domain.com">  <span class="style2">itdepart@domain.com</span></a>)</span><br />
        </td>
        <td style="width: 10%">
        <asp:Label ID="LabelRes" runat="server" Font-Italic="True" 
                Font-Names="Times New Roman" Font-Size="Small" ForeColor="Black"  
                Text="Найдено: " Visible="False"></asp:Label> 
        </td>  
        </tr>        
        </table>
    <asp:GridView ID="GridResults" runat="server"  AllowSorting="True"  OnSorting="Gridresults_Sorting" 
     BackColor="Beige" BorderColor="Black" BorderStyle="None" BorderWidth="1px" CellPadding="3" 
        width="98%" 
        style="text-align: center; margin-left: 10px" Font-Names="Times New Roman" 
        Font-Size="Medium" onrowdatabound="GridResults_RowDataBound" 
            onselectedindexchanged="GridResults_SelectedIndexChanged" 
            onrowcommand="GridResults_RowCommand">
         <Columns>
            <asp:ButtonField CommandName="ColumnClick" Visible="false" />

        </Columns>
        <FooterStyle BackColor="Beige" ForeColor="#000066" />
        <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
         <PagerSettings FirstPageText="First" LastPageText="Last" 
             Mode="NumericFirstLast" />
        <PagerStyle BackColor="Beige" ForeColor="#000066" HorizontalAlign="Left" />
        <RowStyle ForeColor="#000066" />
        <SelectedRowStyle BackColor="#669999" Font-Bold="True" ForeColor="White" />
        <SortedAscendingCellStyle BackColor="#F1F1F1" />
        <SortedAscendingHeaderStyle BackColor="#007DBB" />
        <SortedDescendingCellStyle BackColor="#CAC9C9" />
        <SortedDescendingHeaderStyle BackColor="#00547E" />
       
    </asp:GridView> 
    <br />

    <table width="98%">
    <tr>
 <td class="style4" ><asp:Button ID="Button4" runat="server" onclick="Button4_Click" Text="Назад" 
            style="font-family: Arial, Helvetica, sans-serif; font-size: small; margin-left: 26px" /></td>
 <td ><asp:Button ID="Button5" runat="server" onclick="Button5_Click" Text="Вперед" 
            style="font-family: Arial, Helvetica, sans-serif;  font-size: small; margin-left: 26px" /></td>
     <td class="style5" ><asp:Button ID="Button6" runat="server" onclick="Button6_Click" Text="Выводить на странице:" 
            
             style="font-family: Arial, Helvetica, sans-serif;  font-size: small; margin-left: 50px" />
         <asp:DropDownList ID="DropDownList1" runat="server" style="margin-left: 10px">
             <asp:ListItem>20</asp:ListItem>
             <asp:ListItem>50</asp:ListItem>
             <asp:ListItem>100</asp:ListItem>
             <asp:ListItem>2000</asp:ListItem>
         </asp:DropDownList>
         </td>
         <td Width="170px">
         <asp:Button ID="Button8" runat="server" onclick="Button8_Click" 
             Text="Выгрузить в файл" style="margin-right: 1px" Width="170px" 
             ToolTip="Выгрузка в файл всех пользователей данного справочника"/>
        </td>
    </tr>
    </table>
    <div id="Footer" style="
    position: relative;
    height: 1px;
    width: 99%;
    margin: auto;
    bottom: 0px;
    text-align: center;">
        <asp:HyperLink ID="HyperLink2" runat="server">©"Информационные технологии"</asp:HyperLink>
            </div>
    </form>
</body>

</html>
