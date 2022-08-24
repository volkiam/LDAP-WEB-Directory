<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Users.aspx.cs" Inherits="directory.Users" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">



<head id="Head1" runat="server">
    <title>Справочник Name of organization</title>

    </head>
<body bgcolor="Beige">

<form id="form1" runat="server">
        <div style="text-align: center; color: #FFFFFF; font-size: xx-large; background-color: #006699">

            <asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="~/Main.aspx" 
                style="color:white;" ToolTip="Перейти на главную страницу">Справочник сотрудников Name of organization</asp:HyperLink>
          
        </div>
<br />
    <br />

    <asp:GridView ID="GridResults_Users" runat="server"  AllowSorting="True" 
     BackColor="Beige" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" CellPadding="3" 
        width="99%" 
        style="text-align: center; margin-left: 10px" Font-Names="Times New Roman" 
        Font-Size="Medium" onrowcommand="GridResults_Users_RowCommand" 
            onrowdatabound="GridResults_RowDataBound" 
            onselectedindexchanged="GridResults_Users_SelectedIndexChanged" 
            onsorting="GridResults_Users_Sorting">
        <Columns>
            <asp:ButtonField CommandName="ColumnClick" Visible="false" />
        </Columns>
        <FooterStyle BackColor="White" ForeColor="#000066" />
        <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
        <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" />
        <RowStyle ForeColor="#000066" />
        <SelectedRowStyle BackColor="#669999" Font-Bold="True" ForeColor="White" />
        <SortedAscendingCellStyle BackColor="#F1F1F1" />
        <SortedAscendingHeaderStyle BackColor="#007DBB" />
        <SortedDescendingCellStyle BackColor="#CAC9C9" />
        <SortedDescendingHeaderStyle BackColor="#00547E" />
       
      
    </asp:GridView> 

    <div id="Footer1" runat="server"  style="
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

