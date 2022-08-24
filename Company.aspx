<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Company.aspx.cs" Inherits="directory.WebForm2" %>

<%@ Register assembly="AjaxControlToolkit" namespace="AjaxControlToolkit" tagprefix="cc1" %>

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
      </style>
<body bgcolor=Beige>
    <form id="form1" runat="server">
       <div style="text-align: center; color: #FFFFFF; font-size: xx-large; background-color: #006699">

            <asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="~/Main.aspx" 
                style="color:white;" ToolTip="Перейти на главную страницу">Справочник сотрудников</asp:HyperLink>
          
        </div>
        <br />
    <table  style="width: 90%">
    <tr>
    <td>
    <asp:TreeView ID="TreeView1" runat="server" 
            onselectednodechanged="TreeView1_SelectedNodeChanged" 
            OnClientClick="aspnetForm.target ='_blank';" Font-Size="Large" ShowLines="True">
        <Nodes>
            <asp:TreeNode Text="1" Value="Name of organization">
                <asp:TreeNode Text="2" Value="Department1">
                </asp:TreeNode>
            </asp:TreeNode>
        </Nodes>
    </asp:TreeView>
    </td>
    <td>
    
    
        &nbsp;</td>
    </tr>
</table>
    
    
            <asp:HyperLink ID="HyperLink2" runat="server" width="98%" 
            style="text-align: center">©&quot;Информационные технологии&quot;</asp:HyperLink>
    </form>
</body>
</html>
