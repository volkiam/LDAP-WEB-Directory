using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.DirectoryServices;

namespace directory
{
    public partial class Users : System.Web.UI.Page
    {

        public const string LDAPuser = "user@domain";
        public const string LDAPpass = "12345678"; // да я в курсе что так не надо

        public static string sDisplayName;
        public static string AuthName;

        System.Data.DataTable DTBL_MAIN = null;
        System.Data.DataTable DTBL_MAIN3 = null;
        System.Data.DataTable DTBL_MAIN_FIND = null;

    //____________________________________________________________________________________________________________________//

        protected void Page_Load(object sender, EventArgs e)
        {
            AuthName = Page.User.Identity.Name;
            AuthName = AuthName.Substring(AuthName.IndexOf(@"\") + 1, AuthName.Length - AuthName.IndexOf(@"\") - 1);


            DirectoryEntry de = new DirectoryEntry("WinNT://" + HttpContext.Current.User.Identity.Name.Replace("\\", "/"));
            string sTempDisplayName = de.Properties["FullName"].Value.ToString();
            sDisplayName = sTempDisplayName;

            Session["DisplayName"] = sDisplayName;
            Session["AuthName"] = AuthName;

            if (!IsPostBack)
            {
                Session["UserTiTleNapr"] = "DESK";

                Session["TiTleNapr"] = "DESK";

                Session["UserSortnapravl"] = "DESK";

                Session["UserSortfield"] = "*";
            }

            if (Session["MyResults"] != null)
            {
                GridResults_Users.DataSource = Session["MyResults"];
  

                GridResults_Users.DataBind(); 
               
                if (GridResults_Users.Rows.Count < 20)
                {
                    Footer1.Style["position"] = "absolute";
                    Footer1.Style["height"] = "20";
                }
                else
                {
                    Footer1.Style["position"] = "relative";
                    Footer1.Style["height"] = "1";
                }

            }

        }



        private DataTable findLDAPuser(string lCompany, string lDepartment)
        {

            // Create a new DataTable.


            DTBL_MAIN_FIND = new DataTable("Results");
            // Create new DataColumn, set DataType, 
            // ColumnName and add to DataTable.  

            DataColumn workCol = DTBL_MAIN_FIND.Columns.Add("№", typeof(String));
            workCol.AllowDBNull = true;
            workCol.Unique = false;
            DTBL_MAIN_FIND.Columns.Add("ФИО", typeof(String));
            DTBL_MAIN_FIND.Columns.Add("Телефон", typeof(String));
            DTBL_MAIN_FIND.Columns.Add("IP телефон", typeof(String));
            DTBL_MAIN_FIND.Columns.Add("Должность", typeof(String));
            DTBL_MAIN_FIND.Columns.Add("Отдел", typeof(String));
            DTBL_MAIN_FIND.Columns.Add("Организация", typeof(String));
            DTBL_MAIN_FIND.Columns.Add("Кабинет", typeof(String));
            DTBL_MAIN_FIND.Columns.Add("Адрес", typeof(String));
            DTBL_MAIN_FIND.Columns.Add("Почта", typeof(String));

            try
            {

                DirectoryEntry root = new DirectoryEntry("cityhall.voronezh-city.ru");
                root.Path = "LDAP://OU=Employees,DC=cityhall,DC=voronezh-city,DC=ru";
                DirectorySearcher search = new DirectorySearcher(root);
                root.Username = LDAPuser;
                root.Password = LDAPpass;


                //ss  search.Filter = "(&(objectClass=user)(objectCategory=person)(!userAccountControl:1.2.840.113556.1.4.803:=2))";

                //search.Sort.Direction = System.DirectoryServices.SortDirection.Ascending;

                if ((lCompany == "*") & (lDepartment == "*")) //если выбрать всех
                {
                    search.Filter = "(&(objectClass=user)(objectCategory=person)(!userAccountControl:1.2.840.113556.1.4.803:=2))";
                    SortOption option = new SortOption("company", System.DirectoryServices.SortDirection.Ascending);
                    search.Sort = option;
                }
                if ((lCompany != "*") & (lDepartment != "*")) //если выбрана организация и отдел
                {
                    search.Filter = "(&(objectClass=user)(objectCategory=person)(!userAccountControl:1.2.840.113556.1.4.803:=2)(department=*" + lDepartment + "*)(company=*" + lCompany + "*))";
                    SortOption option = new SortOption("displayname", System.DirectoryServices.SortDirection.Ascending);
                    search.Sort = option;
                }
                if ((lCompany != "*") & (lDepartment == "*"))  //если выбрана организация
                {
                    search.Filter = "(&(objectClass=user)(objectCategory=person)(!userAccountControl:1.2.840.113556.1.4.803:=2)(company=*" + lCompany + "*))";
                    SortOption option = new SortOption("department", System.DirectoryServices.SortDirection.Ascending);
                    search.Sort = option;
                }

                search.SizeLimit = 2000;

                search.PropertiesToLoad.Add("samaccountname");
                search.PropertiesToLoad.Add("displayname");
                search.PropertiesToLoad.Add("mail");
                search.PropertiesToLoad.Add("telephoneNumber");
                search.PropertiesToLoad.Add("company");
                search.PropertiesToLoad.Add("title");
                search.PropertiesToLoad.Add("department");
                search.PropertiesToLoad.Add("streetAddress");
                search.PropertiesToLoad.Add("physicalDeliveryOfficeName");
                search.PropertiesToLoad.Add("Ipphone");
                search.PropertiesToLoad.Add("sn");
                search.PropertiesToLoad.Add("givenname");

                string str = "";
                SearchResultCollection results = search.FindAll();

                if (results != null)
                {
                    int nom = 1;
                    int last_count = results.Count;
                    foreach (SearchResult result in results)
                    {
                        DataRow row = DTBL_MAIN_FIND.NewRow();

                        row["№"] = nom.ToString();
                        nom = nom + 1;
                        str = "";

                        if (result.Properties["streetAddress"].Count > 0)
                        {
                            row["Адрес"] = result.Properties["streetAddress"][0].ToString();
                        }

                        if (result.Properties["physicalDeliveryOfficeName"].Count > 0)
                        {
                            row["Кабинет"] = result.Properties["physicalDeliveryOfficeName"][0].ToString();
                        }


                        if (result.Properties["displayname"].Count > 0)
                        {
                            str = result.Properties["displayname"][0].ToString();
                            if (result.Properties["sn"].Count > 0)
                            {
                                str = result.Properties["sn"][0].ToString();
                            }
                            if (result.Properties["givenname"].Count > 0)
                            {
                                str = str + " " + result.Properties["givenname"][0].ToString();
                            }
                            row["ФИО"] = str;
                        }

                        if (result.Properties["department"].Count > 0)
                        {
                            row["Отдел"] = result.Properties["department"][0].ToString();
                        }

                        if (result.Properties["company"].Count > 0)
                        {
                            row["Организация"] = result.Properties["company"][0].ToString();
                        }

                        if (result.Properties["title"].Count > 0)
                        {
                            row["Должность"] = result.Properties["title"][0].ToString();
                        }

                        if (result.Properties["telephoneNumber"].Count > 0)
                        {
                            row["Телефон"] = result.Properties["telephoneNumber"][0].ToString();
                        }

                        if (result.Properties["Ipphone"].Count > 0)
                        {
                            row["IP телефон"] = result.Properties["Ipphone"][0].ToString();
                        }
                        if (str != "") DTBL_MAIN_FIND.Rows.Add(row);
                        //   }
                        // }
                    }
                }

            }
            catch (Exception ex)
            {
                string exp = ex.Message;
            }

            return DTBL_MAIN_FIND;
        }

        protected override void Render(HtmlTextWriter writer)
        {
            foreach (GridViewRow r in GridResults_Users.Rows)
            {
                if (r.RowType == DataControlRowType.DataRow)
                {
                    for (int columnIndex = 5; columnIndex <= 8; columnIndex++)  //
                    {
                        Page.ClientScript.RegisterForEventValidation(r.UniqueID + "$ctl00", columnIndex.ToString());
                    }
                    Page.ClientScript.RegisterForEventValidation(r.UniqueID + "$ctl00", "9");
                }

            }
            base.Render(writer);
        }

        
        protected void GridResults_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            //string onmouseoverStyle1 = "this.style.fontWeight='UltraBold'; this.style.color='Navy';";
            //string onmouseoutStyle1 = "this.style.fontWeight='normal'; this.style.color='#333333';";

            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                if (e.Row.RowType == DataControlRowType.DataRow)
                {
                    LinkButton _singleClickButton = (LinkButton)e.Row.Cells[0].Controls[0];
                    string _jsSingle = ClientScript.GetPostBackClientHyperlink(_singleClickButton, "");
                    // Add events to each editable cell
                    for (int columnIndex = 5; columnIndex <= 8; columnIndex++)  // for (int columnIndex = 0; columnIndex < e.Row.Cells.Count; columnIndex++)
                    {
                        // Add the column index as the event argument parameter
                        string js = _jsSingle.Insert(_jsSingle.Length - 2, columnIndex.ToString());
                        // Add this javascript to the onclick Attribute of the cell
                        e.Row.Cells[columnIndex].Attributes["onclick"] = js;
                        // Add a cursor style to the cells
                        e.Row.Cells[columnIndex].Attributes["style"] += "cursor:pointer;cursor:hand;this.style.textDecoration='underline';";

                        e.Row.Cells[columnIndex].Attributes["onmouseover"] += "this.style.textDecoration='underline';";
                        e.Row.Cells[columnIndex].Attributes["onmouseout"] += "this.style.textDecoration='none';";

                        e.Row.Cells[columnIndex].ForeColor = System.Drawing.ColorTranslator.FromHtml("#2D1CD0");
                        // e.Row.Cells[columnIndex].Font.Underline= true;
                        e.Row.Cells[columnIndex].BorderColor = System.Drawing.Color.Black;

                    }

                    string _jsSingle2 = ClientScript.GetPostBackClientHyperlink(_singleClickButton, "");
                    string js2 = _jsSingle2.Insert(_jsSingle2.Length - 2, "9");
                    // Add this javascript to the onclick Attribute of the cell
                    e.Row.Cells[9].Attributes["onclick"] = js2;
                    // Add a cursor style to the cells
                    e.Row.Cells[9].Attributes["style"] += "cursor:pointer;cursor:hand;this.style.textDecoration='underline';";
                    e.Row.Cells[9].ForeColor = System.Drawing.ColorTranslator.FromHtml("#2D1CD0");
                    e.Row.Cells[9].BorderColor = System.Drawing.Color.Black;
                    e.Row.Cells[9].Text = "<img src='Images/mail.png' height='25' width='35' onmouseover=javascript:this.src='images/mail1.png' onmouseout=javascript:this.src='images/mail.png' title='Нажмите для отправки письма пользователю'/>";

                }

            }


            if (e.Row.RowType == System.Web.UI.WebControls.DataControlRowType.DataRow)
            {

                // when mouse is over the row, save original color to new attribute, and change it to highlight color
                e.Row.Attributes.Add("onmouseover", "this.originalstyle=this.style.backgroundColor;this.style.backgroundColor='#EEFFAA';");

                // when mouse leaves the row, change the bg color to its original value   
                e.Row.Attributes.Add("onmouseout", "this.style.backgroundColor=this.originalstyle;");


            }
        }

        string smail;


        private string findLDAPuserMail(string sname)
        {

            try
            {

                string[] words = sname.Split(new char[] { ' ' });
                string sF = words[0];
                string sIO = words[1] + " " + words[2];
                DirectoryEntry root = new DirectoryEntry("cityhall.voronezh-city.ru");
                root.Path = "LDAP://OU=Employees,DC=cityhall,DC=voronezh-city,DC=ru";
                DirectorySearcher search = new DirectorySearcher(root);
                root.Username = LDAPuser;
                root.Password = LDAPpass;

                //ss  search.Filter = "(&(objectClass=user)(objectCategory=person)(!userAccountControl:1.2.840.113556.1.4.803:=2))";

                //search.Sort.Direction = System.DirectoryServices.SortDirection.Ascending;


                search.Filter = "(&(objectClass=user)(objectCategory=person)(!userAccountControl:1.2.840.113556.1.4.803:=2)(sn=" + sF + ")(givenname=" + sIO + "))";
                //;  search.Filter = "(&(objectClass=user)(objectCategory=person)(!userAccountControl:1.2.840.113556.1.4.803:=2)(company=*" + lCompany + "*))";
                search.SizeLimit = 10;

                search.PropertiesToLoad.Add("samaccountname");
                search.PropertiesToLoad.Add("displayname");
                search.PropertiesToLoad.Add("mail");
                search.PropertiesToLoad.Add("sn");
                search.PropertiesToLoad.Add("givenname");

                
                SearchResultCollection results = search.FindAll();

                if (results != null)
                {
                    
                    int last_count = results.Count;
                    foreach (SearchResult result in results)
                    {

                        if (result.Properties["mail"].Count > 0)
                        {
                            smail = result.Properties["mail"][0].ToString();
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                string exp = ex.Message;
            }

            return smail;
        }

        private DataTable findLDAPAddress(string sKab, string sAddress)
        {

            // Create a new DataTable.


            DTBL_MAIN3 = new DataTable("Results");
            // Create new DataColumn, set DataType, 
            // ColumnName and add to DataTable.  

            DataColumn workCol = DTBL_MAIN3.Columns.Add("№", typeof(String));
            workCol.AllowDBNull = true;
            workCol.Unique = false;
            DTBL_MAIN3.Columns.Add("ФИО", typeof(String));
            DTBL_MAIN3.Columns.Add("Телефон", typeof(String));
            DTBL_MAIN3.Columns.Add("IP телефон", typeof(String));
            DTBL_MAIN3.Columns.Add("Должность", typeof(String));
            DTBL_MAIN3.Columns.Add("Отдел", typeof(String));
            DTBL_MAIN3.Columns.Add("Организация", typeof(String));
            DTBL_MAIN3.Columns.Add("Кабинет", typeof(String));
            DTBL_MAIN3.Columns.Add("Адрес", typeof(String));
            DTBL_MAIN3.Columns.Add("Почта", typeof(String));

            try
            {

                DirectoryEntry root = new DirectoryEntry("cityhall.voronezh-city.ru");
                root.Path = "LDAP://OU=Employees,DC=cityhall,DC=voronezh-city,DC=ru";
                DirectorySearcher search = new DirectorySearcher(root);
                root.Username = LDAPuser;
                root.Password = LDAPpass;

                //ss  search.Filter = "(&(objectClass=user)(objectCategory=person)(!userAccountControl:1.2.840.113556.1.4.803:=2))";

                //search.Sort.Direction = System.DirectoryServices.SortDirection.Ascending;


                if (sKab == "*")
                {
                    search.Filter = "(&(objectClass=user)(objectCategory=person)(!userAccountControl:1.2.840.113556.1.4.803:=2)(streetAddress=" + sAddress + "))";
                    SortOption option = new SortOption("displayname", System.DirectoryServices.SortDirection.Ascending);
                    search.Sort = option;
                }
                else
                {
                    search.Filter = "(&(objectClass=user)(objectCategory=person)(!userAccountControl:1.2.840.113556.1.4.803:=2)(physicalDeliveryOfficeName=" + sKab + ")(streetAddress=" + sAddress + "))";
                    SortOption option = new SortOption("displayname", System.DirectoryServices.SortDirection.Ascending);
                    search.Sort = option;
                }



                search.SizeLimit = 2000;

                search.PropertiesToLoad.Add("samaccountname");
                search.PropertiesToLoad.Add("displayname");
                search.PropertiesToLoad.Add("mail");
                search.PropertiesToLoad.Add("telephoneNumber");
                search.PropertiesToLoad.Add("company");
                search.PropertiesToLoad.Add("title");
                search.PropertiesToLoad.Add("department");
                search.PropertiesToLoad.Add("streetAddress");
                search.PropertiesToLoad.Add("physicalDeliveryOfficeName");
                search.PropertiesToLoad.Add("Ipphone");
                search.PropertiesToLoad.Add("sn");
                search.PropertiesToLoad.Add("givenname");

                string str = "";
                SearchResultCollection results = search.FindAll();

                if (results != null)
                {
                    int nom = 1;
                    int last_count = results.Count;
                    foreach (SearchResult result in results)
                    {
                        DataRow row = DTBL_MAIN3.NewRow();

                        row["№"] = nom.ToString();
                        nom = nom + 1;
                        str = "";

                        if (result.Properties["streetAddress"].Count > 0)
                        {
                            row["Адрес"] = result.Properties["streetAddress"][0].ToString();
                        }

                        if (result.Properties["physicalDeliveryOfficeName"].Count > 0)
                        {
                            row["Кабинет"] = result.Properties["physicalDeliveryOfficeName"][0].ToString();
                        }


                        if (result.Properties["displayname"].Count > 0)
                        {
                            str = result.Properties["displayname"][0].ToString();
                            if (result.Properties["sn"].Count > 0)
                            {
                                str = result.Properties["sn"][0].ToString();
                            }
                            if (result.Properties["givenname"].Count > 0)
                            {
                                str = str + " " + result.Properties["givenname"][0].ToString();
                            }
                            row["ФИО"] = str;
                        }

                        if (result.Properties["department"].Count > 0)
                        {
                            row["Отдел"] = result.Properties["department"][0].ToString();
                        }

                        if (result.Properties["company"].Count > 0)
                        {
                            row["Организация"] = result.Properties["company"][0].ToString();
                        }

                        if (result.Properties["title"].Count > 0)
                        {
                            row["Должность"] = result.Properties["title"][0].ToString();
                        }

                        if (result.Properties["telephoneNumber"].Count > 0)
                        {
                            row["Телефон"] = result.Properties["telephoneNumber"][0].ToString();
                        }

                        if (result.Properties["Ipphone"].Count > 0)
                        {
                            row["IP телефон"] = result.Properties["Ipphone"][0].ToString();
                        }
                        //if (result.Properties["mail"].Count > 0)
                        //{
                        //    row["Почта"] = result.Properties["mail"][0].ToString();
                        //}
                        row["Почта"] = "mail";

                        if (str != "") DTBL_MAIN3.Rows.Add(row);
                        //   }
                        // }
                    }
                }

            }
            catch (Exception ex)
            {
                string exp = ex.Message;
            }

            return DTBL_MAIN3;
        }



        public void WriteLog(string sLog)
        {
            try
            {
                DateTime dt = DateTime.Now;
                string sFileName = Server.MapPath(@"Logs\log_" + dt.ToString("MMddyyyy") + ".txt");
                String strDate = dt.ToString("MM/dd/yyyy HH:mm:ss");


                if (!File.Exists(sFileName))
                    File.Create(sFileName).Close();



                using (StreamWriter sw = new StreamWriter(sFileName, true))
                {

                    sw.WriteLine(strDate + "| DisplayName: " + Session["DisplayName"] + " | Login: " + Session["AuthName"] + " | Text: " + sLog);
                }

            }
            catch
            { }
        }


        protected void GridResults_Users_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            String sCompany = "";
            String sDepartment = "";
           

            // Create a new DataTable.
            DTBL_MAIN3 = new DataTable("Results");
            // Create new DataColumn, set DataType, 
            // ColumnName and add to DataTable.  

            //DataColumn workCol = DTBL_MAIN3.Columns.Add("№", typeof(String));
            //workCol.AllowDBNull = true;
            //workCol.Unique = false;
            DTBL_MAIN3.Columns.Add("ФИО", typeof(String));
            DTBL_MAIN3.Columns.Add("Телефон", typeof(String));
            DTBL_MAIN3.Columns.Add("IP телефон", typeof(String));
            DTBL_MAIN3.Columns.Add("Должность", typeof(String));
            DTBL_MAIN3.Columns.Add("Отдел", typeof(String));
            DTBL_MAIN3.Columns.Add("Организация", typeof(String));
            DTBL_MAIN3.Columns.Add("Кабинет", typeof(String));
            DTBL_MAIN3.Columns.Add("Адрес", typeof(String));
            DTBL_MAIN3.Columns.Add("Почта", typeof(String));

            if (e.CommandName.ToString() == "ColumnClick")
            {
                GridViewRow selectedRow = GridResults_Users.SelectedRow;
                int selectedRowIndex = Convert.ToInt32(e.CommandArgument.ToString());
                int selectedColumnIndex = Convert.ToInt32(Request.Form["__EVENTARGUMENT"].ToString()); 
                // GridResults.Rows[selectedRowIndex].Cells[selectedColumnIndex].Attributes["style"] += "background-color:Red;";

                
                if ((selectedColumnIndex == 6) )
                {
                    if (sCompany.IndexOf("nbsp") == -1)
                    {
                        DTBL_MAIN3.Clear();
                        sDepartment = "*";
                        sCompany = GridResults_Users.Rows[selectedRowIndex].Cells[selectedColumnIndex].Text;
                        sCompany = sCompany.Replace("&quot;", "\"");
                        if (sCompany.IndexOf("Отдел") >= 0)
                        {
                            sDepartment = GridResults_Users.Rows[selectedRowIndex].Cells[selectedColumnIndex].Text;
                            sCompany = "Администрация городского округа город Воронеж";
                        }

                        DTBL_MAIN3 = findLDAPuser(sCompany, sDepartment);
                        WriteLog("Организация: " + sCompany);
                    }

                }

                
                if ((selectedColumnIndex == 5))
                {

                    if (sDepartment.IndexOf("nbsp") == -1)
                    {
                        sDepartment = GridResults_Users.Rows[selectedRowIndex].Cells[selectedColumnIndex].Text;
                        DTBL_MAIN3.Clear();

                        sCompany = GridResults_Users.Rows[selectedRowIndex].Cells[selectedColumnIndex + 1].Text;

                        DTBL_MAIN3 = findLDAPuser(sCompany, sDepartment);
                        WriteLog("Организация: " + sCompany + ", отдел: " + sDepartment);
                    }
                }

                if (selectedColumnIndex == 7)
                {

                    DTBL_MAIN3.Clear();

                    DTBL_MAIN3 = findLDAPAddress(GridResults_Users.Rows[selectedRowIndex].Cells[selectedColumnIndex].Text, GridResults_Users.Rows[selectedRowIndex].Cells[selectedColumnIndex + 1].Text);
                    WriteLog("Кабинет: " + GridResults_Users.Rows[selectedRowIndex].Cells[selectedColumnIndex].Text + ", Адрес: " + GridResults_Users.Rows[selectedRowIndex].Cells[selectedColumnIndex + 1].Text);
                }

                if (selectedColumnIndex == 8)
                {

                    DTBL_MAIN3.Clear();

                    DTBL_MAIN3 = findLDAPAddress("*", GridResults_Users.Rows[selectedRowIndex].Cells[selectedColumnIndex].Text);
                    WriteLog("Адрес: " + GridResults_Users.Rows[selectedRowIndex].Cells[selectedColumnIndex + 1].Text);

                }

                if (selectedColumnIndex == 9)
                {

                    findLDAPuserMail(GridResults_Users.Rows[selectedRowIndex].Cells[2].Text);

                    ClientScript.RegisterStartupScript(this.GetType(), "mailto", "parent.location='mailto:" + smail + "'", true);
                    WriteLog("Отправить письмо: " + smail);

                }

                //   Response.Redirect("~/Users.aspx");
                if (DTBL_MAIN3.Rows.Count > 0)
                {
                    DTBL_MAIN3.Columns.Remove("№");
                    Session["MyResults"] = DTBL_MAIN3;
                    Response.Redirect("Users.aspx");
                    //Response.Write("<script>");
                    //Response.Write("window.open('Users.aspx','_blank')");
                    //Response.Write("</script>");
                }
                // Label2.Text = GridResults.Rows[selectedRowIndex].Cells[selectedColumnIndex].Text;
            }
        }

        protected void GridResults_Users_SelectedIndexChanged(object sender, EventArgs e)
        {
            GridViewRow selectedRow = GridResults_Users.SelectedRow;
        }


        public string sParseHTMLTAG(string s)
        {
            string sNew = String.Empty;
            sNew = System.Text.RegularExpressions.Regex.Replace(s, @"&nbsp;", "");
            sNew = System.Text.RegularExpressions.Regex.Replace(sNew, @"&quot;", "\"");
            sNew = System.Text.RegularExpressions.Regex.Replace(sNew, @"&#39;", "\"");
            return sNew;
        }

        public DataTable gridviewToDataTable(GridView gv)
        {

            // Create a new DataTable.
            DTBL_MAIN = new DataTable("TableTitleSort");
            // Create new DataColumn, set DataType, 
            // ColumnName and add to DataTable.  

            //DataColumn workCol = DTBL_MAIN.Columns.Add("№", typeof(String));
            //workCol.AllowDBNull = true;
            //workCol.Unique = false;
            DTBL_MAIN.Columns.Add("ФИО", typeof(String));
            DTBL_MAIN.Columns.Add("Телефон", typeof(String));
            DTBL_MAIN.Columns.Add("IP телефон", typeof(String));
            DTBL_MAIN.Columns.Add("Должность", typeof(String));
            DTBL_MAIN.Columns.Add("Отдел", typeof(String));
            DTBL_MAIN.Columns.Add("Организация", typeof(String));
            DTBL_MAIN.Columns.Add("Кабинет", typeof(String));
            DTBL_MAIN.Columns.Add("Адрес", typeof(String));
            DTBL_MAIN.Columns.Add("Почта", typeof(String));

            foreach (GridViewRow row in gv.Rows)
            {
                DataRow dr;
                dr = DTBL_MAIN.NewRow();


               // dr["№"] = row.Cells[1].Text;
                dr["ФИО"] = sParseHTMLTAG(row.Cells[1].Text);
                dr["Телефон"] = sParseHTMLTAG(row.Cells[2].Text);
                dr["IP телефон"] = sParseHTMLTAG(row.Cells[3].Text);
                dr["Должность"] = sParseHTMLTAG(row.Cells[4].Text);
                dr["Отдел"] = sParseHTMLTAG(row.Cells[5].Text);
                dr["Организация"] = sParseHTMLTAG(row.Cells[6].Text);
                dr["Кабинет"] = sParseHTMLTAG(row.Cells[7].Text);
                dr["Адрес"] = sParseHTMLTAG(row.Cells[8].Text);
                dr["Почта"] = sParseHTMLTAG(row.Cells[9].Text);


                DTBL_MAIN.Rows.Add(dr);
            }



            return DTBL_MAIN;
        }

        protected void GridResults_Users_Sorting(object sender, GridViewSortEventArgs e)
        {
     
            DataTable dataTable = new DataTable();
           
                dataTable = gridviewToDataTable(GridResults_Users);


            
            if (dataTable != null)
            {

                            if (Session["UserSortfield"].ToString() == e.SortExpression)
                            {
                                if (Session["UserSortnapravl"].ToString() == "DESC")
                                {
                    
                                    Session["UserSortnapravl"] = "ASC";
                                }
                                else
                                {
                   
                                    Session["UserSortnapravl"] = "DESC";
                                }

                            }
                            else
                            {
                                Session["UserSortnapravl"] = "DESC";
                                Session["UserSsortfield"] = e.SortExpression;
                            }

                            if (e.SortExpression == "Должность")
                            {

                                Session["UserSortfield"] = "title";


                                   dataTable.Columns.Add("TitleNom", typeof(Int32));

                                    string LookTitle = "";
                                    foreach (DataRow dr in dataTable.Rows)
                                    {
                                        dr["TitleNom"] = 10000;
                                        LookTitle = dr[3].ToString();
                                       
                                        
                                        if ((LookTitle.IndexOf("Методист") >= 0))
                                        {
                                            dr["TitleNom"] = 1300;
                                        }

                                        if ((LookTitle.IndexOf("Специалист второй категории") >= 0))
                                        {
                                            dr["TitleNom"] = 1210;
                                        }

                                        if ((LookTitle.IndexOf("Специалист первой категории") >= 0))
                                        {
                                            dr["TitleNom"] = 1200;
                                        }

                                        if ((LookTitle.IndexOf("Ведущий инженер") >= 0))
                                        {
                                            dr["TitleNom"] = 1100;
                                        }

                                        if ((LookTitle.IndexOf("Ведущий эксперт") >= 0))
                                        {
                                            dr["TitleNom"] = 1000;
                                        }

                                        if ((LookTitle.IndexOf("Ведущий специалист") >= 0))
                                        {
                                            dr["TitleNom"] = 900;
                                        }


                                        if ((LookTitle.IndexOf("Главный специалист") >= 0))
                                        {
                                            dr["TitleNom"] = 800;
                                        }

                                        if ((LookTitle.IndexOf("Главный специалист по") >= 0))
                                        {
                                            dr["TitleNom"] = 810;
                                        }

                                        if ((LookTitle.IndexOf("Консультант") >= 0))
                                        {
                                            dr["TitleNom"] = 700;
                                        }

                                        if ((LookTitle.IndexOf("Бухгалтер второй категории") >= 0))
                                        {
                                            dr["TitleNom"] = 631;
                                        }

                                        if ((LookTitle.IndexOf("Бухгалтер") >= 0))
                                        {
                                            dr["TitleNom"] = 630;
                                        }

                                        if ((LookTitle.IndexOf("Ведущий бухгалтер") >= 0))
                                        {
                                            dr["TitleNom"] = 620;
                                        }

                                        if ((LookTitle.IndexOf("Заместитель главного бухгалтера") >= 0))
                                        {
                                            dr["TitleNom"] = 610;
                                        }

                                        if ((LookTitle.IndexOf("Главный бухгалтер") >= 0))
                                        {
                                            dr["TitleNom"] = 600;
                                        }

                                        if ((LookTitle.IndexOf("Заместитель начальника отдела") >= 0))
                                        {
                                            dr["TitleNom"] = 500;
                                        }

                                        if (LookTitle.IndexOf("Заместитель директора") >= 0)
                                        {
                                            dr["TitleNom"] = 490;
                                        }

                                        if (LookTitle.IndexOf("И.о. начальник отдела") >= 0)
                                        {
                                            dr["TitleNom"] = 400;
                                        }

                                        if (LookTitle.IndexOf("Начальник отдела") >= 0)
                                        {
                                            dr["TitleNom"] = 400;
                                        }

                                        if (LookTitle.IndexOf(" заместитель руководителя") >= 0)
                                        {
                                            dr["TitleNom"] = 302;
                                        }

                                        if (LookTitle.IndexOf("И.о. заместителя руководителя") >= 0)
                                        {
                                            dr["TitleNom"] = 301;
                                        }

                                        if (LookTitle.IndexOf("Заместитель руководителя") >= 0)
                                        {
                                            dr["TitleNom"] = 300;
                                        }

                                        if (LookTitle.IndexOf("И.о. директора") >= 0)
                                        {
                                            dr["TitleNom"] = 391;
                                        }

                                        if (LookTitle.IndexOf("Директор") >= 0)
                                        {
                                            dr["TitleNom"] = 390;
                                        }

                                        if ((LookTitle.IndexOf("И.о. руководителя управы") >= 0))
                                        {
                                            dr["TitleNom"] = 211;
                                        }

                                        if (LookTitle.IndexOf("Руководитель управы") >= 0)
                                        {
                                            dr["TitleNom"] = 210;
                                        }

                                        if (LookTitle.IndexOf("И.о. руководителя управления") >= 0)
                                        {
                                            dr["TitleNom"] = 201;
                                        }

                                        if (LookTitle.IndexOf("Руководитель управления") >= 0)
                                        {
                                            dr["TitleNom"] = 200;
                                        }

                                        if (LookTitle.IndexOf("И.о. заместителя главы администрации") >= 0)
                                        {
                                            dr["TitleNom"] = 101;
                                        }

                                        if (LookTitle.IndexOf("Заместитель главы администрации") >= 0)
                                        {
                                            dr["TitleNom"] = 100;
                                        }

                                        if (LookTitle.IndexOf("Глава городского округа") >= 0)
                                        {
                                            dr["TitleNom"] = 10;
                                        }

                                    } // END foreach (DataRow dr in DTBL_MAIN.Rows)

                                    if (Session["TiTleNapr"].ToString() == "ASC")
                                    {
                                        Session["TiTleNapr"] = "DESC";
                                    }
                                    else
                                    {
                                        Session["TiTleNapr"] = "ASC";
                                    }

                               // }


                                    dataTable.DefaultView.Sort = "TitleNom " + Session["TiTleNapr"].ToString()+", " + "Отдел DESC";

                                dataTable.Columns.Remove("TitleNom");



                            } // end if (e.SortExpression == "Должность")
                            else
                            {
                                Session["UserSortfield"] = e.SortExpression;
                                dataTable.DefaultView.Sort = e.SortExpression + " " + Session["UserSortnapravl"].ToString();

                            }// end else if (e.SortExpression == "Должность")


                                GridResults_Users.DataSource = dataTable;

                                //for (int i = 1; i < GridResults_Users.Rows.Count; i++)
                                //{
                                //    GridResults_Users.Rows[i].Cells[0].Text = "1";//i.ToString();
                                //}

                                    GridResults_Users.DataBind();

               

           }


        }








         
       
    }
}