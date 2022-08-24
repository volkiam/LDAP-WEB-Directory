using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.DirectoryServices;
using System.Configuration;
using System.Data;
using System.Collections;
using System.Security.Principal;
using System.Threading;
using System.Threading.Tasks;

namespace directory
{

    public partial class WebForm1 : System.Web.UI.Page
    {

        public const string LDAPuser = "user@domain";
        public const string LDAPpass = "12345678";
        public const string RichUsers = "samaccountname1;samaccountname2;samaccountname3;";
        public const string AdminUsers = "samaccountname1;samaccountname2;samaccountname3;";

        System.Data.DataTable DTBL_MAIN = null;

        System.Data.DataTable DTBL_MAIN2 = null;

        System.Data.DataTable DTBL_MAIN3 = null;

        private string sortnapravl = "DESC";
        private string sortfield = "ФИО";
        private Int32 last_nom = 0;
        private string sortfield_next = "displayname";
        private Int32 last_count = 0;
        private Int32 page_count_item = 20;

        protected override object SaveViewState()
        {  // Change Text Property of Label when this function is invoked. 
            //if (HasControls() && (Page.IsPostBack))
            //{
            //    ((Label)(Controls[0])).Text = "Custom Control Has Saved State";
            //}
            // Save State as a cumulative array of objects. 
            object baseState = base.SaveViewState();
            object[] allStates = new object[4];
            allStates[0] = baseState;
            allStates[1] = last_nom;
            allStates[2] = sortfield_next;
            allStates[3] = page_count_item;
            return allStates;
        }

        protected override void LoadViewState(object savedState)
        {
            
            if (savedState != null)
            {
                // Load State from the array of objects that was saved at ; 
                // SavedViewState. 
                object[] myState = (object[])savedState;
                if (myState[0] != null)
                    base.LoadViewState(myState[0]);
                if (myState[1] != null)
                    last_nom = (Int32)myState[1];
                if (myState[2] != null)
                    sortfield_next = (string)myState[2];
                if (myState[3] != null)
                    page_count_item = (Int32)myState[3];

            }
        }



        private void find(string sname, string SortBy,string  Sort_napravl, Int32 find_count, Int32 find_kol)
        {

            // Create a new DataTable.
            DTBL_MAIN = new DataTable("Results");
            // Create new DataColumn, set DataType, 
            // ColumnName and add to DataTable.  

            DataColumn workCol = DTBL_MAIN.Columns.Add("№", typeof(String)); 
            workCol.AllowDBNull = true;
            workCol.Unique = false;
            DTBL_MAIN.Columns.Add("ФИО", typeof(String));
            DTBL_MAIN.Columns.Add("Телефон", typeof(String));
            DTBL_MAIN.Columns.Add("IP телефон", typeof(String));
            DTBL_MAIN.Columns.Add("Должность", typeof(String));
            DTBL_MAIN.Columns.Add("Отдел", typeof(String));
            DTBL_MAIN.Columns.Add("Организация", typeof(String));
            DTBL_MAIN.Columns.Add("Кабинет", typeof(String));
            DTBL_MAIN.Columns.Add("Адрес", typeof(String));
            DTBL_MAIN.Columns.Add("Почта", typeof(String));
            
            // Create a new DataTable.
            DTBL_MAIN2 = new DataTable("Results2");
            // Create new DataColumn, set DataType, 
            // ColumnName and add to DataTable.  

            DataColumn workCol2 = DTBL_MAIN2.Columns.Add("№", typeof(String));
            workCol2.AllowDBNull = true;
            workCol2.Unique = false;
            DTBL_MAIN2.Columns.Add("ФИО", typeof(String));
            DTBL_MAIN2.Columns.Add("Телефон", typeof(String));
            DTBL_MAIN2.Columns.Add("IP телефон", typeof(String));
            DTBL_MAIN2.Columns.Add("Должность", typeof(String));
            DTBL_MAIN2.Columns.Add("Отдел", typeof(String));
            DTBL_MAIN2.Columns.Add("Организация", typeof(String));
            DTBL_MAIN2.Columns.Add("Кабинет", typeof(String));
            DTBL_MAIN2.Columns.Add("Адрес", typeof(String));
            DTBL_MAIN2.Columns.Add("Почта", typeof(String));

            try
            {            

                DirectoryEntry root = new DirectoryEntry("cityhall.voronezh-city.ru");
                root.Path = "LDAP://OU=Employees,DC=cityhall,DC=voronezh-city,DC=ru";
                DirectorySearcher search = new DirectorySearcher(root);
                root.Username = LDAPuser;
                root.Password = LDAPpass;

                if (sname == "*")
                {
                    search.Filter = "(&(objectClass=user)(objectCategory=person)(!userAccountControl:1.2.840.113556.1.4.803:=2))";
                }
                else
                {

                    string sub_a = "(&(!userAccountControl:1.2.840.113556.1.4.803:=2)(|(sn=*" + sname + "*)(displayname=*" + sname + "*)(givenName=*" + sname + "*)(department=*" + sname + "*)(company=*" + sname + "*)";
                    sub_a = sub_a + "(streetAddress=*" + sname + "*)(Ipphone=*" + sname + "*)(telephoneNumber=*" + sname + "*)(title=*" + sname + "*)))";
                    search.Filter = sub_a; 

                }

                if (Sort_napravl == "ASC")
                {
                    SortOption option = new SortOption(SortBy, System.DirectoryServices.SortDirection.Ascending);
                    search.Sort = option;
                    search.Sort.Direction = System.DirectoryServices.SortDirection.Ascending;
                }
                else
                {
                    SortOption option = new SortOption(SortBy, System.DirectoryServices.SortDirection.Descending);
                    search.Sort = option;
                }
               // SortOption option = new SortOption(SortBy, System.DirectoryServices.SortDirection.Descending);
                if (sname == "*")
                {
                    LabelRes.Visible = false;
                    search.SizeLimit = last_nom + page_count_item;
                }
                else
                {
                    LabelRes.Visible = true;
                    search.SizeLimit = 2000;
                }
                
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
                    LabelRes.Text = "Найдено: " + results.Count.ToString();
                    last_count = results.Count;
                    foreach (SearchResult result in results)
                    {
                        DataRow row = DTBL_MAIN.NewRow();

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


                        row["Почта"] = "mail"; 
                        if (str != "")  DTBL_MAIN.Rows.Add(row);

                    }                
                }

            }
            catch (Exception ex)
            {
                string exp = ex.Message;
            }
            if (find_kol == 0)
            {
                DTBL_MAIN2.Clear();
                Int32 j = 0;
                while ((j < page_count_item) & (last_nom < DTBL_MAIN.Rows.Count))
                {
                    DTBL_MAIN2.ImportRow(DTBL_MAIN.Rows[last_nom]);
                    last_nom = last_nom + 1;
                    j = j + 1;

                }

                GridResults.DataSource = DTBL_MAIN2;
            }
            else
            {
                GridResults.DataSource = DTBL_MAIN;
            }


            GridResults.DataBind();

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

        private static string sDisplayName;
        private static string AuthName;

        //___________________________________________________________________________________________________________________________________________//

        protected void Page_Load(object sender, EventArgs e) //Загрузка страницы
        {
            
            try
            {
                AuthName = Page.User.Identity.Name;
                AuthName = AuthName.Substring(AuthName.IndexOf(@"\") + 1, AuthName.Length - AuthName.IndexOf(@"\") - 1);
                Session["AuthName"] = AuthName;


                DirectoryEntry de = new DirectoryEntry("WinNT://" + HttpContext.Current.User.Identity.Name.Replace("\\", "/"));
                string sTempDisplayName = de.Properties["FullName"].Value.ToString();
                sDisplayName = sTempDisplayName;
                Session["DisplayName"] = sDisplayName;

                if (!IsPostBack)
                {
                    Session["TiTleNapr"] = "DESK";

                    Session["sortnapravl"] = "DESK";
                    Session["sortfield"] = "*";
                    find("*", "displayname", "DESK", last_nom + 20, 0);
                    TextBox1.Text = "";
                    TextBox1.Attributes.Add("onkeydown", "if(event.which || event.keyCode){if ((event.which == 13) || (event.keyCode == 13)) {document.getElementById('ctl00_ContentPlaceHolder1_Button3').click();return false;}} else {return true}; ");

                    
                    //AuthName = Page.User.Identity.Name;
                    //AuthName = AuthName.Substring(AuthName.IndexOf(@"\") + 1, AuthName.Length - AuthName.IndexOf(@"\") - 1);
                    //Session["AuthName"] = AuthName;
                    

                    //DirectoryEntry de = new DirectoryEntry("WinNT://" + HttpContext.Current.User.Identity.Name.Replace("\\", "/"));
                    //string sTempDisplayName = de.Properties["FullName"].Value.ToString();
                    //sDisplayName = sTempDisplayName;
                    //Session["DisplayName"] = sDisplayName;

                    Label2.Text = "Пользователь: " + Session["DisplayName"];
                    Label2.Visible = true;
                    if (RichUsers.IndexOf(AuthName) >= 0)
                    {

                        LinkButton2.Visible = true;
                    }
                    if (AdminUsers.IndexOf(AuthName) >= 0)
                    {

                        LinkButton4.Visible = true;
                    }

                }
            }
            catch
            {
            }
        }


        protected void Button1_Click(object sender, EventArgs e) //Нажатие кнопки "Сброс"
        {
            last_nom = 0;
            TextBox1.Text = "";
            find("*", "displayname", "DESK", 20, 0);
        }

        protected void Button3_Click(object sender, EventArgs e) //Нажатие кнопки "Искать"
        {
            WriteLog(TextBox1.Text);
            last_nom = 0;
            find(TextBox1.Text, "displayname", "DESK", page_count_item, 0);
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

            DataColumn workCol = DTBL_MAIN.Columns.Add("№", typeof(String));
            workCol.AllowDBNull = true;
            workCol.Unique = false;
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
 

                dr["№"] = row.Cells[1].Text;
                dr["ФИО"] = sParseHTMLTAG(row.Cells[2].Text);
                dr["Телефон"] = sParseHTMLTAG(row.Cells[3].Text);
                dr["IP телефон"] = sParseHTMLTAG(row.Cells[4].Text);
                dr["Должность"] =sParseHTMLTAG(row.Cells[5].Text);
                dr["Отдел"] = sParseHTMLTAG(row.Cells[6].Text);
                dr["Организация"] = sParseHTMLTAG(row.Cells[7].Text);
                dr["Кабинет"] = sParseHTMLTAG(row.Cells[8].Text);
                dr["Адрес"] = sParseHTMLTAG(row.Cells[9].Text);
                dr["Почта"] = sParseHTMLTAG(row.Cells[10].Text);

                
                DTBL_MAIN.Rows.Add(dr);
            }



            return DTBL_MAIN;
        }

        protected void Gridresults_Sorting(object sender, GridViewSortEventArgs e)
        {

            string whatfind = "*";
            last_nom = 0;
            if (Session["sortfield"].ToString() != "")
            {
                sortfield = Session["sortfield"].ToString();
            }


            if (TextBox1.Text !="") 
            {
                whatfind = TextBox1.Text;
            }

            if (sortfield == e.SortExpression)
            {
                if (Session["sortnapravl"].ToString() != "")
                {
                    sortnapravl = Session["sortnapravl"].ToString();
                }
                if (sortnapravl == "DESK")
                {
                    sortnapravl = "ASC";
                    Session["sortnapravl"] = "ASC";
                }
                else
                {
                    sortnapravl = "DESK";
                    Session["sortnapravl"] = "DESK";
                }

            }
            else 
            {
                sortfield = e.SortExpression;
                sortnapravl = "DESK";
                Session["sortnapravl"] = "DESK";
                Session["sortfield"] = e.SortExpression;
            }
    
          if (e.SortExpression == "Телефон")
          {
            sortfield_next = "telephoneNumber";
          }
          if (e.SortExpression == "IP телефон")
          {
              sortfield_next = "Ipphone";
          }
          if (e.SortExpression == "Отдел")
          {
              sortfield_next = "department";
          }
          if (e.SortExpression == "Организация")
          {
              sortfield_next = "company";
          }
          if (e.SortExpression == "ФИО")
          {
              sortfield_next = "displayname";
          }
          if (e.SortExpression == "Адрес")
          {
              sortfield_next = "streetAddress";
          }

          if (e.SortExpression == "Кабинет")
          {
              sortfield_next = "physicalDeliveryOfficeName";
          }

          if (e.SortExpression == "Должность")
          {

              sortfield_next = "title";


             
              DataTable dataTable = new DataTable();
              dataTable = gridviewToDataTable(GridResults);
              if (dataTable != null)
              {   
                  
                  dataTable.Columns.Add("TitleNom", typeof(Int32));

                  string LookTitle = "";
                  foreach (DataRow dr in DTBL_MAIN.Rows)
                  {
                      dr["TitleNom"] = 10000;
                      LookTitle = dr[4].ToString();



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


                      if (LookTitle.IndexOf("И.о. директора") >= 0)
                      {
                          dr["TitleNom"] = 391;
                      }

                      if (LookTitle.IndexOf("Директор") >= 0)
                      {
                          dr["TitleNom"] = 390;
                      }

                  }

                  if (Session["TiTleNapr"] == "ASC")
                      {
                          Session["TiTleNapr"] = "DESC";
                      }
                      else
                      {
                          Session["TiTleNapr"] = "ASC";
                      }
                      
                 }

                      dataTable.DefaultView.Sort = "TitleNom " + Session["TiTleNapr"].ToString();

                      //dataTable.Columns.Remove("TitleNom");
                      GridResults.DataSource = dataTable;

                      GridResults.DataBind();
                  
              
          } // end if
          else
          {

              find(whatfind, sortfield_next, sortnapravl, last_nom + 20, 0);
          }// end else
        }

        protected void Button5_Click(object sender, EventArgs e)
        {
            string whatfind = "*";
            if (TextBox1.Text != "")
            {
                whatfind = TextBox1.Text;
            }
            find(whatfind, sortfield_next, "DESK", last_nom + page_count_item, 0);
        }

        protected void Button4_Click(object sender, EventArgs e)
        {
            string whatfind = "*";
            if (TextBox1.Text != "")
            {
                whatfind = TextBox1.Text;
            }
            if (last_nom > 2 * page_count_item)
            {
                Int32 j = 0;
                j = last_nom / page_count_item;
                last_nom = j * page_count_item - page_count_item;
            }
            else
            {
                last_nom=0;
            }
            find(whatfind, sortfield_next, "DESK", last_nom , 0);
      
        }

        protected void Button6_Click(object sender, EventArgs e)
        {
            string whatfind = "*";
            if (TextBox1.Text != "")
            {
                whatfind = TextBox1.Text;
            }
            page_count_item = Int32.Parse(DropDownList1.Text);
            last_nom = 0;
            find(whatfind, sortfield_next, "DESK", page_count_item, 0);
        }



        protected void Button8_Click(object sender, EventArgs e)
        {
            // Create a new DataTable.
            DTBL_MAIN = new DataTable("Results");
            // Create new DataColumn, set DataType, 
            // ColumnName and add to DataTable.  

            DataColumn workCol = DTBL_MAIN.Columns.Add("№", typeof(String));
            workCol.AllowDBNull = true;
            workCol.Unique = false;
            DTBL_MAIN.Columns.Add("ФИО", typeof(String));
            DTBL_MAIN.Columns.Add("Телефон", typeof(String));
            DTBL_MAIN.Columns.Add("IP телефон", typeof(String));
            DTBL_MAIN.Columns.Add("Должность", typeof(String));
            DTBL_MAIN.Columns.Add("Отдел", typeof(String));
            DTBL_MAIN.Columns.Add("Организация", typeof(String));
            DTBL_MAIN.Columns.Add("Кабинет", typeof(String));
            DTBL_MAIN.Columns.Add("Адрес", typeof(String));
            DTBL_MAIN.Columns.Add("Почта", typeof(String));

            try
            {
                
                string sname = "*";
                string SortBy = "company";
                string  Sort_napravl = "DESC";

                DirectoryEntry root = new DirectoryEntry("cityhall.voronezh-city.ru");
                root.Path = "LDAP://OU=Employees,DC=cityhall,DC=voronezh-city,DC=ru";
                DirectorySearcher search = new DirectorySearcher(root);
                root.Username = LDAPuser;
                root.Password = LDAPpass;

                if (sname == "*")
                {
                    search.Filter = "(&(objectClass=user)(objectCategory=person)(!userAccountControl:1.2.840.113556.1.4.803:=2))";
                }
                else
                {

                    string sub_a = "(&(!userAccountControl:1.2.840.113556.1.4.803:=2)(|(sn=*" + sname + "*)(displayname=*" + sname + "*)(givenName=*" + sname + "*)(department=*" + sname + "*)(company=*" + sname + "*)";
                    sub_a = sub_a + "(streetAddress=*" + sname + "*)(Ipphone=*" + sname + "*)(telephoneNumber=*" + sname + "*)(title=*" + sname + "*)))";
                    search.Filter = sub_a;

                }

                if (Sort_napravl == "ASC")
                {
                    SortOption option = new SortOption(SortBy, System.DirectoryServices.SortDirection.Ascending);
                    search.Sort = option;
                    search.Sort.Direction = System.DirectoryServices.SortDirection.Ascending;
                }
                else
                {
                    SortOption option = new SortOption(SortBy, System.DirectoryServices.SortDirection.Descending);
                    search.Sort = option;
                }
                // SortOption option = new SortOption(SortBy, System.DirectoryServices.SortDirection.Descending);

                    search.SizeLimit = 2000;
                    
                    


                // если все равно нужно найти вывести все записи на одну страницу
                //if (find_kol != 0)
                //{
                //    LabelRes.Visible = true;
                //    search.SizeLimit = last_nom + page_count_item;
                //}
                //search.Sort = option;             
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
                    LabelRes.Text = "Найдено: " + results.Count.ToString();
                    last_count = results.Count;
                    foreach (SearchResult result in results)
                    {
                        DataRow row = DTBL_MAIN.NewRow();

                        row["№"] = nom.ToString();
                        nom = nom + 1;


                        if (result.Properties["streetAddress"].Count > 0)
                        {
                            row["Адрес"] = result.Properties["streetAddress"][0].ToString();
                        }

                        if (result.Properties["physicalDeliveryOfficeName"].Count > 0)
                        {
                            row["Кабинет"] = result.Properties["physicalDeliveryOfficeName"][0].ToString();
                        }

                        str = "";
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
                       
                        if (str !="") DTBL_MAIN.Rows.Add(row);
                        //   }
                        // }
                    }
                }

            }
            catch (Exception ex)
            {
                string exp = ex.Message;
            }

            Session["MyDTBL"] = DTBL_MAIN;
            
            Response.ClearContent();
            Response.ClearHeaders();
            Response.AddHeader("Content-Disposition", "attachment; filename=Телефонный_справочник_АГО_город_Воронеж.csv");

            Response.ContentType = "application/excel";
            Response.ContentEncoding = System.Text.Encoding.UTF8;
            Response.BinaryWrite(System.Text.Encoding.UTF8.GetPreamble());
            string tab = "";
            foreach (DataColumn dc in DTBL_MAIN.Columns)
            {
                Response.Write(tab + dc.ColumnName);
                tab = ";";//\t
            }
            Response.Write("\n");
            int i;
            foreach (DataRow dr in DTBL_MAIN.Rows)
            {
                tab = "";
                for (i = 0; i < DTBL_MAIN.Columns.Count; i++)
                {
                    Response.Write(tab + dr[i].ToString());
                    tab = ";";
                }
                Response.Write("\n");
            }
            Response.End();

        }


        protected void LinkButton1_Click(object sender, EventArgs e)
        {

            System.Web.HttpResponse response = System.Web.HttpContext.Current.Response;
            response.ClearContent();

            Response.Redirect("~/Download/Телефонный справочник.docx");
        }

        protected void LinkButton2_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/Company.aspx");
        }


       protected void GridResults_RowDataBound(object sender, GridViewRowEventArgs e)
        {

            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                if (e.Row.RowType == DataControlRowType.DataRow)
                {
                    LinkButton _singleClickButton = (LinkButton)e.Row.Cells[0].Controls[0];
                    string _jsSingle = ClientScript.GetPostBackClientHyperlink(_singleClickButton, "");
                    // Add events to each editable cell
                    for (int columnIndex = 6; columnIndex <= 9; columnIndex++)  // for (int columnIndex = 0; columnIndex < e.Row.Cells.Count; columnIndex++)
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
                    string js2 = _jsSingle2.Insert(_jsSingle2.Length - 2, "10");
                    // Add this javascript to the onclick Attribute of the cell
                    e.Row.Cells[10].Attributes["onclick"] = js2;
                    // Add a cursor style to the cells
                    e.Row.Cells[10].Attributes["style"] += "cursor:pointer;cursor:hand;this.style.textDecoration='underline';";
                    e.Row.Cells[10].ForeColor = System.Drawing.ColorTranslator.FromHtml("#2D1CD0");
                    e.Row.Cells[10].BorderColor = System.Drawing.Color.Black;
                    e.Row.Cells[10].Text = "<img src='Images/mail.png' height='25' width='35' onmouseover=javascript:this.src='images/mail1.png' onmouseout=javascript:this.src='images/mail.png' title='Нажмите для отправки письма пользователю'/>";
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

        protected override void Render(HtmlTextWriter writer)
        {
            foreach (GridViewRow r in GridResults.Rows)
            {
                if (r.RowType == DataControlRowType.DataRow)
                {
                    for (int columnIndex = 6; columnIndex <= 9; columnIndex++)  //
                    {
                        Page.ClientScript.RegisterForEventValidation(r.UniqueID + "$ctl00", columnIndex.ToString());
                    }

                    Page.ClientScript.RegisterForEventValidation(r.UniqueID + "$ctl00", "10");
                }
            }
            base.Render(writer);
        }

        protected void GridResults_SelectedIndexChanged(object sender, EventArgs e)
        {
            GridViewRow selectedRow = GridResults.SelectedRow;
        }

        private DataTable findLDAPuser(string lCompany, string lDepartment)
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
                    int nom = 1;
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

        protected void GridResults_RowCommand(object sender, GridViewCommandEventArgs e)
        {

            String sCompany = "";
            String sDepartment = "";
            

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

            if (e.CommandName.ToString() == "ColumnClick")
            {
                GridViewRow selectedRow = GridResults.SelectedRow;
                int selectedRowIndex = Convert.ToInt32(e.CommandArgument.ToString());
                int selectedColumnIndex = Convert.ToInt32(Request.Form["__EVENTARGUMENT"].ToString());

                sCompany = GridResults.Rows[selectedRowIndex].Cells[selectedColumnIndex].Text;
                sCompany = sCompany.Replace("&quot;","\"");
                if ((selectedColumnIndex == 7) & (sCompany.IndexOf("nbsp") == -1))
                {
                    DTBL_MAIN3.Clear();
                    sDepartment = "*";
                    
                    if (sCompany.IndexOf("Отдел") >= 0)
                    {
                        sDepartment = GridResults.Rows[selectedRowIndex].Cells[selectedColumnIndex].Text;
                        sCompany = "Name of organization";
                    }

                    DTBL_MAIN3 = findLDAPuser(sCompany, sDepartment);
                    WriteLog("Организация: " + sCompany);

                }

                sDepartment = GridResults.Rows[selectedRowIndex].Cells[selectedColumnIndex].Text;
                if ((selectedColumnIndex == 6) & (sDepartment.IndexOf("nbsp") ==-1))
                {

                    DTBL_MAIN3.Clear();
                    
                    sCompany = GridResults.Rows[selectedRowIndex].Cells[selectedColumnIndex+1].Text;

                    DTBL_MAIN3 = findLDAPuser(sCompany,sDepartment );
                    WriteLog("Организация: " + sCompany + ", отдел: " + sDepartment);

                }

                if (selectedColumnIndex == 8)
                {

                    DTBL_MAIN3.Clear();
                    
                    DTBL_MAIN3 = findLDAPAddress(GridResults.Rows[selectedRowIndex].Cells[selectedColumnIndex].Text, GridResults.Rows[selectedRowIndex].Cells[selectedColumnIndex + 1].Text);
                    WriteLog("Кабинет: " + GridResults.Rows[selectedRowIndex].Cells[selectedColumnIndex].Text + ", Адрес: " + GridResults.Rows[selectedRowIndex].Cells[selectedColumnIndex + 1].Text);
                }

                if (selectedColumnIndex == 9)
                {

                    DTBL_MAIN3.Clear();

                    DTBL_MAIN3 = findLDAPAddress("*", GridResults.Rows[selectedRowIndex].Cells[selectedColumnIndex].Text);
                    WriteLog("Адрес: " + GridResults.Rows[selectedRowIndex].Cells[selectedColumnIndex + 1].Text);

                }

                if (selectedColumnIndex == 10)
                {

                    findLDAPuserMail(GridResults.Rows[selectedRowIndex].Cells[2].Text);

                    ClientScript.RegisterStartupScript(this.GetType(), "mailto", "parent.location='mailto:" + smail + "'", true);
                    WriteLog("Отправить письмо: " + smail);

                }
               

              
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

        protected void GridResults_RowCreated(object sender, GridViewRowEventArgs e)
        {

          
          
        }

        protected void LinkButton4_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/LogViewer.aspx");
        }

        protected void GridResults_PageIndexChanging(object sender, System.Web.UI.WebControls.GridViewPageEventArgs e)
        {
            GridResults.PageIndex = e.NewPageIndex;
            GridResults.DataBind();
        }



    }
}