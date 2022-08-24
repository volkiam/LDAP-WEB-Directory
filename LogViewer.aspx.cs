using System;
using System.IO;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

namespace directory
{
    public partial class LogViewer : System.Web.UI.Page
    {

        public const string AdminUsers = "samaccountname1;samaccountname2;samaccountname3"; //хахардкожено, так было надо
        public static string sFileName;
        public static string sColumnSort;
        public static string sNapravlSort;

        System.Data.DataTable DTBL_MAIN = null;
        System.Data.DataTable DTBL_Log = null;

       
        private void GetFilesLog()
        {
            string sDir = Server.MapPath("/Logs");
            string[] dirs = Directory.GetFiles(sDir, "*.txt");
            DropDownList1.Items.Clear();
            foreach (string dir in dirs)
            {
                DropDownList1.Items.Add(dir.Substring(dir.IndexOf(@"Logs\") + 5));
                
            }
            if (sFileName != "")
            {
                ListItem Item = DropDownList1.Items.FindByText(sFileName.Substring(sFileName.IndexOf(@"Logs\") + 5));
                int nom = DropDownList1.Items.IndexOf(Item);
                DropDownList1.SelectedIndex = nom;

            }
        }


        private void GetLogs(string sFileName)
    {
                    // Create a new DataTable.
             DTBL_MAIN = new DataTable("UserResults");
            // Create new DataColumn, set DataType, 
            // ColumnName and add to DataTable.  

            DataColumn workCol = DTBL_MAIN.Columns.Add("№", typeof(String)); 
            workCol.AllowDBNull = true;
            workCol.Unique = false;
            DTBL_MAIN.Columns.Add("Дата", typeof(String));
            DTBL_MAIN.Columns.Add("Пользователь", typeof(String));
            DTBL_MAIN.Columns.Add("Логин", typeof(String));
            DTBL_MAIN.Columns.Add("Запрос", typeof(String));

            DateTime dt = DateTime.Now;
            try
            {   // Open the text file using a stream reader.
                int counter = 0;
                string line;
                

                // Read the file and display it line by line.  
                System.IO.StreamReader file = new System.IO.StreamReader(sFileName);
                while ((line = file.ReadLine()) != null)
                {
                    string[] tokens = line.Split('|');
                    DataRow row = DTBL_MAIN.NewRow();
                    row["№"] = (counter+1).ToString();
                    row["Дата"] = tokens[0]; 
                    row["Пользователь"] = tokens[1].Substring(tokens[1].IndexOf(':') + 1);
                    row["Логин"] = tokens[2].Substring(tokens[2].IndexOf(':') + 1); ;
                    row["Запрос"] = tokens[3].Substring(tokens[3].IndexOf(':') + 1); ;
                    DTBL_MAIN.Rows.Add(row);
                    counter++;
                }

                file.Close();

                if( (sNapravlSort != null) & (sColumnSort != null))
                {
                    if (sColumnSort == "Пользователь")
                    {
                        if (sNapravlSort == "ASC")
                        {
                            DTBL_MAIN.DefaultView.Sort = "Пользователь ASC";
                        }
                        else 
                        {
                            DTBL_MAIN.DefaultView.Sort = "Пользователь DESC";
                        }
                    }

                    if (sColumnSort == "Запрос")
                    {
                        if (sNapravlSort == "ASC")
                        {
                            DTBL_MAIN.DefaultView.Sort = "Запрос ASC";
                        }
                        else
                        {
                            DTBL_MAIN.DefaultView.Sort = "Запрос DESC";
                        }
                    }

                    if (sColumnSort == "Логин")
                    {
                        if (sNapravlSort == "ASC")
                        {
                            DTBL_MAIN.DefaultView.Sort = "Логин ASC";
                        }
                        else
                        {
                            DTBL_MAIN.DefaultView.Sort = "Логин DESC";
                        }
                    }

                    if (sColumnSort == "Дата")
                    {
                        if (sNapravlSort == "ASC")
                        {
                            DTBL_MAIN.DefaultView.Sort = "Дата ASC";
                        }
                        else
                        {
                            DTBL_MAIN.DefaultView.Sort = "Дата DESC";
                        }
                    }
                }

                GridView1.DataSource = DTBL_MAIN;
                Session["LoadingLogs"] = DTBL_MAIN;
                DTBL_Log = DTBL_MAIN;
                
                GridView1.DataBind();
                GridView1.SetPageIndex(GridView1.PageCount);
               
                
            }
            catch (Exception e)
            {
                Console.WriteLine("The file could not be read:");
            }

    }
       
        public static string AuthName;
//____________________________________________________________________________________________________//
        protected void Page_Load(object sender, EventArgs e)
        {
            try //main body
            {
                if (!IsPostBack)
                {
                    Session["LogSortnapravl"] = "ASC";
                    Session["LogSortfield"] = "Дата";
                }

                if (Session["LogSortnapravl"] != null)
                {
                    sNapravlSort = Session["LogSortnapravl"].ToString();
                }
                if (Session["LogSortfield"] != null)
                {
                    sColumnSort = Session["LogSortfield"].ToString();
                }                    
                    
                    try //AuthName
                    {
                        AuthName = Page.User.Identity.Name;
                        AuthName = AuthName.Substring(AuthName.IndexOf(@"\") + 1, AuthName.Length - AuthName.IndexOf(@"\") - 1);
                    }
                    catch //AuthName
                    { }

                    try // subbody 1
                    {
                        if (AdminUsers.IndexOf(AuthName) >= 0)
                        {
                            DateTime dt = DateTime.Now;

                            if (sFileName == null)
                            {
                                sFileName = Server.MapPath(@"Logs\log_" + dt.ToString("ddMMyyyy") + ".txt");
                                GetFilesLog();
                            }
                            else
                            {
                                if (DropDownList1.SelectedIndex == -1)
                                {
                                    GetFilesLog();
                                    sFileName = Server.MapPath(@"Logs\log_" + dt.ToString("ddMMyyyy") + ".txt");
                                }
                                else
                                {
                                    sFileName = Server.MapPath(@"Logs\" + DropDownList1.Items[DropDownList1.SelectedIndex]);
                                }
                            }

                            GetLogs(sFileName);

                        }
                    }
                    catch // subbody 1
                    { }
                
            }
            catch //main body
                { }

        }

        protected void BindData()
        {

            GridView1.DataBind();
            
        }  

        protected void GridView1_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            GridView1.PageIndex = e.NewPageIndex;
            BindData(); 
        }

        protected void DropDownList1_SelectedIndexChanged(object sender, EventArgs e)
        {
           

        }

        protected void GridView1_Sorting(object sender, GridViewSortEventArgs e)
        {
            if (sColumnSort == e.SortExpression)
            {
                if (Session["LogSortnapravl"].ToString() != "")
                {
                    sNapravlSort = Session["LogSortnapravl"].ToString();
                }
                if (sNapravlSort == "DESK")
                {
                    sNapravlSort = "ASC";
                    Session["LogSortnapravl"] = "ASC";
                }
                else
                {
                    sNapravlSort = "DESK";
                    Session["LogSortnapravl"] = "DESK";
                }

            }
            else
            {
                
                sNapravlSort = "DESK";
                Session["LogSortnapravl"] = "DESK";
                Session["LogSortfield"] = e.SortExpression;
            }

            sColumnSort = e.SortExpression;
           
        }// end of GridView1_Sorting

        protected void Button1_Click(object sender, EventArgs e)
        {

              sFileName = Server.MapPath(@"Logs\" + DropDownList1.Items[DropDownList1.SelectedIndex]);

              GetLogs(sFileName);    
        }

        public string sParseHTMLTAG(string s)
        {
            string sNew = String.Empty;
            sNew = System.Text.RegularExpressions.Regex.Replace(s, @"&nbsp;", "");
            sNew = System.Text.RegularExpressions.Regex.Replace(sNew, @"&quot;", "\"");
            sNew = System.Text.RegularExpressions.Regex.Replace(sNew, @"&#39;", "\"");
            return sNew;
        }

       
        protected void Button2_Click(object sender, EventArgs e)
        {
          
            DataTable dt = DTBL_Log; // gridviewToDataTable(GridView1);
            DataView dataView = dt.DefaultView;
            dataView.RowFilter = "Пользователь LIKE '%" + TextBox1.Text + "%'";

            GridView1.DataSource = dataView;
            GridView1.DataBind();
        } 
    }
}