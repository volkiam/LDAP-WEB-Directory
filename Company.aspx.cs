using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.DirectoryServices;
using System.Configuration;
using System.Data;
using System.Collections;
using System.IO;
using System.Web.Security;
using System.Security.Principal;

namespace directory
{
    public partial class WebForm2 : System.Web.UI.Page
    {
        public const string LDAPuser = "user@domain";
        public const string LDAPpass = "12345678";

        System.Data.DataTable DTBL_MAIN = null;

        System.Data.DataTable DTBL_MAIN2 = null;

        
        TreeNode sNode;


        private DataTable findLDAPuser(string lCompany, string lDepartment)
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

                DirectoryEntry root_Company = new DirectoryEntry("cityhall.voronezh-city.ru");
                root_Company.Path = "LDAP://OU=Employees,DC=cityhall,DC=voronezh-city,DC=ru";
                DirectorySearcher search = new DirectorySearcher(root_Company);
                root_Company.Username = LDAPuser;
                root_Company.Password = LDAPpass;

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

                        if (str != "") DTBL_MAIN.Rows.Add(row);

                    }
                }

            }
            catch (Exception ex)
            {
                string exp = ex.Message;
            }

            return DTBL_MAIN;
        }

        

        protected void Page_Load(object sender, EventArgs e)
        {
            
           
            try
            {
                if (!IsPostBack)
                {
                    TreeView1.Nodes.Clear();
                    findLDAPuser("*","*");
                    PopulateTreeView(); 
                }
            }
            catch
            {
            }
        }


        private void PopulateTreeView()
        {
            TreeNode rootNode;
            
     
 

            rootNode = new TreeNode("Name of the organization");

            List<string> spCompany = new List<string>();
            List<string> spDepartment = new List<string>();


            StreamReader sr = new StreamReader(Server.MapPath("~/Download/") + "DataCompany.dat");
            
            String readstr = "";
                    while (sr.Peek() >= 0)
                    {
                        readstr =sr.ReadLine(); 
                        spCompany.Add(readstr);
                        spDepartment.Add("");
                        rootNode.ChildNodes.Add(new TreeNode(readstr));
                    }
                    sr.Close();

            TreeView1.Nodes.Add(rootNode);
               
            Int32 j = 0;
            string sp = "";
            string sdepartment = "";
            string sname = "";
         
            int p;

            while ((j < DTBL_MAIN.Rows.Count))
            {

                int k = 0;
                for (k = 0; k < spCompany.Count; k++)
                {
                    if ((DTBL_MAIN.Rows[j][6]).ToString() == spCompany[k])
                    {
                        sp = spDepartment[k];
                        sname = spCompany[k];
                        sdepartment = (DTBL_MAIN.Rows[j][5]).ToString();
                        if (sp.IndexOf(sdepartment) == -1 & (DTBL_MAIN.Rows[j][5]).ToString() != sname)
                        {
                            spDepartment[k] = sp + (DTBL_MAIN.Rows[j][5]).ToString() + ";";
                            sNode = new TreeNode((DTBL_MAIN.Rows[j][5]).ToString());
                            p = 0;
                            while ((p<rootNode.ChildNodes.Count) & (rootNode.ChildNodes[p].Value != sname))
                            {
                                p++;
                            }
                            if (rootNode.ChildNodes[p].Value == sname)
                            {
                                rootNode.ChildNodes[p].ChildNodes.Add(sNode);
                            }

                        }
                    }
                }

                j = j + 1;
            }

            TreeView1.CollapseAll();
            TreeView1.Nodes[0].Expand();
        }

        private void GetUsers()
        {
  // ((TreeView)sender).SelectedNode.

            String sCompany = "";
            String sDepartment = "";
            int k = 0;

          
           
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

            TreeNode node = TreeView1.SelectedNode;
            int lvl = 0;
            while (node != null)
            {
                node = node.Parent;
                lvl++;
            }

            if (lvl == 3)
            {

                DTBL_MAIN2.Clear();
                sDepartment = TreeView1.SelectedNode.Value;
                sCompany = TreeView1.SelectedNode.Parent.Value;

                DTBL_MAIN2 = findLDAPuser(sCompany, sDepartment);

                //for (k = 0; k < DTBL_MAIN.Rows.Count; k++)
                //{
                //    if (((DTBL_MAIN.Rows[k][6]).ToString() == sCompany) & ((DTBL_MAIN.Rows[k][5]).ToString() == sDepartment))
                //    {
                //        DTBL_MAIN2.ImportRow(DTBL_MAIN.Rows[k]);
                //    }
                //}
            }
            if (lvl == 2)
            {  
                ////

                ////
                DTBL_MAIN2.Clear();
                sDepartment = "*";
                sCompany = TreeView1.SelectedNode.Value;
                if (sCompany.IndexOf("Отдел") >= 0)
                {
                    sDepartment =TreeView1.SelectedNode.Value;
                    sCompany =  TreeView1.SelectedNode.Parent.Value;
                }

                DTBL_MAIN2 = findLDAPuser(sCompany, sDepartment);

                //for (k = 0; k < DTBL_MAIN.Rows.Count; k++)
                //{
                //    if (((DTBL_MAIN.Rows[k][6]).ToString() == sCompany))
                //    {
                //        DTBL_MAIN2.ImportRow(DTBL_MAIN.Rows[k]);
                //    }
                //}
              
            }
                Session["MyResults"] = DTBL_MAIN2;
               
             //   Response.Redirect("~/Users.aspx");

                
                Response.Write("<script>");
                Response.Write("window.open('Users.aspx','_blank')");
                Response.Write("</script>");
                
                //GridResults.DataSource = DTBL_MAIN2;
                //GridResults.DataBind();
        }

        protected void TreeView1_SelectedNodeChanged(object sender, EventArgs e)
        {
         
          GetUsers();
        }

      

       
    }
}