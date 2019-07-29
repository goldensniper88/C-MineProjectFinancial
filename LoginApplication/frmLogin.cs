using System;
using System.Data;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Collections.Generic;
using System.Diagnostics;

//using System.Data.SqlClient;

namespace LoginApplication
{
    public partial class frmLogin : Form
    {
        bool flagLogin = false;
        List<string[]> _names = new List<string[]>();
        List<double[]> _pwdArray = new List<double[]>();
        

        public frmLogin()
        {
            InitializeComponent();
        }

        //Create and Return a OleDbConnection obj.
        private static OleDbConnection GetConnection()
        {
            OleDbConnection conn = new OleDbConnection();
            try
            {
                String connectionString = @"Provider=Microsoft.JET.OlEDB.4.0;"
               + @"Data Source=" + Application.StartupPath + "\\mydb.mdb";
        conn = new OleDbConnection(connectionString);
                conn.Open();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return conn;
        }

        private static void CloseConnection(OleDbConnection conn)
        {
            try
            {
                conn.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        //Return List<Product>
        public static List<Product> GetProducts()

        {
            List<Product> productList = new List<Product>();

            DataSet ds = new DataSet();
            OleDbConnection conn = GetConnection();
            OleDbDataAdapter da = new OleDbDataAdapter("Select * from MyTable", conn);
            da.Fill(ds);
            conn.Close();
            DataTable dt = ds.Tables[0];
            foreach (DataRow rows in dt.Rows)
            {
                Product product = new Product();
                product.id = int.Parse(rows["ID"].ToString());
                product.name = rows["Name"].ToString();
                product.pwd = rows["Password"].ToString();
                product.identity = int.Parse(rows["Identity"].ToString());
                productList.Add(product);
            }
            return productList;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            flagLogin = true;
            if (txt_UserName.Text=="" || txt_Password.Text=="")
            {
                string message = "Please provide UserName and Password";
                string title = "Information";
                MessageBox.Show(message, title);
                return; 
            }

            try
            {
                GetConnection();

                List<Product> getDBItems = new List<Product>();
                getDBItems = GetProducts();
                foreach (var item in getDBItems)
                    if (item.name == txt_UserName.Text && item.pwd == txt_Password.Text)
                    {
                        flagLogin = true;
                        break;
                    }
                    else
                    {
                        flagLogin = false;
                    }
                        //Console.WriteLine(item);

                CloseConnection(GetConnection());

                if (flagLogin == true)
                {
                    string message = "Log in Succeed!";
                    string title = "Information";
                    MessageBox.Show(message, title);

                    this.Hide();
                    frmMain fm = new frmMain();
                    fm.Show();
                }
                else
                {
                    string message = "Log in Failed!";
                    string title = "Information";
                    MessageBox.Show(message, title);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void frmLogin_FormClosing(object sender, FormClosingEventArgs e)
        {
            foreach (Process clsProcess in Process.GetProcesses())
                if (clsProcess.ProcessName.Equals("EXCEL"))  //Process Excel?
                    clsProcess.Kill();
            Application.Exit();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            foreach (Process clsProcess in Process.GetProcesses())
                if (clsProcess.ProcessName.Equals("EXCEL"))  //Process Excel?
                    clsProcess.Kill();
            Application.Exit();
        }

        private void frmLogin_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }
    }
}
