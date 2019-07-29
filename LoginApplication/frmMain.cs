using System;
using System.Data;
using System.Windows.Forms;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;


namespace LoginApplication
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
           
        }

        public void inputExcelFile()
        {

            bool isNumeric0 = textBox1.Text.All(char.IsDigit);
            if (isNumeric0 == false)
            {
                string message = "Please Insert Numbers";
                string title = "Information";
                MessageBox.Show(message, title);
                return;
            }
            bool isNumeric1 = textBox2.Text.All(char.IsDigit);
            if (isNumeric1 == false)
            {
                string message = "Please Insert Numbers";
                string title = "Information";
                MessageBox.Show(message, title);
                return;
            }
            bool isNumeric2 = textBox3.Text.All(char.IsDigit);
            if (isNumeric2 == false)
            {
                string message = "Please Insert Numbers";
                string title = "Information";
                MessageBox.Show(message, title);
                return;
            }
            bool isNumeric3 = textBox4.Text.All(char.IsDigit);
            if (isNumeric3 == false)
            {
                string message = "Please Insert Numbers";
                string title = "Information";
                MessageBox.Show(message, title);
                return;
            }
            bool isNumeric4 = textBox5.Text.All(char.IsDigit);
            if (isNumeric4 == false)
            {
                string message = "Please Insert Numbers";
                string title = "Information";
                MessageBox.Show(message, title);
                return;
            }
            List<Product> productItemIn = new List<Product>();
            productItemIn = frmLogin.GetProducts();

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Object pwd = productItemIn[0].pwd;
            Object MissingValue = System.Reflection.Missing.Value;
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(Application.StartupPath + "\\Logic.xlsx", MissingValue, MissingValue, MissingValue, pwd);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            try
            {
                xlWorksheet.Cells[3, 7] = long.Parse(textBox1.Text);
                xlWorksheet.Cells[4, 7] = long.Parse(textBox2.Text);
                xlWorksheet.Cells[6, 7] = long.Parse(textBox3.Text);
                xlWorksheet.Cells[10, 7] = long.Parse(textBox4.Text);
                xlWorksheet.Cells[11, 7] = long.Parse(textBox5.Text);
                xlWorkbook.Save();

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [something].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);

                string message = "Data Inputed!";
                string title = "Information";
                MessageBox.Show(message, title);
                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";
                textBox4.Text = "";
                textBox5.Text = "";
                textBox6.Text = "";
                textBox7.Text = "";
                textBox8.Text = "";
            }
            catch (Exception ex)
            {
                
                string message = ex.Message; 
                string title = "Information";
                MessageBox.Show(message, title);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                return;
            }
        }


        public void outputExcelFile()
        {

                List<Product> productItemOut = new List<Product>();
                productItemOut = frmLogin.GetProducts();
                
            //Create COM Objects. Create a COM object for everything that is referenced
                Excel.Application xlApp = new Excel.Application();
                Object pwd = productItemOut[0].pwd;
                Object MissingValue = System.Reflection.Missing.Value;
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(Application.StartupPath + "\\Logic.xlsx",MissingValue, MissingValue, MissingValue, pwd);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;
            try
            {
                textBox6.Text = xlRange.Cells[5, 7].Value2.ToString();
                textBox7.Text = xlRange.Cells[18, 7].Value2.ToString();
                textBox8.Text = xlRange.Cells[20, 7].Value2.ToString();

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [something].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }
            catch (Exception ex)
            {
                //Console.Write("No Students exists for the specified name.");
                string message = ex.Message;
                string title = "Information";
                MessageBox.Show(message, title);

                GC.Collect();
                GC.WaitForPendingFinalizers();
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                return;
            }
        }

        //btn_LogOut Click Event
        private void btn_LogOut_Click(object sender, EventArgs e)
        {
            foreach (Process clsProcess in Process.GetProcesses())
                if (clsProcess.ProcessName.Equals("EXCEL"))  //Process Excel?
                    clsProcess.Kill();

            this.Hide();
            //this.Close();
            frmLogin fl = new frmLogin();
            fl.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == "" || textBox4.Text == "" || textBox5.Text == "")
            {
                string message = "Input Data exactly!";
                string title = "Information";
                MessageBox.Show(message, title);
            }
            else
            {
                inputExcelFile();
            }
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            outputExcelFile();
        }

        private void frmMain_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                foreach (Process clsProcess in Process.GetProcesses())
                    if (clsProcess.ProcessName.Equals("EXCEL"))  //Process Excel?
                        clsProcess.Kill();
            }
            catch
            {
                Application.Exit();
                return;
            }
            Application.Exit();
        }
    }
}
