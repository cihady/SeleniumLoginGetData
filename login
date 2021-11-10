namespace SeleniumLoginGetData
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        ChromeDriver driver;

        
        public void Login()
        {
            driver = new ChromeDriver();
            driver.Navigate().GoToUrl("URL");
            driver.FindElement(By.Id("txtUserName")).SendKeys(""); // UserName
            driver.FindElement(By.Id("txtPassword")).SendKeys(""); // Password
            driver.FindElement(By.ClassName("submitButton")).Click();

        }
        
        public void SendData()
        {
            string gonderilecek;
            driver.FindElement(By.Id("TypeTextID")).SendKeys(sendingText.Text);

            /*driver.SwitchTo().Frame(0);
            driver.FindElement(By.Id("recaptcha-anchor")).Click();
            driver.SwitchTo().DefaultContent();

            driver.SwitchTo().Frame(0);*/ // I am not robot
            //  MessageBox.Show("devam ?"); If captcha is text, manual
        }
        
        public void GetData()
        {

            string gData = driver.FindElement(By.XPath("/html/body/div........")).GetAttribute("value");
             //Full xpath
              string[] Dline = new string[] { gData };
              dataGridView1.Rows.Add(Dline); // Get data to datagridView
        }


        void excel_export(DataGridView dg) // Export to excel
        {
            dg.AllowUserToAddRows = false;
            System.Globalization.CultureInfo lang = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-us");
            Microsoft.Office.Interop.Excel.Application Table = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook book = Table.Workbooks.Add(true);
            Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)Table.ActiveSheet;
            System.Threading.Thread.CurrentThread.CurrentCulture = lang;
            Table.Visible = true;
            sheet = book.ActiveSheet;
            for (int i = 0; i < dg.Rows.Count; i++)
            {
                for (int j = 0; j < dg.ColumnCount; j++)
                {
                    if (i == 0)
                    {
                        Table.Cells[1, j + 1] = dg.Columns[j].HeaderText;
                    }
                    Table.Cells[i + 2, j + 1] = dg.Rows[i].Cells[j].Value.ToString();
                }
            }
            Table.Visible = true;
            Table.UserControl = true;
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            SendData();
            GetData();

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Login();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            excel_export(dataGridView1);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SendData();
            GetData();
        }
    }
}
