namespace SeleniumVeriGonderAl
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        ChromeDriver driver;

        
        public void Giris()
        {
            driver = new ChromeDriver();
            driver.Navigate().GoToUrl("URL");


           // driver.FindElement(By.Id("termsChkbx")).Click();
           // driver.FindElement(By.Id("sub1")).Click(); login kısmına geçmek için tıklanacak yer varsa id sini bul tıkla

            driver.FindElement(By.Id("txtKullanıcı")).SendKeys(""); // Kullanıcı Adı kısmının girileceği text id si
            driver.FindElement(By.Id("txtSifre")).SendKeys(""); // Şifre kısmının girileceği text id si
            driver.FindElement(By.ClassName("submitButton")).Click();

        }
        public void VeriGonder()
        {
            string gonderilecek;
            driver.FindElement(By.Id("gönderilecekveri idsi")).SendKeys(gonderilecek.Text);

            /*driver.SwitchTo().Frame(0);
            driver.FindElement(By.Id("recaptcha-anchor")).Click();
            driver.SwitchTo().DefaultContent();

            driver.SwitchTo().Frame(0);*/ // Robot değilim kısmına tıklayacağımız yer
        }
        public void VeriAl()
        {

            string alinacakveri;
            alinacakveri = driver.FindElement(By.XPath("/html/body/div........")).GetAttribute("value");
             //Çekilecek değerin full xpathini yaz

              string[] satir = new string[] { alinacakveri };
              dataGridView1.Rows.Add(satir); // datagride çekilen veriyi aktar

        }


        void excele_aktar(DataGridView dg) // datagriddeki veriyi excele aktar
        {
            dg.AllowUserToAddRows = false;
            System.Globalization.CultureInfo dil = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-us");
            Microsoft.Office.Interop.Excel.Application Tablo = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook kitap = Tablo.Workbooks.Add(true);
            Microsoft.Office.Interop.Excel.Worksheet sayfa = (Microsoft.Office.Interop.Excel.Worksheet)Tablo.ActiveSheet;
            System.Threading.Thread.CurrentThread.CurrentCulture = dil;
            Tablo.Visible = true;
            sayfa = kitap.ActiveSheet;
            for (int i = 0; i < dg.Rows.Count; i++)
            {
                for (int j = 0; j < dg.ColumnCount; j++)
                {
                    if (i == 0)
                    {
                        Tablo.Cells[1, j + 1] = dg.Columns[j].HeaderText;
                    }
                    Tablo.Cells[i + 2, j + 1] = dg.Rows[i].Cells[j].Value.ToString();
                }
            }
            Tablo.Visible = true;
            Tablo.UserControl = true;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            VeriGonder();
          //  MessageBox.Show("devam ?"); captchayı elle girmemiz gerekirse
            VeriAl();

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Giris();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            excele_aktar(dataGridView1);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            VeriGonder();
            //  MessageBox.Show("devam ?"); captchayı elle girmemiz gerekirse
            VeriAl();
        }
    }
}
