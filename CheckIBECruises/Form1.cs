using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TuicContentLoader;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;

namespace CheckIBECruises
{

    public partial class Form1 : Form
    {
        IWebDriver driver;
        String excelDatei;
        public Form1()
        {
            InitializeComponent();

        }

        public void setGrid(int nr, String URL, string status)
        {
            dataGridView1.BeginInvoke(new MethodInvoker(() => dataGridView1.Rows.Add(nr, URL, status)));

            switch (status)
            {
                case "OK":
                    dataGridView1.BeginInvoke(new MethodInvoker(() => dataGridView1.Rows[dataGridView1.Rows.Count - 2].Cells[2].Style.BackColor = Color.LightGreen));
                    break;
                case "ERROR":
                    dataGridView1.BeginInvoke(new MethodInvoker(() => dataGridView1.Rows[dataGridView1.Rows.Count - 2].Cells[2].Style.BackColor = Color.OrangeRed));
                    break;
            }
            // dataGridView1.Rows.Add(nr, URL, status);
        }
        private void button1_Click(object sender, EventArgs e)
        {

            new Thread(() =>
            {
                ExcelLoader loader = new ExcelLoader();
                Excel.Workbook wb = loader.getWorkbook(excelDatei);

                Excel.Worksheet sheet = (Excel.Worksheet)wb.Worksheets.get_Item(1);
                Excel.Range range = sheet.UsedRange;

                Console.WriteLine(range.Rows.Count);
                for (int i = 2; i <= range.Rows.Count; i++)
                //for (int i = 2; i < 10; i++)
                {
                    driver = new ChromeDriver(Directory.GetCurrentDirectory());
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));

                    //Console.WriteLine( driver.Manage().Logs.GetLog("browser") );
                    Console.WriteLine(i + ": " + (range.Cells[i, 5] as Excel.Range).Value2);
                    driver.Url = (range.Cells[i, 5] as Excel.Range).Value2;
                    try
                    {
                        wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("booking-page-headline")));
                        setGrid(i, (range.Cells[i, 5] as Excel.Range).Value2, "OK");
                        //dataGridView1.Rows.Add(i, (range.Cells[i, 5] as Excel.Range).Value2, "OK");
                    }
                    catch (Exception ee)
                    {
                        Console.WriteLine("*** Row " + i + " has error");
                        setGrid(i, (range.Cells[i, 5] as Excel.Range).Value2, "ERROR");
                        //dataGridView1.Rows.Add(i, (range.Cells[i, 5] as Excel.Range).Value2, "ERROR");
                    }
                    //dataGridView1.FirstDisplayedScrollingRowIndex = i - 1;
                    dataGridView1.BeginInvoke(new MethodInvoker(() => dataGridView1.FirstDisplayedScrollingRowIndex = i - 1));

                    driver.Quit();

                }
            }).Start();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                excelDatei = openFileDialog1.FileName;
            }
        }
    }
}
