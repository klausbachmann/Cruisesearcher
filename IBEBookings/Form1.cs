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
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace IBEBookings
{
    public partial class Form1 : Form
    {
        String excelDatei;
        IWebDriver driver;

        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                excelDatei = openFileDialog1.FileName;
                Console.WriteLine(excelDatei);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            new Thread(() =>
            {
                excelDatei = "C:/VSProjekte/git/IBEBookings/IBE_Buchungen.xlsx";
                ExcelLoader loader = new ExcelLoader();
                Excel.Workbook wb = loader.getWorkbook(excelDatei);

                Excel.Worksheet sheet = (Excel.Worksheet)wb.Worksheets.get_Item(1);
                Excel.Range range = sheet.UsedRange;
                
                Console.WriteLine(range.Rows.Count);
                //for (int i = 2; i <= range.Rows.Count; i++)

                /*
                
                driver = new ChromeDriver(Directory.GetCurrentDirectory());
                driver.Url = "https://tuic-ibe-web.stage.cellular.de/?tripCode=MSD173536SEE&ePackageCode=EPKATMSD17351812";
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));

                // Klick Wohlfühlpreis
                wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("div[class='swiper-slide level-1 is-selected swiper-slide-active']")));
                IWebElement sliderBox = driver.FindElement(By.CssSelector("div[class='swiper-slide level-1 is-selected swiper-slide-active']"));
                wait.Until(ExpectedConditions.ElementToBeClickable(sliderBox.FindElement(By.CssSelector("label[for^='input_cabintype_cabintype-feelgood']"))));
                sliderBox.FindElement(By.CssSelector("label[for^='input_cabintype_cabintype-feelgood']")).Click();

                driver.FindElement(By.CssSelector("[class='js-price-change-button-selection']")).Click();

                // Klick weiter
                wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("a[class^='button button-cta button-next-page']")));
                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.CssSelector("[class'spinner-container is-hidden']")));
                driver.FindElement(By.CssSelector("a[class^='button button-cta button-next-page']")).Click();

                

                driver.Quit();

                */
                for (int i = 2; i < 10; i++)
                {
                    /*driver = new ChromeDriver(Directory.GetCurrentDirectory());
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
                    */
                    (range.Cells[i, 18] as Excel.Range).Value2 = "moin";


                    Console.WriteLine(i + ": " + (range.Cells[i, 8] as Excel.Range).Value2);

                }
                loader.saveWorkbook();
                loader.quit();
                //wb.Save();
            }).Start();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            driver.Quit();
        }
    }
}
