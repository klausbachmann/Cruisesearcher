using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
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
                /*excelDatei = @"C:\Users\M\Documents\Visual Studio 2015\Projects\Cruisesearcher\IBEBookings\IBE_Buchungen.xlsx";
                ExcelLoader loader = new ExcelLoader();
                Excel.Workbook wb = loader.getWorkbook(excelDatei);

                Excel.Worksheet sheet = (Excel.Worksheet)wb.Worksheets.get_Item(1);
                Excel.Range range = sheet.UsedRange;

                Console.WriteLine(range.Rows.Count);
                //for (int i = 2; i <= range.Rows.Count; i++)

                */

                driver = new ChromeDriver(Directory.GetCurrentDirectory());
                driver.Manage().Window.Maximize();
                driver.Url = "https://tuic-ibe-web.stage.cellular.de/?tripCode=MSD173536SEE&ePackageCode=EPKATMSD17351812";
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));

                // Klick Wohlfühlpreis
                wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("div[class='swiper-slide level-1 is-selected swiper-slide-active']")));
                IWebElement sliderBox = driver.FindElement(By.CssSelector("div[class='swiper-slide level-1 is-selected swiper-slide-active']"));
                wait.Until(ExpectedConditions.ElementToBeClickable(sliderBox.FindElement(By.CssSelector("label[for^='input_cabintype_cabintype-feelgood']"))));
                sliderBox.FindElement(By.CssSelector("label[for^='input_cabintype_cabintype-feelgood']")).Click();

                Thread.Sleep(2000);
                if (driver.FindElement(By.CssSelector("[class^='js-price-change-button-selection']")).Displayed)
                {
                    driver.FindElement(By.CssSelector("[class='js-price-change-button-selection']")).Click();
                }

                // Klick weiter
                wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("a[class^='button button-cta button-next-page']")));
                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.CssSelector("[class^='spinner-container']")));
                var footer = driver.FindElement(By.CssSelector("[class='booking-pager bottom']"));
                footer.FindElement(By.CssSelector("a[class^='button button-cta button-next-page']")).Click();

                // Cabin selector
                // Klick weiter
                wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("a[class^='button button-cta button-next-page']")));
                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.CssSelector("[class^='spinner-container']")));

                footer = driver.FindElement(By.CssSelector("[class='booking-pager bottom']"));
                footer.FindElement(By.CssSelector("a[class^='button button-cta button-next-page']")).Click();

                // Flug
                wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("input_shipboundairport_airportcode")));
                new SelectElement(driver.FindElement(By.Id("input_shipboundairport_airportcode"))).SelectByIndex(1);

                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.CssSelector("[class^='spinner-container']")));
                driver.FindElement(By.CssSelector("a[class^='button button-cta button-next-page']")).Click();

                // Paarversicherung
                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.CssSelector("[class^='spinner-container']")));
                var coupleElement = driver.FindElement(By.CssSelector("div[class='number-input-element js-input-element-couple']"));
                coupleElement.FindElement(By.CssSelector("[class='button button-plus']")).Click();

                // Klick auf Versicherungen anzeigen
                driver.FindElement(By.CssSelector("[class='button button-show-insurances js-button-show-insurances '")).Click();

                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.CssSelector("[class^='spinner-container']")));

                // Police klicken
                new SelectElement(driver.FindElement(By.Id("input_policytype_1_policy"))).SelectByIndex(2);

                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.CssSelector("[class^='spinner-container']")));

                // Daten eingeben
                new SelectElement(driver.FindElement(By.Id("input_adult_1_salutationcode"))).SelectByIndex(1);
                driver.FindElement(By.Id("input_adult_1_firstname")).SendKeys("Horst");
                driver.FindElement(By.Id("input_adult_1_lastname")).SendKeys("Ungerer");
                driver.FindElement(By.Id("input_adult_1_dateofbirth")).SendKeys("01.02.1977");

                new SelectElement(driver.FindElement(By.Id("input_adult_2_salutationcode"))).SelectByIndex(1);
                driver.FindElement(By.Id("input_adult_2_firstname")).SendKeys("Horst");
                driver.FindElement(By.Id("input_adult_2_lastname")).SendKeys("Ungerer");
                driver.FindElement(By.Id("input_adult_2_dateofbirth")).SendKeys("24.03.1977");

                driver.FindElement(By.CssSelector("a[class^='button button-cta button-next-page']")).Click();

                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.CssSelector("[class^='spinner-container']")));

                // Invoice
                driver.FindElement(By.Id("input_invoice_streetandhousenumber")).SendKeys("Albert-Schweizer Straße 2");
                driver.FindElement(By.Id("input_invoice_additionaladdress")).SendKeys("oben links");
                driver.FindElement(By.Id("input_invoice_postalcode")).SendKeys("24119");
                driver.FindElement(By.Id("input_invoice_city")).SendKeys("Kronshagen");
                driver.FindElement(By.Id("input_invoice_telephone")).SendKeys("0431 00012544");
                driver.FindElement(By.Id("input_invoice_mobilenumber")).SendKeys("0171-999888111");
                driver.FindElement(By.Id("input_invoice_email")).SendKeys("martin.wolters@tuicruises.com");
                driver.FindElement(By.Id("input_invoice_emailrepeat")).SendKeys("martin.wolters@tuicruises.com");


                footer = driver.FindElement(By.CssSelector("[class='booking-pager bottom']"));
                footer.FindElement(By.CssSelector("a[class^='button button-cta button-next-page']")).Click();

                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.CssSelector("[class^='spinner-container']")));

                // Zahlungsmethode
                //driver.FindElement(By.Id("input_paymenttype_paymentoption_0")).Click();
                driver.FindElement(By.CssSelector("[class^='js-item item']")).Click();

                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.CssSelector("[class^='spinner-container']")));

                //driver.FindElement(By.Id("input_overviewandapproval_approvalofterms")).Click();
                driver.FindElement(By.CssSelector("label[for='input_overviewandapproval_approvalofterms']")).Click();

                footer = driver.FindElement(By.CssSelector("[class='booking-pager bottom']"));
                footer.FindElement(By.CssSelector("a[class^='button button-cta button-next-page']")).Click();

                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.CssSelector("[class^='spinner-container']")));

                // Bankdaten
                driver.FindElement(By.Id("input_sepa_iban")).SendKeys("DE27100777770209299700");
                driver.FindElement(By.Id("input_sepa_bic")).SendKeys("THE BIC");
                driver.FindElement(By.Id("input_sepa_bankname")).SendKeys("SELENIUM BANK");

                driver.FindElement(By.CssSelector("label[for='input-accept']")).Click();
                //driver.FindElement(By.Id("input-accept")).Click();


                driver.FindElement(By.CssSelector("a[class^='button button-cta button-next-page']")).Click();

                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.CssSelector("[class^='spinner-container']")));



                Thread.Sleep(50000);
                driver.Quit();


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
                    //(range.Cells[i, 18] as Excel.Range).Value2 = "moin";


                    //Console.WriteLine(i + ": " + (range.Cells[i, 8] as Excel.Range).Value2);

                }
                //loader.saveWorkbook();
                //loader.quit();
                //wb.Save();
            }).Start();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            driver.Quit();
        }
    }
}
