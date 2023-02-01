using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using WebDriverManager.DriverConfigs.Impl;
using Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    { 
        public Form1()
        {
            InitializeComponent();
            
        }

        IWebDriver driver;
        public static Excel.Application xlApp = new Excel.Application();
        public static Excel.Application xlApp2 = new Excel.Application();
        public static Excel.Workbook xlWorkbook = null;
        public static Excel._Worksheet xlWorksheet = null;
        public static Excel.Range xlRange = null;

        public int rowCount = 0;
        public int colCount = 0;

        public static int i = 0;
        public static int j = 0;

        public string fileName;

        public void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            //dlg.ShowDialog();
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = dlg.FileName;
            }
        }
        
        public void button2_Click(object sender, EventArgs e)
        {
            xlWorkbook = xlApp.Workbooks.Open(textBox1.Text);
            xlWorksheet = xlWorkbook.Sheets[1];
            xlRange = xlWorksheet.UsedRange;
            rowCount = xlRange.Rows.Count;
            colCount = xlRange.Columns.Count;

            new WebDriverManager.DriverManager().SetUpDriver(new ChromeConfig());
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            

            for (int k = 3; k <= colCount; k++)
            {
                string loc = xlRange.Cells[1, k].Value2.ToString();
                driver.Url = "https://www.microsoft.com/" + loc + "/d/surface-pro-signature-keyboard/8qq3k7gn2tz5?activetab=pivot:overviewtab";
                Thread.Sleep(3000);
                driver.FindElement(By.XPath("//*[@aria-label='Cancel']")).Click();
                Thread.Sleep(5000);
                try
                {
                    driver.FindElement(By.XPath("//*[@class='close']")).Click();
                }
                catch
                {
                    Console.WriteLine("\nEmail popup not present for " + xlRange.Cells[1, k].Value2.ToString() + " locale\n");
                }
                IWebElement lang = driver.FindElement(By.XPath("//*[@lang='" + loc + "']"));
                string actual_lang = lang.GetAttribute("lang");
                Console.WriteLine(actual_lang);
                if (actual_lang == "en-US")
                {
                    test(3);
                }
                if (actual_lang == "en-CA")
                {
                    test(5);
                }
                if (actual_lang == "fr-CA")
                {
                    test(6);
                }
                if (actual_lang == "en-AU")
                {
                    test(4);
                }
                //test();
            }
            CloseBroswer();
        }

        public void test(int j)
        {
            try
            {
                for (i = 2; i <= rowCount - 1; i++)
                {
                    string image_row = xlRange.Cells[i, 2].Value2.ToString();
                    IWebElement test = driver.FindElement(By.XPath("//*[starts-with(@src, '" + image_row + "')]"));
                    string actual_alt = test.GetAttribute("alt");
                    string expected_alt = xlRange.Cells[i, j].Value2.ToString();
                    Console.WriteLine(actual_alt);
                    Assert.AreEqual(expected_alt, actual_alt, "Result Not Found");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine("The locale doesn't have needed info in input");
            }
        }

        public void CloseBroswer()
        {
            driver.Quit();
            GC.Collect();
            GC.WaitForPendingFinalizers();

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

        public void Form1_Load(object sender, EventArgs e)
        {

        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }
    }
}
