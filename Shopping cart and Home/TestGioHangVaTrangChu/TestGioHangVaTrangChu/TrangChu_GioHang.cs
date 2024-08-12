using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace TestGioHangVaTrangChu
{
    [TestClass]
    public class UnitTest1
    {
        IWebDriver driver1;
        Excel.Application dataApp;
        Excel.Workbook dataWorkbook;
        Excel.Worksheet dataWorkSheet;
        Excel.Range xlRange;
        [TestInitialize]
        public void Init()
        {
            driver1 = new ChromeDriver();
            driver1.Url = "https://localhost:44310/";
            driver1.Navigate();
            driver1.Manage().Window.Maximize();
            Thread.Sleep(2000);

            dataApp = new Excel.Application();
            dataWorkbook = dataApp.Workbooks.Open(@"D:\LTBDCLPM\Nhóm 19\nhóm 19\Đoàn Quỳnh Như - 21DH112769\Đoàn Quỳnh Như_21DH112769_GioHang_TrangChu.xlsx");
            Thread.Sleep(2000);
        }
        [TestMethod]
        public void TestSearch()
        {
            dataWorkSheet = dataWorkbook.Sheets[2];
            xlRange = dataWorkSheet.UsedRange;

            int j = 3; //Chọn số dòng của j là data của dòng đó
            if (xlRange.Cells[2][j].value != null)
            {
                driver1.FindElement(By.XPath("//input[@name='SearchString']")).SendKeys(xlRange.Cells[2][j].value.ToString());
            }

            Thread.Sleep(2000);
            IWebElement btnlogin = driver1.FindElement(By.XPath("//input[@value='Tìm kiếm']"));
            btnlogin.Click();
            Thread.Sleep(2000);
        }

        [TestCleanup]
        public void CleanUp()
        {
            dataWorkbook.Save();
            dataWorkbook.Close();
            dataApp.Quit();
            driver1.Close();
        }
    }
}
