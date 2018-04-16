using System;
using System.Collections.Generic;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using excel=Microsoft.Office.Interop.Excel;


namespace Login_using_Excel_Sheets
{
    [TestClass]
    public class UnitTest1
    {
        public IWebDriver driver;

        [TestMethod]
        public void TestMethod1()
        {
           driver=new ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("https://rapgotesting.emcare.com");
            Thread.Sleep(5000);


            //Filling entries for Login
            driver.FindElement(By.XPath("//*[@id='hospitalSelect']/span/a")).Click();
            Thread.Sleep(1000);
            IList<IWebElement> option = driver.FindElements(By.ClassName("ui-menu-item"));
            foreach (var i in option)
                if (i.Text == "Cartersville Medical Center")
                {
                    i.Click();
                    break;
                }

            excel.Application x1app=new excel.Application();
            Thread.Sleep(2000);

            excel.Workbook x1Workbook = x1app.Workbooks.Open(@"C:\Users\Sachan\Desktop\test_data.xlsx");
            Thread.Sleep(2000);


            excel.Worksheet x1Worksheet = x1Workbook.Sheets[1];
            Thread.Sleep(2000);

            excel.Range x1range = x1Worksheet.UsedRange;
            excel.Range x2range = x1Worksheet.UsedRange;

            string username,password;
            try
            {
                for (int i = 1; i <= 3; i++)
                {
                    username = x1range.Cells[i][1].Value2;
                    driver.FindElement(By.XPath("//*[@id='UserName']")).SendKeys(username);
                    Thread.Sleep(1000);
                    password = x2range.Cells[i][2].Value2;
                    driver.FindElement(By.XPath("//*[@id='Password']")).SendKeys(password);
                    driver.FindElement(By.XPath("//*[@id='btnLogin']")).Click();
                    Thread.Sleep(4000);

                    if (driver.FindElement(By.XPath("//*[@id='logoutUser']/span/a")).Displayed)
                    {
                        break;
                    }
                }
            }//break
            catch (Exception e)
            {

            }





            //driver.FindElement(By.XPath("//*[@id='UserName']")).SendKeys("asimmons");
            //driver.FindElement(By.XPath("//*[@id='Password']")).SendKeys("Xyz123!");
           // driver.FindElement(By.XPath("//*[@id='btnLogin']")).Click();

            driver.Close();
        }
    }
}
