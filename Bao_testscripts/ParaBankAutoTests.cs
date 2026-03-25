using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace Bao_testscripts
{
    [TestFixture]
    public class Tests
    {
        private static string excelFilePath = @"D:\Nhom5_THBĐ\Report_Nhom_5.xlsx";
        private string screenshotFolder = @"D:\Nhom5_THBĐ\Screenshots";

        [SetUp]
        public void Setup()
        {
            if (!Directory.Exists(screenshotFolder)) Directory.CreateDirectory(screenshotFolder);
        }

        public static IEnumerable<TestCaseData> GetTestData()
        {
            var sheets = new[] { "Login", "Transfer", "BillPay", "Loan", "Account" };
            var testCases = new List<TestCaseData>();

            using (var stream = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var workbook = new XLWorkbook(stream))
                {
                    foreach (var sheetName in sheets)
                    {
                        if (!workbook.Worksheets.TryGetWorksheet(sheetName, out var worksheet)) continue;
                        var rows = worksheet.RowsUsed();
                        bool isFirstRow = true;
                        int currentRow = 0;
                        foreach (var row in rows)
                        {
                            currentRow++;
                            if (isFirstRow) { isFirstRow = false; continue; }
                            string testCaseId = row.Cell(1).GetString().Trim();
                            if (string.IsNullOrEmpty(testCaseId)) continue;

                            var data = new TestCaseData(
                                testCaseId, sheetName,
                                row.Cell(2).GetString(),
                                row.Cell(3).GetString(),
                                row.Cell(4).GetString(),
                                row.Cell(5).GetString(),
                                currentRow
                            );
                            data.SetName($"{testCaseId}_{sheetName}");
                            testCases.Add(data);
                        }
                    }
                }
            }
            return testCases;
        }

        [Test, TestCaseSource(nameof(GetTestData))]
        public void ExecuteAutoTest(string testCaseId, string sheetName, string username, string password, string amount, string expected, int rowIndex)
        {
            string actualMessage = "";
            bool isPass = false;

            using (IWebDriver driver = new ChromeDriver())
            {
                driver.Manage().Window.Maximize();
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(8); // Tăng thời gian đợi

                try
                {
                    driver.Navigate().GoToUrl("https://parabank.parasoft.com/parabank/index.htm");

                    // 1. LOGIN (Trừ trang Register)
                    if (sheetName != "Register")
                    {
                        driver.FindElement(By.Name("username")).SendKeys(username);
                        driver.FindElement(By.Name("password")).SendKeys(password);
                        driver.FindElement(By.CssSelector("input.button[value='Log In']")).Click();
                        Thread.Sleep(2000);
                    }

                    // 2. THỰC THI THEO SHEET
                    switch (sheetName)
                    {
                        case "Login":
                        case "Account":
                            if (sheetName == "Account")
                            {
                                try { driver.FindElement(By.LinkText("Accounts Overview")).Click(); Thread.Sleep(1000); } catch { }
                            }
                            actualMessage = GetPageResult(driver);
                            break;

                        case "Transfer":
                            driver.FindElement(By.LinkText("Transfer Funds")).Click();
                            Thread.Sleep(1500);
                            if (!string.IsNullOrEmpty(amount)) driver.FindElement(By.Id("amount")).SendKeys(amount);
                            driver.FindElement(By.CssSelector("input.button[value='Transfer']")).Click();
                            Thread.Sleep(1500);
                            actualMessage = GetPageResult(driver);
                            break;

                        case "BillPay":
                            driver.FindElement(By.LinkText("Bill Pay")).Click();
                            Thread.Sleep(1000);
                            driver.FindElement(By.Name("payee.name")).SendKeys("Test User");
                            driver.FindElement(By.Name("payee.address.street")).SendKeys("Street");
                            driver.FindElement(By.Name("payee.address.city")).SendKeys("City");
                            driver.FindElement(By.Name("payee.address.state")).SendKeys("State");
                            driver.FindElement(By.Name("payee.address.zipCode")).SendKeys("12345");
                            driver.FindElement(By.Name("payee.phoneNumber")).SendKeys("09090909");
                            driver.FindElement(By.Name("payee.accountNumber")).SendKeys("12345");
                            driver.FindElement(By.Name("verifyAccount")).SendKeys("12345");
                            if (!string.IsNullOrEmpty(amount)) driver.FindElement(By.Name("amount")).SendKeys(amount);
                            driver.FindElement(By.CssSelector("input.button[value='Send Payment']")).Click();
                            Thread.Sleep(1500);
                            actualMessage = GetPageResult(driver);
                            break;

                        case "Loan":
                            driver.FindElement(By.LinkText("Request Loan")).Click();
                            Thread.Sleep(1000);
                            if (!string.IsNullOrEmpty(amount)) driver.FindElement(By.Id("amount")).SendKeys(amount);
                            driver.FindElement(By.Id("downPayment")).SendKeys("10");
                            driver.FindElement(By.CssSelector("input.button[value='Apply Now']")).Click();
                            Thread.Sleep(1500);
                            actualMessage = GetPageResult(driver);
                            break;
                    }

                    // 3. LOGIC SO SÁNH THÔNG MINH
                    string cleanActual = actualMessage.ToLower();
                    string cleanExpected = expected.ToLower().Trim();

                    // Ưu tiên 1: Nếu mong đợi Accounts Overview và thực tế có chữ đó -> PASS
                    if (cleanExpected.Contains("accounts overview") && cleanActual.Contains("accounts overview"))
                    {
                        isPass = true;
                        actualMessage = "Accounts Overview displayed";
                    }
                    // Ưu tiên 2: So sánh chứa từ khóa
                    else if (cleanActual.Contains(cleanExpected))
                    {
                        isPass = true;
                    }
                    // Ưu tiên 3: Xử lý lỗi Internal Server của web (cho Fail nhưng ghi rõ)
                    else if (cleanActual.Contains("internal error"))
                    {
                        isPass = false;
                    }
                }
                catch (Exception ex)
                {
                    actualMessage = "Lỗi hệ thống: " + ex.Message;
                    isPass = false;
                }

                // GHI EXCEL
                WriteResultToExcel(sheetName, rowIndex, actualMessage, isPass, testCaseId, driver);

                // HIỂN THỊ ASSERT
                Assert.That(actualMessage, Does.Contain(expected).IgnoreCase, $"Mã Test: {testCaseId}");
            }
        }

        private void WriteResultToExcel(string sheetName, int rowIndex, string actual, bool isPass, string testCaseId, IWebDriver driver)
        {
            using (var stream = new FileStream(excelFilePath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
            {
                using (var workbook = new XLWorkbook(stream))
                {
                    var worksheet = workbook.Worksheet(sheetName);
                    worksheet.Cell(rowIndex, 6).Value = actual;
                    worksheet.Cell(rowIndex, 7).Value = isPass ? "PASS" : "FAIL";

                    if (!isPass)
                    {
                        string fileName = $"{testCaseId}_{DateTime.Now:HHmmss}.png";
                        string fullPath = Path.Combine(screenshotFolder, fileName);
                        ((ITakesScreenshot)driver).GetScreenshot().SaveAsFile(fullPath);
                        worksheet.Cell(rowIndex, 8).Value = "Link Ảnh";
                        worksheet.Cell(rowIndex, 8).SetHyperlink(new XLHyperlink(fullPath));
                    }
                    workbook.Save();
                }
            }
        }

        private string GetPageResult(IWebDriver driver)
        {
            // Tìm lỗi đỏ (validation) trước
            var spans = driver.FindElements(By.CssSelector("span.error"));
            foreach (var s in spans) if (!string.IsNullOrEmpty(s.Text)) return s.Text;

            // Tìm thông báo lỗi hệ thống
            var errors = driver.FindElements(By.ClassName("error"));
            foreach (var e in errors) if (!string.IsNullOrEmpty(e.Text)) return e.Text;

            // Lấy tiêu đề vùng nội dung (ví dụ: Accounts Overview, Transfer Complete)
            try
            {
                var title = driver.FindElement(By.ClassName("title")).Text;
                if (!string.IsNullOrEmpty(title)) return title;
            }
            catch { }

            // Cuối cùng mới lấy toàn bộ text
            try
            {
                return driver.FindElement(By.Id("rightPanel")).Text.Replace("\r", "").Replace("\n", " ").Trim();
            }
            catch
            {
                return "Không tìm thấy nội dung";
            }
        }
    }
}
