using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using ClosedXML.Excel;
using System;
using System.IO;
using System.Threading;

namespace Bao_testscripts
{
    public class Tests
    {
        private string excelFilePath = @"C:\Tester\Nhom_5\Report_Nhom_5.xlsx";
        private string screenshotFolder = @"C:\Tester\Nhom_5\Screenshots";

        [SetUp]
        public void Setup()
        {
            if (!Directory.Exists(screenshotFolder)) Directory.CreateDirectory(screenshotFolder);
        }

        // TẠO 6 TESTCASE CHẠY RIÊNG BIỆT CHO 6 SHEET
        [Test] public void Test_01_Login() { RunTestsFromSheet("Login"); }
        [Test] public void Test_02_Transfer() { RunTestsFromSheet("Transfer"); }
        [Test] public void Test_03_BillPay() { RunTestsFromSheet("BillPay"); }
        [Test] public void Test_04_Loan() { RunTestsFromSheet("Loan"); }
        [Test] public void Test_05_Account() { RunTestsFromSheet("Account"); }
        [Test] public void Test_06_Register() { RunTestsFromSheet("Register"); } // THÊM MỚI Ở ĐÂY

        // --- HÀM XỬ LÝ LÕI ĐỌC EXCEL VÀ CHẠY SELENIUM ---
        private void RunTestsFromSheet(string sheetName)
        {
            using (var workbook = new XLWorkbook(excelFilePath))
            {
                // Kiểm tra xem Sheet có tồn tại không để tránh lỗi
                if (!workbook.Worksheets.TryGetWorksheet(sheetName, out var worksheet))
                {
                    Console.WriteLine($"Không tìm thấy Sheet: {sheetName}");
                    return;
                }

                var rows = worksheet.RowsUsed();
                bool isFirstRow = true;

                foreach (var row in rows)
                {
                    if (isFirstRow) { isFirstRow = false; continue; }

                    string testCaseId = row.Cell(1).GetString().Trim();
                    if (string.IsNullOrEmpty(testCaseId)) continue;

                    string username = row.Cell(2).GetString();
                    string password = row.Cell(3).GetString();
                    string amount = row.Cell(4).GetString(); // Cột Amount
                    string expected = row.Cell(5).GetString();

                    string actualMessage = "";
                    bool isPass = false;

                    using (IWebDriver driver = new ChromeDriver())
                    {
                        driver.Manage().Window.Maximize();
                        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);

                        try
                        {
                            driver.Navigate().GoToUrl("https://parabank.parasoft.com/parabank/index.htm");

                            // BƯỚC 1: NẾU KHÔNG PHẢI LÀ REGISTER THÌ MỚI ĐĂNG NHẬP
                            if (sheetName != "Register")
                            {
                                driver.FindElement(By.Name("username")).SendKeys(username);
                                driver.FindElement(By.Name("password")).SendKeys(password);
                                driver.FindElement(By.CssSelector("input.button[value='Log In']")).Click();
                                Thread.Sleep(1500);
                            }

                            // BƯỚC 2: RẼ NHÁNH THEO TÊN SHEET
                            switch (sheetName)
                            {
                                case "Register":
                                    // Bấm vào menu Register
                                    driver.FindElement(By.LinkText("Register")).Click();
                                    Thread.Sleep(1500);

                                    // Điền Dummy Data cho các ô thông tin cá nhân
                                    driver.FindElement(By.Id("customer.firstName")).SendKeys("Nguyen");
                                    driver.FindElement(By.Id("customer.lastName")).SendKeys("Bao");
                                    driver.FindElement(By.Id("customer.address.street")).SendKeys("123 Duong ABC");
                                    driver.FindElement(By.Id("customer.address.city")).SendKeys("HCM");
                                    driver.FindElement(By.Id("customer.address.state")).SendKeys("VN");
                                    driver.FindElement(By.Id("customer.address.zipCode")).SendKeys("70000");
                                    driver.FindElement(By.Id("customer.phoneNumber")).SendKeys("012345678");
                                    driver.FindElement(By.Id("customer.ssn")).SendKeys("123-45-678");

                                    // Điền thông tin Username và Password lấy từ Excel (Để test validation)
                                    if (!string.IsNullOrEmpty(username)) driver.FindElement(By.Id("customer.username")).SendKeys(username);
                                    if (!string.IsNullOrEmpty(password))
                                    {
                                        driver.FindElement(By.Id("customer.password")).SendKeys(password);
                                        driver.FindElement(By.Id("repeatedPassword")).SendKeys(password); // Confirm password luôn giống password
                                    }

                                    driver.FindElement(By.CssSelector("input.button[value='Register']")).Click();
                                    Thread.Sleep(1500);
                                    actualMessage = GetPageResult(driver, expected);
                                    break;

                                case "Login":
                                    actualMessage = GetPageResult(driver, expected);
                                    break;

                                case "Account":
                                    driver.FindElement(By.LinkText("Accounts Overview")).Click();
                                    Thread.Sleep(1500);
                                    actualMessage = GetPageResult(driver, expected);
                                    break;

                                case "Transfer":
                                    driver.FindElement(By.LinkText("Transfer Funds")).Click();
                                    Thread.Sleep(1500);
                                    if (!string.IsNullOrEmpty(amount)) driver.FindElement(By.Id("amount")).SendKeys(amount);
                                    driver.FindElement(By.CssSelector("input.button[value='Transfer']")).Click();
                                    Thread.Sleep(1500);
                                    actualMessage = GetPageResult(driver, expected);
                                    break;

                                case "BillPay":
                                    driver.FindElement(By.LinkText("Bill Pay")).Click();
                                    Thread.Sleep(1500);
                                    driver.FindElement(By.Name("payee.name")).SendKeys("Nguyen Van A");
                                    driver.FindElement(By.Name("payee.address.street")).SendKeys("123 Duong ABC");
                                    driver.FindElement(By.Name("payee.address.city")).SendKeys("HCM");
                                    driver.FindElement(By.Name("payee.address.state")).SendKeys("VN");
                                    driver.FindElement(By.Name("payee.address.zipCode")).SendKeys("70000");
                                    driver.FindElement(By.Name("payee.phoneNumber")).SendKeys("012345678");
                                    driver.FindElement(By.Name("payee.accountNumber")).SendKeys("12345");
                                    driver.FindElement(By.Name("verifyAccount")).SendKeys("12345");
                                    if (!string.IsNullOrEmpty(amount)) driver.FindElement(By.Name("amount")).SendKeys(amount);

                                    driver.FindElement(By.CssSelector("input.button[value='Send Payment']")).Click();
                                    Thread.Sleep(1500);
                                    actualMessage = GetPageResult(driver, expected);
                                    break;

                                case "Loan":
                                    driver.FindElement(By.LinkText("Request Loan")).Click();
                                    Thread.Sleep(1500);
                                    if (!string.IsNullOrEmpty(amount)) driver.FindElement(By.Id("amount")).SendKeys(amount);
                                    driver.FindElement(By.Id("downPayment")).SendKeys("10");

                                    driver.FindElement(By.CssSelector("input.button[value='Apply Now']")).Click();
                                    Thread.Sleep(1500);
                                    actualMessage = GetPageResult(driver, expected);
                                    break;
                            }

                            // SO SÁNH (PASS/FAIL)
                            if (actualMessage.Contains(expected))
                            {
                                isPass = true;
                                actualMessage = expected;
                            }
                        }
                        catch (Exception ex)
                        {
                            actualMessage = "Lỗi Exception: " + ex.Message;
                            isPass = false;
                        }

                        // GHI KẾT QUẢ VÀO EXCEL
                        row.Cell(6).Value = actualMessage; // Cột Actual
                        row.Cell(7).Value = isPass ? "PASS" : "FAIL"; // Cột Result

                        // XỬ LÝ CHỤP ẢNH
                        if (!isPass)
                        {
                            string fileName = $"{testCaseId}_{sheetName}_{DateTime.Now:HHmmss}.png";
                            string fullPath = Path.Combine(screenshotFolder, fileName);
                            ((ITakesScreenshot)driver).GetScreenshot().SaveAsFile(fullPath);

                            row.Cell(8).Value = "Xem ảnh lỗi";
                            row.Cell(8).SetHyperlink(new XLHyperlink(fullPath));
                            row.Cell(8).Style.Font.FontColor = XLColor.Blue;
                            row.Cell(8).Style.Font.Underline = XLFontUnderlineValues.Single;
                        }
                        else
                        {
                            row.Cell(8).Clear();
                        }
                    }
                }
                workbook.Save();
            }
        }

        // --- HÀM HỖ TRỢ LẤY THÔNG BÁO TỪ TRÌNH DUYỆT ---
        private string GetPageResult(IWebDriver driver, string expected)
        {
            var errorElements = driver.FindElements(By.CssSelector(".error"));

            if (errorElements.Count > 0 && !string.IsNullOrWhiteSpace(errorElements[0].Text))
            {
                return errorElements[0].Text;
            }

            if (expected == "Accounts Overview displayed" && (driver.PageSource.Contains("Accounts Overview") || driver.Url.Contains("overview.htm")))
            {
                return "Accounts Overview displayed";
            }

            try
            {
                string text = driver.FindElement(By.Id("rightPanel")).Text.Replace("\r", "").Replace("\n", " ");
                return text.Length > 60 ? text.Substring(0, 60) + "..." : text;
            }
            catch
            {
                return "Không lấy được thông báo (Lỗi Load trang)";
            }
        }
    }
}