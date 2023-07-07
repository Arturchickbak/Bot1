using System;
using System.IO;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

class Program
{
    static void Main()
    {
        string chromeDriverPath = "C:\\Users\\хозяин\\Desktop\\Botik";
        string excelFilePath = "C:\\Users\\хозяин\\Desktop\\Bot\\Записи.xlsx";

        using (IWebDriver driver = new ChromeDriver(chromeDriverPath))
        {
            driver.Navigate().GoToUrl("https://segezhsky--kar.sudrf.ru/modules.php?name=sud_delo&name_op=sf&delo_id=1540005"); // Замените "http://example.com" на адрес нужного сайта

            IWebElement caseNumberInput = driver.FindElement(By.Id("<input name=\"g1_case__CASE_NUMBERSS\" type=\"text\" class=\"Lookup\" size=\"70\">")); // Замените "caseNumberInput" на идентификатор поля ввода на сайте
            caseNumberInput.SendKeys("2-213/2023"); // Замените "Номер_дела" на нужный номер дела

            IWebElement searchButton = driver.FindElement(By.Id("searchButton")); // Замените "searchButton" на идентификатор кнопки на сайте
            searchButton.Click();

            System.Threading.Thread.Sleep(5000); // Ожидание загрузки информации (пример)

            IWebElement informationElement = driver.FindElement(By.CssSelector("<a href=\"#\" onclick=\"index(2, 4); return false;\">&nbsp;ДВИЖЕНИЕ ДЕЛА&nbsp;</a>")); // Замените ".information" на CSS-селектор элемента, содержащего информацию на сайте
            string information = informationElement.Text;

            if (!string.IsNullOrEmpty(information))
            {
                FileInfo excelFile = new FileInfo(excelFilePath);
                using (ExcelPackage package = new ExcelPackage(excelFile))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Изменения");
                    int lastRow = worksheet.Dimension?.Rows ?? 0;
                    worksheet.Cells[lastRow + 1, 1].Value = information;

                    package.Save();
                }
            }

            driver.Quit();
        }
    }
}
