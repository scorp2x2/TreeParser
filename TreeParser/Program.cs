using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading;

namespace TreeParser
{
    internal class Program
    {
        public struct Row
        {
            public string number;
            public string sellerName;
            public string sellerInn;
            public string buyerName;
            public string buyerInn;
            public DateTime dealDate;
            public string volume;
        }

        static SqlConnection myConn;

        static void Main(string[] args)
        {
            try
            {
                myConn = new SqlConnection("Data source = DESKTOP-8EG4NM4; Initial catalog = WoodTrades; Integrated security=True");
                myConn.Open();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

            IWebDriver driver = new ChromeDriver();

            while (true)
            {
                driver.Navigate().GoToUrl("https://www.lesegais.ru/open-area/deal");
                Thread.Sleep(2000);

                var countText = driver.FindElement(By.XPath(".//*[@id='root']/div/div/div/div[2]/div/div/div[3]/div[2]/span[3]")).Text.Split(' ').Last();
                var countPages = int.Parse(countText);

                List<IWebElement> listOfElements;
                Row row = new Row();

                for (int j = 0; j < countPages; j++)
                {
                    listOfElements = driver.FindElements(By.XPath(".//div[@role = 'gridcell']")).ToList();

                    for (int i = 0; i < listOfElements.Count; i += 7)
                    {
                        var number = listOfElements[i].Text;
                        if (number == "" || number == "Итого")
                            continue;
                        row.number = number;
                        row.sellerName = listOfElements[i + 1].Text;
                        row.sellerInn = listOfElements[i + 2].Text;
                        row.buyerName = listOfElements[i + 3].Text;
                        row.buyerInn = listOfElements[i + 4].Text;

                        var date = listOfElements[i + 5].Text;
                        if (date != "")
                            row.dealDate = DateTime.Parse(date);
                        row.volume = listOfElements[i + 6].Text;

                        AddRow(row);
                    }

                    var button = driver.FindElement(By.XPath(".//span[@class = 'x-btn-icon-el x-btn-icon-el-plain-toolbar-small x-tbar-page-next']"));
                    button.Click();

                    Thread.Sleep(700);
                }

                Thread.Sleep(10 * 60 * 1000);
            }

            driver.Quit();
        }

        static void AddRow(Row row)
        {
            if (myConn.State != ConnectionState.Open) return;

            string query = $"SELECT count(id) FROM trades WHERE id = @number";
            SqlCommand command = new SqlCommand(query, myConn);
            command.Parameters.AddWithValue("@number", row.number);

            var count = (int)command.ExecuteScalar();

            if (count == 0)
            {
                query = $"INSERT INTO trades VALUES (@number, @buyerName ,@buyerInn, @sellerName, @sellerInn, @dealDate, @volume)";
                command = new SqlCommand(query, myConn);

                command.Parameters.AddWithValue("number", row.number);
                command.Parameters.AddWithValue("buyerName", row.buyerName);
                command.Parameters.AddWithValue("buyerInn", row.buyerInn);
                command.Parameters.AddWithValue("sellerName", row.sellerName);
                command.Parameters.AddWithValue("sellerInn", row.sellerInn);
                command.Parameters.AddWithValue("dealDate", row.dealDate);
                command.Parameters.AddWithValue("volume", row.volume);

                command.ExecuteNonQuery();
                Console.WriteLine($"{DateTime.Now}: Добавлена новая запись - {row.number}");
            }
        }
    }
}
