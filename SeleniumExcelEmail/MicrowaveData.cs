using System;
using System.Collections.Generic;
using System.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SeleniumExcelEmail
{
    class MicrowaveData
    {
        IWebDriver driver;
        WebDriverWait wait;

        public List<Data> MicrowaveDataSearch()
        {
            List<Data> dataList = new List<Data>();

            using (driver = new ChromeDriver())
            {
                wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));

                driver.Navigate().GoToUrl(@"https://onliner.by");

                wait.Until<bool>((d) => { return d.Title.Contains("Onliner"); });

                driver.FindElement(By.XPath("//*[@class='b-main-navigation__link']/span")).Click();

                driver.FindElement(By.XPath("//*[@class='catalog-navigation']/ul/li[4]")).Click();

                wait.Until(d => driver.FindElement(By.XPath("//*[@class='catalog-navigation']/div[2]/div/div[3]/div/div")));

                driver.FindElement(By.XPath("//*[@class='catalog-navigation']/div[2]/div/div[3]/div/div/div[6]")).Click();

                driver.FindElement(By.XPath("//*[@class='catalog-navigation']/div[2]/div/div[3]/div/div/div[6]/div[2]/div/a[1]")).Click();

                driver.FindElement(By.XPath("//*[@class='schema-order__link']")).Click();

                Thread.Sleep(1000);

                driver.FindElement(By.XPath("//*[@class='schema-order__list']/div[5]")).Click();

                Thread.Sleep(1000);

                var links = driver.FindElements(By.XPath(".//div[@class='schema-product__title']/a"));

                var prices = driver.FindElements(By.XPath(".//a[(contains(@class,'schema-product__price-value_primary')" +
                                                            "or contains(@class,'schema-product__price-value_additional')) " +
                                                            "and not(contains(@class,'schema-product__price-value_secondary'))]"));

                foreach (var item in links.Zip(prices, Tuple.Create))
                {
                    Data data = new Data();
                    data.name = item.Item1.Text.Remove(0, 19);
                    data.href = item.Item1.GetAttribute("href");
                    data.price = item.Item2.Text;
                    dataList.Add(data);
                }
            }

            foreach (var i in dataList)
            {
                Console.WriteLine(i.name + " " + i.price + " " + i.href + " ");
            }

            return dataList;


        }

    }
}
