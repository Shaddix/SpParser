using OpenQA.Selenium;

namespace SpParser
{
    public class PageBase
    {
        public IWebDriver Driver { get; }

        public PageBase(IWebDriver driver)
        {
            Driver = driver;
        }

    }
}