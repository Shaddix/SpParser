using System.Collections.Generic;
using OpenQA.Selenium;

namespace SpParser
{
    public class LoginPage : PageBase
    {
        //[FindsBy(How = How.Id, Using = "UserName")]
        public IWebElement Login => Driver.FindElement(By.Id("username"));
        public IWebElement Password => Driver.FindElement(By.Name("password"));

        public IWebElement LoginButton => Driver.FindElement(By.XPath("//input[@type='submit' and @name='login']"));

        public LoginPage(IWebDriver driver) : base(driver)
        {
        }

        public void EnterLogin(string login, string pass)
        {
            Login.SendKeys(login);
            Password.SendKeys(pass);
            LoginButton.Click();
        }
    }
}