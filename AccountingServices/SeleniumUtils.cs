using System;
using OpenQA.Selenium;

namespace AccountingServices
{
    public static class SeleniumUtils
    {
        public static bool IsElementPresent(IWebDriver driver, By by)
        {
            try
            {
                driver.FindElement(by);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }
    }
}
