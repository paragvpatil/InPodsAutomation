using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;


namespace Automation.Development.Browsers
{
    public class Chrome : Browser
    {
        
        public Chrome(string driverPath)
            : base(new ChromeDriver(driverPath)) 
        {
        }
        /// <summary>
        /// Navigate to url
        /// </summary>
        /// <param name="url"></param>
        public override void Navigate(string url)
        {
            this.driver.Navigate().GoToUrl(url);
        }

        /// <summary>
        /// This function returns title of page
        /// </summary>
        public override string PageTitle
        {
            get 
            {
                return this.driver.Title;
            }
        }
    }
}
