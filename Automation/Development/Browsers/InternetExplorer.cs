using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenQA.Selenium.IE;

namespace Automation.Development.Browsers
{
    public class InternetExplorer : Browser
    {
        public InternetExplorer(string driverPath)
            : base(new InternetExplorerDriver(driverPath))
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
            get { return this.driver.Title; }
        }
    }
}
