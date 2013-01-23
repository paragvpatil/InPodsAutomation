using OpenQA.Selenium.Firefox;

namespace Automation.Development.Browsers
{
    public class FireFox : Browser
    {
        public FireFox() : base(new FirefoxDriver())
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
