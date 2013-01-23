using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Automation.Development.Browsers;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System.Collections.ObjectModel;
using System.Threading;
using System.Runtime.InteropServices;

namespace Automation.Development.Pages
{
    /// <summary>
    /// Covers all functionality related to browser session i.e. store browser session
    /// </summary>
    public class SessionHandler
    {
        //local static variables
        private static SessionHandler _instance = null;
        private static Browser browserSession = null;

        public SessionHandler() { }

        //getter method to get browser session
        public static SessionHandler Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new SessionHandler();
                }
                return _instance;
            }
        }

        #region Function to get browser session
        /// <summary>
        /// Function to get browser session
        /// </summary>
        /// <returns>return browser session</returns>
        public Browser GetBrowserInstance()
        {
            try
            {
                //gets common xpath for all pages of ComputeNext application 
                string xpathAspDotNetform = "//*[@id=\"aspnetForm\"]";
                //checking browser session available or not
                IWebElement aspDotNetForm = browserSession.Driver.FindElement(By.XPath(xpathAspDotNetform));
                if (aspDotNetForm != null)
                {
                    return browserSession;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception)
            {
                return null;
            }
        }
        #endregion

        #region Function to store browser session
        /// <summary>
        /// Function to store browser session
        /// </summary>
        /// <param name="browser">browser session to store</param>
        public void StoreBrowserInstance(Browser browser)
        {
            try
            {
                //gets common xpath for all pages of ComputeNext application 
                string xpathAspDotNetform = "//*[@id=\"aspnetForm\"]";
                //checking browser session available or not
                IWebElement aspDotNetForm = browser.Driver.FindElement(By.XPath(xpathAspDotNetform));
                if (aspDotNetForm != null)
                {
                    browserSession = browser;
                }
                else
                {
                    browserSession = null;
                }
            }
            catch (Exception)
            {
                browserSession = null;
            }
        }
        #endregion
    }
}
