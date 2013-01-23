using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web.UI;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System.Threading;
using System.Collections.ObjectModel;
using Automation.Development.Pages;
using System.Diagnostics;

namespace Automation.Development.Browsers
{
    /// <summary>
    /// This class contains browser related functionalities e.g. switch to window, find controls by different parameters
    /// </summary>
    public abstract class Browser
    {

        protected IWebDriver driver;

        public abstract void Navigate(string url);
        public abstract string PageTitle { get; }
                
        public Browser(IWebDriver driver)
        {
            this.driver = driver;
        }

        public IWebDriver Driver { get { return this.driver; } }

        #region This function finds control by selector
        /// <summary>
        /// This function finds control by selector
        /// </summary>
        /// <param name="selector"></param>
        /// <returns>webelement</returns>
        public virtual IWebElement FindControlBySelector(string selector)
        {
            try
            {
                return this.driver.FindElement(By.CssSelector(selector));
            }
            catch (Exception)
            {
            }
            return null;
        }
        #endregion

        #region This function finds control by Id
        /// <summary>
        /// This function finds control by Id
        /// </summary>
        /// <param name="key"></param>
        /// <returns>weblement</returns>
        public virtual IWebElement FindControlById(string key)
        {
            try
            {
                return this.driver.FindElement(By.Id(key));
            }
            catch (Exception)
            {
            }
            return null;
        }

        public void Close()
        {
            
            this.driver.Quit();
        }
        #endregion

        #region This function finds control by xpath
        /// <summary>
        /// Finds control by Xpath
        /// </summary>
        /// <param name="key"></param>
        /// <returns>webelement</returns>
        public virtual IWebElement FindControlByXpath(string key)
        {
            try
            {
                return this.driver.FindElement(By.XPath(key));
            }
            catch (Exception)
            {
            }
            return null;
        }
        #endregion

        #region This function finds control by xpath
        /// <summary>
        /// Finds select control by xpath
        /// </summary>
        /// <param name="key"></param>
        /// <returns>webelement</returns>
        public virtual SelectElement FindSelectControlByXpath(string key)
        {
            try
            {
                SelectElement select = new SelectElement(this.driver.FindElement(By.XPath(key)));
                return select;
            }
            catch (Exception)
            {
            }
            return null;
        }
        #endregion

        /// <summary>
        /// This function finds select control by selector
        /// </summary>
        /// <param name="selector"></param>
        /// <returns>webelement</returns>
        public virtual SelectElement FindSelectControlBySelector(string selector)
        {
            try
            {
                SelectElement select = new SelectElement(this.driver.FindElement(By.CssSelector(selector)));
                return select;
            }
            catch (Exception)
            {
            }
            return null;
        }

        /// <summary>
        /// This function finds select control by class
        /// </summary>
        /// <param name="className"></param>
        /// <returns>webelement</returns>
        public virtual SelectElement FindSelectControlByClass(string className)
        {
            try
            {
                SelectElement select = new SelectElement(this.driver.FindElement(By.ClassName(className)));
                return select;
            }
            catch (Exception)
            {
            }
            return null;
        }

        /// <summary>
        /// This function find select control by Id 
        /// </summary>
        /// <param name="key"></param>
        /// ll<returns>webelement</returns>
        public virtual SelectElement FindSelectControlById(string key)
        {
            try
            {
                SelectElement select = new SelectElement(this.driver.FindElement(By.Id(key)));
                return select;
            }
            catch (Exception)
            {
            }
            return null;
        }
		
        /// <summary>
        /// This function select control by name
        /// </summary>
        /// <param name="key"></param>
        /// <returns>webelement</returns>
		public virtual SelectElement FindSelectControlByName(string key)
        {
            try
            {
                SelectElement select = new SelectElement(this.driver.FindElement(By.Name(key)));
                return select;
            }
            catch (Exception)
            {
            }
            return null;
        }
        /// <summary>
        /// This function finds collection of read only control by xpath
        /// </summary>
        /// <param name="key"></param>
        /// <returns>webelements collection</returns>
        public virtual ReadOnlyCollection<IWebElement> FindControlsByXpath(string key)
        {
            try
            {
                return this.driver.FindElements(By.XPath(key));
            }
            catch (Exception)
            {
            }
            return null;
        }

        /// <summary>
        /// This finction finds control by class
        /// </summary>
        /// <param name="className"></param>
        /// <returns>webelement</returns>
        public virtual IWebElement FindControlByClass(string className)
        {
            try
            {
                return this.driver.FindElement(By.ClassName(className));
            }
            catch (Exception)
            {
            }
            return null;
        }
        /// <summary>
        /// This function finds collection of read only control by class
        /// </summary>
        /// <param name="className"></param>
        /// <returns>webelements collection</returns>
        public virtual ReadOnlyCollection<IWebElement> FindControlsByClass(string className)
        {
            try
            {
                return this.driver.FindElements(By.ClassName(className));
            }
            catch (Exception)
            {
            }
            return null;
        }
        /// <summary>
        /// This function finds control by tag name
        /// </summary>
        /// <param name="tagName"></param>
        /// <returns>webelement</returns>
        public virtual IWebElement FindControlByTagName(string tagName)
        {
            try
            {
                return this.driver.FindElement(By.TagName(tagName));
            }
            catch (Exception)
            {
            }
            return null;
        }

        /// <summary>
        /// This function finds collection of read only control by tagname
        /// </summary>
        /// <param name="tagName"></param>
        /// <returns>webelements collection</returns>
        public virtual ReadOnlyCollection<IWebElement> FindControlsByTagName(string tagName)
        {
            try
            {
                return this.driver.FindElements(By.TagName(tagName));
            }
            catch (Exception)
            {
            }
            return null;
        }

        /// <summary>
        /// This function finds control by name
        /// </summary>
        /// <param name="key"></param>
        /// <returns>webelement</returns>
        public virtual IWebElement FindControlByName(string key)
        {
            try
            {
                return this.driver.FindElement(By.Name(key));
            }
            catch (Exception)
            {
            }
            return null;
        }

        /// <summary>
        /// This function finds control by link text
        /// </summary>
        /// <param name="key"></param>
        /// <returns>webelement</returns>
        public virtual IWebElement FindControlByLinkText(string key)
        {
            try
            {
                return this.driver.FindElement(By.LinkText(key));
            }
            catch (Exception)
            {
            }
            return null;
        }

        /// <summary>
        /// This function find control by partial link text
        /// </summary>
        /// <param name="key"></param>
        /// <returns>webelement</returns>
        public virtual IWebElement FindControlByPartialLinkText(string key)
        {
            try
            {
                return this.driver.FindElement(By.PartialLinkText(key));
            }
            catch (Exception)
            {
            }
            return null;
        }

        /// <summary>
        /// This function switchs to alert window
        /// </summary>
        /// <returns></returns>
        public virtual IAlert SwitchToAlert()
        {
            try
            {
                return this.driver.SwitchTo().Alert();
            }
            catch (Exception)
            {
            }
            return null;
        }

        /// <summary>
        /// This function switches to window using index values
        /// </summary>
        /// <param name="windowHandleIndex"></param>
        public virtual void SwitchToWindowWithIndex(int windowHandleIndex)
        {
            try
            {
                this.driver.SwitchTo().Window(this.driver.WindowHandles[windowHandleIndex]);
            }
            catch (Exception)
            {
            }
        }

        /// <summary>
        /// This function finds collection of read only control by link text
        /// </summary>
        /// <param name="linkText"></param>
        /// <returns>webelement collection</returns>
        public virtual ReadOnlyCollection<IWebElement> FindControlsByLinkText(string linkText)
        {
            try
            {
                return this.driver.FindElements(By.LinkText(linkText));
            }
            catch (Exception)
            {
            }
            return null;
        }
    }
}
