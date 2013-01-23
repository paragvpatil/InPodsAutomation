using System;
using Automation.Development.Browsers;
using OpenQA.Selenium;
using System.Diagnostics;
using OpenQA.Selenium.Support.UI;
using System.Collections.ObjectModel;
using Automation.TestScripts;

namespace Automation.Development.Pages
{
    /// <summary>
    /// To use,access functionalities of login page
    /// </summary>
    public class DashboardPage : HomePage
    {
        

        /// <summary>
        /// Default constructor
        /// </summary>
        private DashboardPage() { }

        /// <summary>
        /// Default Parameterized Constructor
        /// </summary>
        /// <param name="browser">browser value to store session</param>
        public DashboardPage(Browser browser)
            : base(browser)
        {
            
            this.LocateControls();
        }

        ObjectRepository objectRepository;

        /// <summary>
        /// To locate elements on login page using xpaths
        /// </summary>
        public void LocateControls()
        {
            string pageName = GetPageName();
            string objectRepositoryFilePath = PrepareObjectRepositoryFilePath(pageName);
            objectRepository = new ObjectRepository(objectRepositoryFilePath);
            bool isDashboardTextPresent = this.WaitForElement("XPATH", (string)objectRepository.ObjectRepositoryTable["DashboardText"]);
            if (!isDashboardTextPresent)
            {
                throw new Exception("Dashboard Text not found");
            }
           

        }



    }
}

