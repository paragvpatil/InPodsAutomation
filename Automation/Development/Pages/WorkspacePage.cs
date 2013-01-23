using System;
using Automation.Development.Browsers;
using OpenQA.Selenium;
using System.Diagnostics;
using OpenQA.Selenium.Support.UI;
using System.Collections.ObjectModel;
using OpenQA.Selenium.Interactions;
using Automation.TestScripts;

namespace Automation.Development.Pages
{
    /// <summary>
    /// To use,access functionalities of login page
    /// </summary>
    public class WorkspacePage : HomePage
    {        

        //declaring variables to access web elements
        private IWebElement deleteButtonControl;
        private IWebElement activateButtonControl;

        /// <summary>
        /// Default constructor
        /// </summary>
        private WorkspacePage() { }

        /// <summary>
        /// Default Parameterized Constructor
        /// </summary>
        /// <param name="browser">browser value to store session</param>
        public WorkspacePage(Browser browser)
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

            bool isWorkSpaceTextPresent = this.WaitForElement("XPATH", (string)objectRepository.ObjectRepositoryTable["WorkSpaceText"]);
            if (!isWorkSpaceTextPresent)
            {
                throw new Exception("WorkSpace Text not found");
            }
            bool isWorkLoadTableDisplayed = this.WaitForElement("XPATH", (string)objectRepository.ObjectRepositoryTable["WorkloadTable"]);
            if (!isWorkLoadTableDisplayed)
            {
                throw new Exception("WorkLoad Table not found");
            }
            bool isDeleteButtonPresent = this.WaitForElement("XPATH", (string)objectRepository.ObjectRepositoryTable["DeleteButton"]);
            if (!isDeleteButtonPresent)
            {
                throw new Exception("Delete Button not found");
            }
            deleteButtonControl = this.FindControlByXPath((string)objectRepository.ObjectRepositoryTable["DeleteButton"]);
            bool isActivateButtonPresent = this.WaitForElement("XPATH", (string)objectRepository.ObjectRepositoryTable["ActivateButton"]);
            if (!isActivateButtonPresent)
            {
                throw new Exception("activate Button not found");
            }
            activateButtonControl = this.FindControlByXPath((string)objectRepository.ObjectRepositoryTable["ActivateButton"]);

        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public DashboardPage ActivateWorkload()
        {
            bool isActivateButtonDisplayed = this.WaitForElement("XPATH", (string)objectRepository.ObjectRepositoryTable["ActivateButton"]);
            if (isActivateButtonDisplayed)
            {
                activateButtonControl = this.FindControlByXPath((string)objectRepository.ObjectRepositoryTable["ActivateButton"]);
                Actions action = new Actions(this.Browser.Driver);
                action.MoveToElement(activateButtonControl).DoubleClick().Perform();
                ImplicitlyWait(5000);
                bool isActivateTrialButtonDisplayed = this.WaitForElement("XPATH", (string)objectRepository.ObjectRepositoryTable["ActivateTrialButton"]);
                if (isActivateTrialButtonDisplayed)
                {
                    IWebElement activateTrailButtonControl = this.FindControlByXPath((string)objectRepository.ObjectRepositoryTable["ActivateTrialButton"]);
                    action.MoveToElement(activateTrailButtonControl).DoubleClick().Perform();
                    ImplicitlyWait(10000);
                    return new DashboardPage(this.Browser);
                }
                else
                {
                    throw new Exception("Activate Trial Button not found");
                }
            }
            else
            {
                throw new Exception("Activate button not found");
            }
        }       


    }
}

