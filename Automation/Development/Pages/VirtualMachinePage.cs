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
    public class VirtualMachinePage : HomePage
    {
        //declaring variables to access xpaths
        
        //declaring variables to access web elements
        private IWebElement virtualMachineTextControl;
        private SelectElement resourceDropDownControl;

        /// <summary>
        /// Default constructor
        /// </summary>
        private VirtualMachinePage() { }

        /// <summary>
        /// Default Parameterized Constructor
        /// </summary>
        /// <param name="browser">browser value to store session</param>
        public VirtualMachinePage(Browser browser)
            : base(browser)
        {
            LocateControls();

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
            bool isVirtualMachineTextPresent = this.WaitForElement("XPATH", (string)objectRepository.ObjectRepositoryTable["VirtualMachineText"]);
            if (!isVirtualMachineTextPresent)
            {
                throw new Exception("Virtual Machine Text not found");
            }
            virtualMachineTextControl = this.FindControlByXPath((string)objectRepository.ObjectRepositoryTable["VirtualMachineText"]);
            bool isresourceDropDownPresent = this.WaitForElement("XPATH", (string)objectRepository.ObjectRepositoryTable["ResourceTypeDropdown"]);
            if (!isresourceDropDownPresent)
            {
                throw new Exception("Resource DropDown not found");
            }
            resourceDropDownControl = this.FindSelectControlByXpath((string)objectRepository.ObjectRepositoryTable["ResourceTypeDropdown"]);
           
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="resourceType"></param>
        /// <returns></returns>

        public VirtualMachinePage SelectResourceType(string resourceType)
        {
            resourceDropDownControl.SelectByText(resourceType);
            ImplicitlyWait(3000);
            bool isVirtualMachineTextPresent = this.WaitForElement("XPATH", (string)objectRepository.ObjectRepositoryTable["VirtualMachineText"]);
            if (!isVirtualMachineTextPresent)
            {
                throw new Exception("Virtual Machine Text not found");
            }
            return new VirtualMachinePage(this.Browser);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="providerName"></param>
        /// <returns></returns>

        public VirtualMachinePage SelectProvider(string providerName)
        {
            bool isProviderSelected = false;
            int tdCount = this.GetTableCount((string)objectRepository.ObjectRepositoryTable["ProviderDiv"]);
            IWebElement firstDivControl = this.FindControlByXPath((string)objectRepository.ObjectRepositoryTable["ProviderDiv"]);
            if (!firstDivControl.Displayed)
            {
                IWebElement providerLinkControl = this.FindControlByXPath((string)objectRepository.ObjectRepositoryTable["ProviderLink"]);
                providerLinkControl.Click();
                bool isFirstDivDisplayed = this.WaitForElementToGetEnabled("XPATH", (string)objectRepository.ObjectRepositoryTable["ProviderDiv"]);
                if (!isFirstDivDisplayed)
                {
                    throw new Exception("Providers not displayed");
                }
            }

            for (int providerCount = 1; providerCount <= tdCount; providerCount++)
            {
                string providerCheckbox = (string)objectRepository.ObjectRepositoryTable["ProviderDiv"] + "[" + providerCount + "]/input";
                IWebElement providerCheckboxControl = this.FindControlByXPath(providerCheckbox);
                string providerText = providerCheckboxControl.GetAttribute("value");
                if (providerName.Equals(providerText))
                {
                    providerCheckboxControl.Click();
                    isProviderSelected = true;
                    break;
                }
            }
            if (isProviderSelected)
            {
                bool isVirtualMachineTextPresent = this.WaitForElement("XPATH", (string)objectRepository.ObjectRepositoryTable["VirtualMachineText"]);
                if (!isVirtualMachineTextPresent)
                {
                    throw new Exception("Virtual Machine Text not found");
                }
                return new VirtualMachinePage(this.Browser);
            }
            else
            {
                throw new Exception("Provider " + providerName + " not found");
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="resourceDescription"></param>
        /// <returns></returns>
        public VirtualMachinePage AddResourceToWorkSpace(string resourceDescription)
        {
            bool isVMAddedToWorkspace = false;
            int vmCount = this.GetTableCount((string)objectRepository.ObjectRepositoryTable["VMList"]);
            ReadOnlyCollection<IWebElement> vmListControl = this.FindControlsByXPath((string)objectRepository.ObjectRepositoryTable["VMList"]);           
            for (int vm = 1; vm <= vmCount; vm++)
            {
                string vmDescription = vmListControl[vm].Text.ToString();
                IWebElement vmDescriptionDiv = vmListControl[vm].FindElement(By.TagName("div"));
                ReadOnlyCollection<IWebElement> listOfDiv = vmDescriptionDiv.FindElements(By.TagName("div"));
                vmDescriptionDiv = vmDescriptionDiv.FindElement(By.TagName("div"));               
                vmDescription = vmDescriptionDiv.Text.ToString();
                if(vmDescription.Contains(resourceDescription))
                {
                    IWebElement addToWorkSpaceButton = listOfDiv[2].FindElement(By.TagName("button"));
                    string buttonText = addToWorkSpaceButton.Text.ToString();
                    if (buttonText == "Add to Workspace")
                    {
                        addToWorkSpaceButton.Click();
                    }
                    else
                    {
                        throw new Exception("Add to Workspace button not found instead Button with Text " + buttonText+" found ");
                    }
                    isVMAddedToWorkspace = true;
                    bool isResourceAddedSuccessMessage = this.WaitForElement("XPATH", (string)objectRepository.ObjectRepositoryTable["ResourceAddedSuccessMessage"]);
                    if (!isResourceAddedSuccessMessage)
                    {
                        throw new Exception("Resource Added Success Message not found");
                    }
                    break;
                }
            
            }
            if (isVMAddedToWorkspace)
            {
                return new VirtualMachinePage(this.Browser);
            }
            else
            {
                throw new Exception("VM with Description "+resourceDescription+" not found");
            }
        }
      
    }
}

