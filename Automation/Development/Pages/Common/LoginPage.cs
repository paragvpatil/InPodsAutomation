using System;
using Automation.Development.Browsers;
using OpenQA.Selenium;
using System.Diagnostics;
using Automation.TestScripts;

namespace Automation.Development.Pages
{
    /// <summary>
    /// To use,access functionalities of login page
    /// </summary>
    public class LoginPage : HomePage
    {
        //declaring variables to access web elements        
        private IWebElement userNameControl;
        private IWebElement passwordControl;
        private IWebElement loginButtonControl;
        private ObjectRepository objectRepository;

        /// <summary>
        /// Default constructor
        /// </summary>
        private LoginPage() { }

        /// <summary>
        /// Default Parameterized Constructor
        /// </summary>
        /// <param name="browser">browser value to store session</param>
        public LoginPage(Browser browser)
            : base(browser)
        {
            
        }

        #region Function for login to Application
        /// <summary>
        /// Function for logging into Application 
        /// </summary>
        /// <param name="userName">String value for name of user</param>
        /// <param name="password">String value for password</param>
        /// <param name="loginPageUrl">string value for url of login page</param>
        /// <param name="timeoutParam">integer numberfor time-out</param>
        /// <param name="processPath">String array for path of process</param>
        /// <returns>an instance of searchpage</returns>
        public VirtualMachinePage Login(string userName, string password, string loginPageUrl, int timeoutParam, string[] processPath)
        {
            try
            {
                if (string.IsNullOrEmpty(userName))
                {
                    throw new ArgumentNullException("userName");
                }

                if (string.IsNullOrEmpty(password))
                {
                    throw new ArgumentNullException("password");
                }
                timeoutWait = timeoutParam * 2;
                try
                {
                    string[] processName = { "QTAgentHandler", "QTAgentApplicationErrorHandler", "UploadFileScriptError", "InternetExplorerStopped" };
                    for (int count = 0; count <= processName.Length; count++)
                    {
                        Process[] localByName = Process.GetProcessesByName(processName[count]);
                        foreach (Process item in localByName)
                        {
                            // kill the process
                            item.Kill();
                        }
                    }
                }
                catch (Exception) { }
                StackTrace stackTrace = new StackTrace();
                for (int count = 0; count <= processPath.Length; count++)
                {
                    try
                    {
                        Process process = new Process();
                        process.StartInfo.FileName = processPath[count];
                        process.EnableRaisingEvents = true;
                        process.Start();
                    }
                    catch (Exception) { }
                }
                this.Browser.Driver.Navigate().GoToUrl(loginPageUrl);
                try
                {
                    this.Browser.Driver.Manage().Cookies.DeleteAllCookies();
                }
                catch (Exception) { }
                this.MaximizeWindowWithAltD();
                //try
                //{
                //    IWebElement MarketPlaceTabControl = this.FindControlByXPath((string)objectRepository.ObjectRepositoryTable["MarketPlaceTab"]);
                //    if (MarketPlaceTabControl != null)
                //    {
                //        this.LogOut();
                //    }
                //}
                //catch (Exception) { }                

                LocateControls();

               
                //entering the required username and password


                this.userNameControl = this.FindControlByXPath((string)objectRepository.ObjectRepositoryTable["UserNameTextbox"]);
                userNameControl.SendKeys(userName);
                this.passwordControl = this.FindControlByXPath((string)objectRepository.ObjectRepositoryTable["PasswordTextbox"]);
                passwordControl.SendKeys(password);
                this.loginButtonControl = this.FindControlByXPath((string)objectRepository.ObjectRepositoryTable["LoginButton"]);
                
                loginButtonControl.Click();
                bool isLoginTextboxPresent = this.WaitForElement("XPATH", (string)objectRepository.ObjectRepositoryTable["UserNameTextbox"]);
                if(isLoginTextboxPresent)
                {

                    this.userNameControl = this.FindControlByXPath((string)objectRepository.ObjectRepositoryTable["UserNameTextbox"]);
                    userNameControl.SendKeys(userName);
                    this.passwordControl = this.FindControlByXPath((string)objectRepository.ObjectRepositoryTable["PasswordTextbox"]);
                    passwordControl.SendKeys(password);
                    this.loginButtonControl = this.FindControlByXPath((string)objectRepository.ObjectRepositoryTable["LoginButton"]);
                
                    loginButtonControl.Click();
                
                }
               
                return new VirtualMachinePage(this.Browser);
            }
            catch (Exception exception)
            {
                throw new Exception("Exception in login " + exception.Message);
            }
        }

        #endregion
        /// <summary>
        /// To locate elements on login page using xpaths
        /// </summary>
        private void LocateControls()
        {
            string pageName = GetPageName();
            string objectRepositoryFilePath = PrepareObjectRepositoryFilePath(pageName);
            objectRepository = new ObjectRepository(objectRepositoryFilePath);


            bool isUserNameTextBoxPresent = this.WaitForElement("XPATH", (string)objectRepository.ObjectRepositoryTable["UserNameTextbox"]);
            if (!isUserNameTextBoxPresent)
            {
                throw new Exception("User name  Textbox not found"); 
            }
            userNameControl = this.FindControlByXPath((string)objectRepository.ObjectRepositoryTable["UserNameTextbox"]);
            bool isPasswordTextboxPresent = this.WaitForElement("XPATH", (string)objectRepository.ObjectRepositoryTable["PasswordTextbox"]);
            if (!isPasswordTextboxPresent)
            {
                throw new Exception("Password Textbox not found");
            }
            passwordControl = this.FindControlByXPath((string)objectRepository.ObjectRepositoryTable["PasswordTextbox"]);
            bool isLoginButtonPresent = this.WaitForElement("XPATH", (string)objectRepository.ObjectRepositoryTable["LoginButton"]);           
            if (!isLoginButtonPresent)
            {
                throw new Exception("Login Button not found");
            }
            loginButtonControl = this.FindControlByXPath((string)objectRepository.ObjectRepositoryTable["LoginButton"]);
        }

      
    }
}
