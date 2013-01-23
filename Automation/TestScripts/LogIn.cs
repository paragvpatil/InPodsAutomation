using System;
using Automation.Development.Browsers;
using Automation.Development.Pages;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Automation.TestScripts;


namespace Automation.TestScripts
{
    [TestClass]
    public class LogIn : BaseClass
    {
        [TestMethod]
        public void TestCaseLogin()
        {

            int stepNo = 1;

            string test = TestContext.TestName;

            // Gets scriptname through reflection
            string testScriptName = GetTestScriptName();

            // Loads config data and creates a Singleton object of Configuration and loads data into generic test case variables
            this.GetConfigData();

            // Get process exe file path
            string[] processPath = PrepareProcessFilePath();

            // Get debug viewer exe file path
            string configFilesLocation = PrepareConfigureDataFilePath();

            // Get log directory details from xml file
            PrepareLogDirectoryPath(configFilesLocation);

            // Start debug viewer
            ApplicationLog applicationLog = new ApplicationLog(configFilesLocation, reportFileDirectory, testScriptName);
            applicationLog.StartDebugViewer();

            // Prepare test data file path
            string testDataFilePath = PrepareTestDataFilePath(testScriptName);

            // Loads test case specific data
            TestData testData = new TestData(testDataFilePath);
            string resourceType = (string)testData.TestDataTable["ResourceType"];
            string provider = (string)testData.TestDataTable["Provider"];
            string resourceDescription = (string)testData.TestDataTable["ResourceDescription"];
            SessionHandler sessionHandler = new SessionHandler();
            Browser browser = sessionHandler.GetBrowserInstance();

            try
            {
                try
                {
                    // store session in session handler 
                    sessionHandler.StoreBrowserInstance(browser);

                    // get browser session from session handler
                    browser = sessionHandler.GetBrowserInstance();
                }
                catch (Exception)
                { }

                VirtualMachinePage virtualMachinePage = null;
                LoginPage loginPage = null;

                try
                {

                    virtualMachinePage.LocateControls();

                }
                catch (Exception)
                {
                    browser = null;
                }

                if (null == browser)
                {

                    //Create Browser Instance 
                    browser = BrowserFactory.Instance.GetBrowser(browserId, testScriptName, configFilesLocation, driverPath);
                    //Login to ComputeNext Application
                    loginPage = new LoginPage(browser);
                    virtualMachinePage = loginPage.Login(this.userName, this.password, this.applicationURL, this.timeout, processPath);
                    virtualMachinePage.LocateControls();
                    Assert.IsNotNull(virtualMachinePage, "Failed to login to ComputeNext application ");
                    WriteLogs(testScriptName, stepNo, "Login to ComputeNext Application", "Pass", browser);

                    //Select Resource Type
                    stepNo++;
                    virtualMachinePage = virtualMachinePage.SelectResourceType(resourceType);
                    Assert.IsNotNull(virtualMachinePage, "Failed to Select Resource Type ");
                    WriteLogs(testScriptName, stepNo, "Resource Type Selected", "Pass", browser);

                    stepNo++;
                    virtualMachinePage = virtualMachinePage.SelectProvider(provider);
                    Assert.IsNotNull(virtualMachinePage, "Failed To Select Provider");
                    WriteLogs(testScriptName, stepNo, "Provider Selected", "Pass", browser);

                    stepNo++;
                    virtualMachinePage = virtualMachinePage.AddResourceToWorkSpace(resourceDescription);
                    Assert.IsNotNull(virtualMachinePage, "Failed To Add Resource To Workspace");
                    WriteLogs(testScriptName, stepNo, "Resource Added To Workspace", "Pass", browser);

                    stepNo++;
                    WorkspacePage workspacePage = virtualMachinePage.clickWorkSpaceButton();
                    Assert.IsNotNull(workspacePage, "Failed To Click Workspace Button");
                    WriteLogs(testScriptName, stepNo, "Workspace Button Clicked", "Pass", browser);

                    stepNo++;
                    DashboardPage DashboardPage = workspacePage.ActivateWorkload();
                    Assert.IsNotNull(DashboardPage, "Failed To Activate Workload");
                    WriteLogs(testScriptName, stepNo, "Activate Workload", "Pass", browser);

                    stepNo++;
                    //LogOut ComputeNext Application
                    bool isLoggedOut = loginPage.LogOut();
                    Assert.IsTrue(isLoggedOut, "Failed to Log Out to ComputeNext Application");
                    WriteLogs(testScriptName, stepNo, "Log Out to ComputeNext Application", "Pass", browser);
                }

            }
            catch (Exception exception)
            {
                WriteLogs(testScriptName, stepNo, exception.Message.ToString() + " Exception Occured in " + testScriptName, "FAIL", browser);
            }
            finally
            {
                sessionHandler.StoreBrowserInstance(browser);
                HomePage homePage = new HomePage(browser);
                homePage.LogOut();
                stepNo++;
                // to stop debug viewer
                applicationLog.StopDebugViewer();
                bool isExceptionFound = applicationLog.VerifyDebugLogFiles(reportFileDirectory, testScriptName);
                if (!isExceptionFound)
                {
                    WriteLogs(testScriptName, stepNo, "Exception/error found in log file", "INFO", browser);
                }
            }
        }

    }
}
