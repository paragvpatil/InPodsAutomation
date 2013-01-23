using System;
using Automation.Development.Browsers;
using Automation.Development.Pages;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Automation.TestScripts
{
    /// <summary>
    /// Summary description for TestCase_2101
    /// </summary>
    [TestClass]
    public class TestCase_2101 : TestCase
    {
        [TestMethod]
        public void TestCase2101()
        {
            int stepNo = 1;

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
            ApplicationLog applicationLog = new ApplicationLog();
            applicationLog.StartDebugViewer(configFilesLocation, reportFileDirectory, testScriptName);

            // Prepare test data file path
            string testDataFilePath = PrepareTestDataFilePath(testScriptName);

            // Loads test case specific data
            TestData testData = new TestData(testDataFilePath);
            string applicationName = (string)testData.TestDataTable["ApplicationName"];
            string[] applicationsName = applicationName.Split(',');
            string pageNumber = (string)testData.TestDataTable["PageNumber"];
            string fileName = (string)testData.TestDataTable["FileName"];

            //file path to upload
            fileName = PrepareDocumentPath(testScriptName, fileName);

            CleanUpTestScripts.CleanUp_2100 cleanUp_2100 = new CleanUpTestScripts.CleanUp_2100();

            SessionHandler sessionHandler = new SessionHandler();
            Browser browser = sessionHandler.GetBrowserInstance();
            try
            {
                // Execute all clean up scripts if marked as true in xml file
                try
                {
                    // store session in session handler 
                    sessionHandler.StoreBrowserInstance(browser);
                    // Execute all clean up scripts if marked as true in xml file
                    CleanUpExecution cleanUpExecution = new CleanUpExecution();
                    cleanUpExecution.ExecuteCleanUp(configFilesLocation, this.projectDirectory);
                    // get browser session from session handler
                    browser = sessionHandler.GetBrowserInstance();
                }
                catch (Exception) { }
                //// call clean up script for test script
                //try
                //{
                //    sessionHandler.StoreBrowserInstance(browser);
                //    cleanUp_2100.CleanUp2100();
                //    browser = sessionHandler.GetBrowserInstance();
                //}
                //catch (Exception) { }
                try
                {
                    SearchPage searchPageNew = new SearchPage(browser);
                    bool isMenuSelectedFlag = searchPageNew.SelectMenuItem("Search", null, null);
                    Assert.IsTrue(isMenuSelectedFlag, "Navigation to search page failed.");
                    WriteLogs(testScriptName, stepNo, "Navigation to search page passed.", "PASS", browser);
                }
                catch (Exception)
                {
                    browser = null;
                }
                if (null == browser)
                {

                    browser = BrowserFactory.Instance.GetBrowser(browserId, testScriptName, configFilesLocation);
                    LoginPage loginPage = new LoginPage(browser);
                    SearchPage searchPageNew = loginPage.Login(this.userName, this.password, this.applicationURL, this.timeout, processPath);
                    searchPageNew.LocateControls();
                    Assert.IsNotNull(searchPageNew, "Failed to login in application");
                }

                WriteLogs(testScriptName, stepNo, "Verify ILS has been enabled for the application - in prerequisite", "PASS", browser);

                stepNo++;
                SearchPage searchPage = new SearchPage(browser);
                bool isSelected = searchPage.SelectMenuItem("Admin", "User Configurations", null);
                Assert.IsTrue(isSelected, "Failed to navigate to Search page");
                WriteLogs(testScriptName, stepNo, "Go to admin tab. Select user configuration. all iSynergy users are listed", "PASS", browser);

                stepNo++;
                UserConfigurationPage userConfigurationPage = new UserConfigurationPage(browser);
                UserPermissionPage userPermissionPage = new UserPermissionPage(browser);
                userPermissionPage = userConfigurationPage.ClickAccessUserConfiguration(this.userName);
                Assert.IsNotNull(userPermissionPage, "Failed to select the key next to sysadmin user");
                WriteLogs(testScriptName, stepNo, "select the key next to sysadmin user. User permissions screen appears. ", "PASS", browser);

                stepNo++;
                IndexLevelPermissionPage indexLevelPermissionPage = new IndexLevelPermissionPage(browser);
                indexLevelPermissionPage = userPermissionPage.NavigateToIndexLevelPermissionsPage(applicationName);
                Assert.IsNotNull(indexLevelPermissionPage,"Failed to select the Configuration gear icon under ILS configuration for the application "+applicationName);
                WriteLogs(testScriptName,stepNo,"Under application Permissions, select the Configuration gear icon under ILS configuration for the application in which ILS was enabled on An index field. Index Level Permissions configuration screen appears. Each ILS enabled index field for the application should be listed. A view column with check boxes and an edit column with check boxes Should be displayed. ","PASS",browser);

                stepNo++;
                indexLevelPermissionPage = indexLevelPermissionPage.SelectViewOrEditIndexLevelPermissions("View","ST1");
                Assert.IsNotNull(indexLevelPermissionPage, "View ILS options selected");

                userPermissionPage = indexLevelPermissionPage.ClickUpdateButtonAppConfILS();
                Assert.IsNotNull(userPermissionPage, "ILS Permission Update button clicked");

                userConfigurationPage = userPermissionPage.ClickUpdateButtonUserPermission();
                Assert.IsNotNull(userConfigurationPage, "User Permission Update button clicked");
                WriteLogs(testScriptName,stepNo,"Select the index field. Check the view check box To enable. Leave edit disabled. Click update button. ","PASS",browser);



            }
            catch (Exception exception)
            {
                ExceptionCleanUp(testScriptName, stepNo, exception.Message, browser);
            }
            finally
            {
                sessionHandler.StoreBrowserInstance(browser);
                //cleanUp_2100.CleanUp2100();
                stepNo++;
                // to stop debeg viewer
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
