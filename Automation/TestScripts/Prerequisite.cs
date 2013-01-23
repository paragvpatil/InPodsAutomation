using System;
using Automation.Development.Browsers;
using Automation.Development.Pages;
using Automation.TestScripts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Threading;

namespace Automation.TestScripts
{
    [TestClass]
    public class Prerequisite : TestCase
    {

        [TestMethod]
        public void Prerequisites()
        {
            bool isMenuSelected = false;
            int stepNo = 0;
            int applicationNumber = 0;
            int indexNumber = 0;
            int fileNumber = 0;
            // Gets scriptname through reflection
            string testScriptName = GetTestScriptName();

            // Loads config data and creates a Singleton object of Configuration and loads data into generic test case variables
            GetConfigData();

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

            string applicationName = null;

            SessionHandler sessionHandler = new SessionHandler();
            Browser browser = sessionHandler.GetBrowserInstance();
            try
            {
                if (null == browser)
                {
                    browser = BrowserFactory.Instance.GetBrowser(browserId, testScriptName, configFilesLocation);
                    LoginPage loginPage = new LoginPage(browser);
                    stepNo++;
                    SearchPage searchPageNew = loginPage.Login(this.userName, this.password, this.applicationURL, this.timeout, processPath);
                    Assert.IsNotNull(searchPageNew, "Failed to login in application");
                    WriteLogs(testScriptName, stepNo, "Login succssfull.", "PASS", browser);

                }

                SearchPage searchPage = new SearchPage(browser);
                string emailAddress = GetValuesFromXML("Prerequisite", "EmailAddress", testDataFilePath);
                string userPassword = GetValuesFromXML("Prerequisite", "Password", testDataFilePath);

                stepNo++;
                isMenuSelected = searchPage.SelectMenuItem("Admin", "User Configurations", null);
                Assert.IsTrue(isMenuSelected, "Navigation to User Configurations page failed.");
                WriteLogs(testScriptName, stepNo, "Navigation to User Configurations page passed.", "PASS", browser);

                for (int count1 = 1; count1 <= 2; count1++)
                {
                    string userNode = "User";
                    userNode = userNode + count1.ToString();
                    string user = GetValuesFromXML("Prerequisite", userNode, testDataFilePath);

                    stepNo++;
                    UserConfigurationPage userConfigurationPage = new UserConfigurationPage(browser);
                    userConfigurationPage = userConfigurationPage.VerifyUserPresent(user);

                    if (null == userConfigurationPage)
                    {
                        stepNo++;
                        userConfigurationPage = new UserConfigurationPage(browser);
                        UserDetailsPage userDetailsPage = userConfigurationPage.ClickAddUserButton();
                        Assert.IsNotNull(userDetailsPage, "Click add user failed");
                        WriteLogs(testScriptName, stepNo, "Add user button Clicked", "PASS", browser);

                        stepNo++;
                        userConfigurationPage = userDetailsPage.AddUpdateUser(user, emailAddress, userPassword,true,false);
                        Assert.IsNotNull(userConfigurationPage, "Adding user failed");
                        WriteLogs(testScriptName, stepNo, "Adding user passed", "PASS", browser);

                        stepNo++;
                        UserPermissionPage userPermissionPage = userConfigurationPage.ClickAccessUserConfiguration(user);
                        Assert.IsNotNull(userPermissionPage, "Click User aceess button failed");
                        WriteLogs(testScriptName, stepNo, "Click User aceess button passed", "PASS", browser);

                        stepNo++;
                        userPermissionPage = userPermissionPage.UncheckResetPasswordNextLoginCheckbox();
                        Assert.IsNotNull(userPermissionPage, "Uncheck Reset Password Next Login Checkbox failed");
                        WriteLogs(testScriptName, stepNo, "Uncheck Reset Password Next Login Checkbox passed", "PASS", browser);

                        stepNo++;
                        userConfigurationPage = userPermissionPage.ClickUpdateButtonUserPermission();
                        Assert.IsNotNull(userConfigurationPage, "Click Update Button User Permission failed");
                        WriteLogs(testScriptName, stepNo, "Click Update Button User Permission passed", "PASS", browser);

                    }
                    else
                    {
                        WriteLogs(testScriptName, stepNo, "User already present", "PASS", browser);
                    }
                }
                string appliactionCount = GetValuesFromXML("Prerequisite", "ApplicationCount", testDataFilePath);
                int count = Int32.Parse(appliactionCount);
                for (applicationNumber = 1; applicationNumber <= count; applicationNumber++)
                {
                    string applicationNameNode = "Application";
                    applicationNameNode = applicationNameNode + applicationNumber.ToString();
                    applicationName = GetValuesFromXML(applicationNameNode, "Name", testDataFilePath);

                    //Navigating To Applications Page	
                    stepNo++;
                    isMenuSelected = searchPage.SelectMenuItem("Admin", "Applications", null);
                    Assert.IsTrue(isMenuSelected, "Navigation to application page failed.");
                    WriteLogs(testScriptName, stepNo, "Navigation to application page passed.", "PASS", browser);

                    stepNo++;
                    ApplicationPage applicationPage = new ApplicationPage(browser);
                    applicationPage = applicationPage.IsApplicationExist(applicationName);
                    if (null == applicationPage)
                    {
                        applicationPage = new ApplicationPage(browser);
                        AddApplicationPage addApplicationPage = applicationPage.ClickAddApplicationButton();
                        Assert.IsNotNull(addApplicationPage, "Click add application failed");
                        WriteLogs(testScriptName, stepNo, "Add Application Button Clicked", "PASS", browser);

                        stepNo++;
                        EditApplicationPage editApplicationPage = new EditApplicationPage(browser);
                        if (applicationNumber == 4)
                        {
                            editApplicationPage = addApplicationPage.addApplicationDetails(applicationName, applicationName, false, false, false, false, false, true, false, null);
                            Assert.IsNotNull(editApplicationPage, "Application Added Failed");
                            WriteLogs(testScriptName, stepNo, "Application Added Succesfully ", "PASS", browser);
                        }
                        else if (applicationNumber == 1 || applicationNumber == 5)
                        {
                            editApplicationPage = addApplicationPage.addApplicationDetails(applicationName, applicationName, false, false, false, false, false, false, false, null);
                            Assert.IsNotNull(editApplicationPage, "Application Added Failed");
                            WriteLogs(testScriptName, stepNo, "Application Added Succesfully ", "PASS", browser);
                        }
                        else
                        {
                            editApplicationPage = addApplicationPage.addApplicationDetails(applicationName, applicationName, true, false, false, false, false, false, false, null);
                            Assert.IsNotNull(editApplicationPage, "Application Added Failed");
                            WriteLogs(testScriptName, stepNo, "Application Added Succesfully ", "PASS", browser);
                        }
                        string indexCount = GetValuesFromXML(applicationNameNode, "IndexCount", testDataFilePath);
                        int count1 = Int32.Parse(indexCount);
                        //Adding Application Index
                        for (indexNumber = 1; indexNumber <= count1; indexNumber++)
                        {
                            string indexNameNade = "IndexName";
                            string indexTypeNade = "IndexType";
                            indexNameNade = indexNameNade + indexNumber.ToString();
                            indexTypeNade = indexTypeNade + indexNumber.ToString();
                            string indexName = GetValuesFromXML(applicationNameNode, indexNameNade, testDataFilePath);
                            string indexType = GetValuesFromXML(applicationNameNode, indexTypeNade, testDataFilePath);
                            stepNo++;
                            if (1 == indexNumber)
                            {

                                editApplicationPage = editApplicationPage.AddApplicationIndex(true, indexName, indexType, false, false, false, false);
                            }
                            else
                            {
                                editApplicationPage = editApplicationPage.AddApplicationIndex(false, indexName, indexType, false, false, false, false);
                            }
                            Assert.IsNotNull(editApplicationPage, "Adding application index failed");
                            WriteLogs(testScriptName, stepNo, "Index in application Added Succesfully ", "PASS", browser);
                        }
                        //Implementing Application
                        stepNo++;
                        editApplicationPage = editApplicationPage.ImplementApplication(true);
                        Assert.IsNotNull(editApplicationPage, "Implement application failed");
                        WriteLogs(testScriptName, stepNo, "Application Implemented Succesfully ", "PASS", browser);

                        //Navigating To User Configuration Page
                        stepNo++;
                        isMenuSelected = searchPage.SelectMenuItem("Admin", "User Configurations", null);
                        Assert.IsTrue(isMenuSelected, "Navigating to user configuration failed");
                        WriteLogs(testScriptName, stepNo, "Navigation to User Configurations page passed.", "PASS", browser);

                        stepNo++;
                        UserConfigurationPage userConfigurationPage = new UserConfigurationPage(browser);
                        UserPermissionPage userPermissionPage = userConfigurationPage.ClickAccessUserConfiguration(this.userName);
                        Assert.IsNotNull(userPermissionPage, "Click user Access link failed");
                        WriteLogs(testScriptName, stepNo, "Click on User Access on User Configuration Page ", "PASS", browser);

                        stepNo++;
                        ApplicationLevelPermissionPage applicationLevelPermissionPage = userPermissionPage.NavigateToApplicationConfiguration(applicationName);
                        Assert.IsNotNull(applicationLevelPermissionPage, "Click  on application configuration image failed");
                        WriteLogs(testScriptName, stepNo, "Click on Application Configuration image ", "PASS", browser);

                        stepNo++;
                        applicationLevelPermissionPage = applicationLevelPermissionPage.SelectAllAccountFuncMgnt();
                        Assert.IsNotNull(applicationLevelPermissionPage, "Select all account function management failed");
                        WriteLogs(testScriptName, stepNo, "Select All Account Management Checkbox Clicked ", "PASS", browser);

                        stepNo++;
                        applicationLevelPermissionPage = applicationLevelPermissionPage.SelectAllCabinetMgnt();
                        Assert.IsNotNull(applicationLevelPermissionPage, "Select all cabinet mangement failed");
                        WriteLogs(testScriptName, stepNo, "Select All Cabinet Management Checkbox Clicked ", "PASS", browser);

                        stepNo++;
                        userPermissionPage = applicationLevelPermissionPage.ClickUpdateButtonAppConf();
                        Assert.IsNotNull(userPermissionPage, "Update application level permissions failed");
                        WriteLogs(testScriptName, stepNo, "Update Application Level Permissions Clicked ", "PASS", browser);

                        stepNo++;
                        userConfigurationPage = userPermissionPage.ClickUpdateButtonUserPermission();
                        Assert.IsNotNull(userConfigurationPage, "Update user permissions failed");
                        WriteLogs(testScriptName, stepNo, "Update User Permissions Clicked ", "PASS", browser);

                        string user1 = GetValuesFromXML("Prerequisite", "User1", testDataFilePath);
                        string user2 = GetValuesFromXML("Prerequisite", "User2", testDataFilePath);

                        stepNo++;
                        userPermissionPage = userConfigurationPage.ClickAccessUserConfiguration(user1);
                        Assert.IsNotNull(userPermissionPage, "Click user Access link failed");
                        WriteLogs(testScriptName, stepNo, "Click on User Access on User Configuration Page ", "PASS", browser);

                        stepNo++;
                        userPermissionPage = userPermissionPage.CheckUncheckApplicationUserConf(applicationName, true);
                        Assert.IsNotNull(userPermissionPage, "Selecting Application for UserA failed");
                        WriteLogs(testScriptName, stepNo, "Selecting Application for UserA passed ", "PASS", browser);

                        stepNo++;
                        userConfigurationPage = userPermissionPage.ClickUpdateButtonUserPermission();
                        Assert.IsNotNull(userConfigurationPage, "Click Update Button failed");
                        WriteLogs(testScriptName, stepNo, "Click Update Button passed ", "PASS", browser);

                        stepNo++;
                        userPermissionPage = userConfigurationPage.ClickAccessUserConfiguration(user2);
                        Assert.IsNotNull(userPermissionPage, "Click user Access link failed");
                        WriteLogs(testScriptName, stepNo, "Click on User Access on User Configuration Page ", "PASS", browser);

                        stepNo++;
                        userPermissionPage = userPermissionPage.CheckUncheckApplicationUserConf(applicationName, true);
                        Assert.IsNotNull(userPermissionPage, "Selecting Application for UserA failed");
                        WriteLogs(testScriptName, stepNo, "Selecting Application for UserA passed ", "PASS", browser);

                        stepNo++;
                        userConfigurationPage = userPermissionPage.ClickUpdateButtonUserPermission();
                        Assert.IsNotNull(userConfigurationPage, "Click Update Button failed");
                        WriteLogs(testScriptName, stepNo, "Click Update Button passed ", "PASS", browser);

                        stepNo++;
                        userPermissionPage = userConfigurationPage.ClickAccessUserConfiguration(user1);
                        Assert.IsNotNull(userPermissionPage, "Click user Access link failed");
                        WriteLogs(testScriptName, stepNo, "Click on User Access on User Configuration Page ", "PASS", browser);

                        stepNo++;
                        applicationLevelPermissionPage = userPermissionPage.NavigateToApplicationConfiguration(applicationName);
                        Assert.IsNotNull(applicationLevelPermissionPage, "Click  on application configuration image failed");
                        WriteLogs(testScriptName, stepNo, "Click on Application Configuration image ", "PASS", browser);

                        stepNo++;
                        applicationLevelPermissionPage = applicationLevelPermissionPage.SelectAllAccountFuncMgnt();
                        Assert.IsNotNull(applicationLevelPermissionPage, "Select all account function management failed");
                        WriteLogs(testScriptName, stepNo, "Select All Account Management Checkbox Clicked ", "PASS", browser);

                        stepNo++;
                        applicationLevelPermissionPage = applicationLevelPermissionPage.SelectAllCabinetMgnt();
                        Assert.IsNotNull(applicationLevelPermissionPage, "Select all cabinet mangement failed");
                        WriteLogs(testScriptName, stepNo, "Select All Cabinet Management Checkbox Clicked ", "PASS", browser);

                        stepNo++;
                        userPermissionPage = applicationLevelPermissionPage.ClickUpdateButtonAppConf();
                        Assert.IsNotNull(userPermissionPage, "Update application level permissions failed");
                        WriteLogs(testScriptName, stepNo, "Update Application Level Permissions Clicked ", "PASS", browser);

                        stepNo++;
                        userConfigurationPage = userPermissionPage.ClickUpdateButtonUserPermission();
                        Assert.IsNotNull(userConfigurationPage, "Update user permissions failed");
                        WriteLogs(testScriptName, stepNo, "Update User Permissions Clicked ", "PASS", browser);

                        stepNo++;
                        userPermissionPage = userConfigurationPage.ClickAccessUserConfiguration(user2);
                        Assert.IsNotNull(userPermissionPage, "Click user Access link failed");
                        WriteLogs(testScriptName, stepNo, "Click on User Access on User Configuration Page ", "PASS", browser);

                        stepNo++;
                        applicationLevelPermissionPage = userPermissionPage.NavigateToApplicationConfiguration(applicationName);
                        Assert.IsNotNull(applicationLevelPermissionPage, "Click  on application configuration image failed");
                        WriteLogs(testScriptName, stepNo, "Click on Application Configuration image ", "PASS", browser);

                        stepNo++;
                        applicationLevelPermissionPage = applicationLevelPermissionPage.SelectAllAccountFuncMgnt();
                        Assert.IsNotNull(applicationLevelPermissionPage, "Select all account function management failed");
                        WriteLogs(testScriptName, stepNo, "Select All Account Management Checkbox Clicked ", "PASS", browser);

                        stepNo++;
                        applicationLevelPermissionPage = applicationLevelPermissionPage.SelectAllCabinetMgnt();
                        Assert.IsNotNull(applicationLevelPermissionPage, "Select all cabinet mangement failed");
                        WriteLogs(testScriptName, stepNo, "Select All Cabinet Management Checkbox Clicked ", "PASS", browser);

                        stepNo++;
                        userPermissionPage = applicationLevelPermissionPage.ClickUpdateButtonAppConf();
                        Assert.IsNotNull(userPermissionPage, "Update application level permissions failed");
                        WriteLogs(testScriptName, stepNo, "Update Application Level Permissions Clicked ", "PASS", browser);

                        stepNo++;
                        userConfigurationPage = userPermissionPage.ClickUpdateButtonUserPermission();
                        Assert.IsNotNull(userConfigurationPage, "Update user permissions failed");
                        WriteLogs(testScriptName, stepNo, "Update User Permissions Clicked ", "PASS", browser);

                        string fileCount = GetValuesFromXML(applicationNameNode, "FileCount", testDataFilePath);
                        int count2 = Int32.Parse(fileCount);
                        for (fileNumber = 1; fileNumber <= count2; fileNumber++)
                        {
                            //Navigating To upload Page
                            stepNo++;
                            isMenuSelected = userConfigurationPage.SelectMenuItem("Upload", null, null);
                            Assert.IsTrue(isMenuSelected, "Opening upload dialog failed");
                            WriteLogs(testScriptName, stepNo, "Navigation to upload page passed.", "PASS", browser);

                            string fileNameNode = "FileName";
                            string indexValeNode = "IndexValue" + fileNumber.ToString();
                            fileNameNode = fileNameNode + fileNumber.ToString();
                            string fileName = GetValuesFromXML(applicationNameNode, fileNameNode, testDataFilePath);
                            string indexValue = GetValuesFromXML(applicationNameNode, indexValeNode, testDataFilePath);
                            if (indexValue == "null")
                            {
                                indexValue = null;
                            }
                            string fileNameWithPath = this.currentDirectory + this.projectDirectory + "\\" + this.testDataDirectory + "\\" + testScriptName + "\\" + fileName;

                            stepNo++;
                            UploadFilePage uploadFilePage = new UploadFilePage(browser);
                            UploadIndexingPage uploadIndexingPage = uploadFilePage.UploadFile(applicationName, fileNameWithPath);
                            Assert.IsNotNull(uploadFilePage, "Uploading file " + fileNameWithPath + " failed");
                            WriteLogs(testScriptName, stepNo, "Upload File Successfull  ", "PASS", browser);

                            stepNo++;
                            searchPage = uploadIndexingPage.AddIndexDetails(indexValue, applicationName);
                            Assert.IsNotNull(searchPage, "Adding Index details failed");
                            WriteLogs(testScriptName, stepNo, "Upload index values succesfull  ", "PASS", browser);
                        }
                    }
                    else
                    {
                        stepNo++;
                        WriteLogs(testScriptName, stepNo, "Application exist", "PASS", browser);
                    }
                }

            }
            catch (Exception exception)
            {
                stepNo++;
                WriteLogs(testScriptName, stepNo, "Execution of script terminated. Exception is " + exception.Message, "FAIL", browser);
                Assert.IsTrue(false, "Execution of script terminated.");
            }
            finally
            {

                sessionHandler.StoreBrowserInstance(browser);
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
