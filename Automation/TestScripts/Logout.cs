using System;
using Automation.Development.Browsers;
using Automation.Development.Pages;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Net;
using System.IO;
using System.Collections;

namespace Automation.TestScripts
{
    [TestClass]
    public class Logout : TestCaseUtil
    {
        [TestMethod]
        public void zLogOut()
        {
            string embeddedMailContents = null;
            ArrayList notRunTestCases = new ArrayList();
            DateTime startTimeOfExecution = new DateTime();
            DateTime endTimeOfExecution = new DateTime();

            // Loads config data and creates a Singleton object of Configuration and loads data into generic test case variables
            GetConfigData();

            // Get debug viewer exe file path
            string configFilesLocation = PrepareConfigureDataFilePath();

            // Get log directory details from xml file
            PrepareLogDirectoryPath(configFilesLocation);
                        
            try
            {

                KillAlreadyOpenBrowsers();

                // to stop debeg viewer
                ApplicationLog applicationLog = new ApplicationLog();
                applicationLog.StopDebugViewer();
                
                string ExecutionStartDateTime = GetValuesFromXML("TestDataConfig", "ExecutionStartDateTime", configFilesLocation + "\\RunTime.xml");
                startTimeOfExecution = Convert.ToDateTime(ExecutionStartDateTime.ToString());
                endTimeOfExecution = DateTime.Now;

                // Delete run time config file
                DeleteConfigDetailsFile();

                // Generate custome html report form csv logs
                AutomationReport automationReport = new AutomationReport(logFileDirectory, reportFileDirectory, server, startTimeOfExecution, endTimeOfExecution);
                embeddedMailContents = automationReport.CompileReportFromCSV();

                // Prepare path for zip file
                string zippedFolderPath = reportFileDirectory + ".zip";

                // Zip the custome html report folder
                FileZipOperations fileZipOperations = new FileZipOperations(reportFileDirectory, zippedFolderPath, null);
                fileZipOperations.ZipFiles();

                // Send mail/ notification to given mail ids in config file
                Notifications notifications = new Notifications(reportFileDirectory, embeddedMailContents);
                notifications.SendNotification();
            }
            catch (Exception exception)
            {
                Console.WriteLine("Exception in Logout " + exception.Message);
            }
            finally
            {
                
            }
        }
    }
}
