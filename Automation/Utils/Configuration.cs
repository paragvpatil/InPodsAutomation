using System;
using System.Text.RegularExpressions;
using System.Xml;


namespace Automation.TestScripts
{
    public class Configuration
    {
        private string testScriptName = String.Empty;
        private string server = String.Empty;
        private string browserId = String.Empty;
        private string applicationURL = String.Empty;
        private string userName = String.Empty;
        private string password = String.Empty;
        private int timeout = 0;
        private string webConfigFilePath = String.Empty;
        private string applicationLogDirectory = String.Empty;
        private string driverPath = String.Empty;       
        private string outlookPath = String.Empty;
        private string projectDirectory = String.Empty;
        private string testDataDirectory = String.Empty;
        private string testScriptDirectory = String.Empty;
        private string genericDataDirectory = String.Empty;
        private string preReqDataDirectory = String.Empty;
        private string customResultDirectory = String.Empty;
        private string htmlReportDirectory = String.Empty;
        private string csvLogDirectory = String.Empty;
        private string emailTo = String.Empty;
        private string emailFrom = String.Empty;
        private string emailSubject = String.Empty;
        private string emailMessage = String.Empty;
        private string emailSMTPServer = String.Empty;
        private string emailUserName = String.Empty;
        private string emailPassword = String.Empty;
        private string logFileDirectory = String.Empty;
        private string reportFileDirectory = String.Empty;
        private string currentDirectory = String.Empty;

        private static Configuration _instance = null;

        public Configuration()
        {
            LoadConfigData();
        }

        public Configuration Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new Configuration();
                }
                return _instance;
            }

        }

        // get set for server
        public string Server
        {
            get
            {
                return server;
            }
            set
            {
                server = value;
            }
        }

        // get set for browser type(internet explorer, firefox, chrome etc)
        public string BrowserId
        {
            get
            {
                return browserId;
            }
            set
            {
                browserId = value;
            }
        }

        // get set for server/application url
        public string ApplicationURL
        {
            get
            {
                return applicationURL;
            }
            set
            {
                applicationURL = value;
            }
        }

        // get set for user name
        public string UserName
        {
            get
            {
                return userName;
            }
            set
            {
                userName = value;
            }
        }

        // get set for password
        public string Password
        {
            get
            {
                return password;
            }
            set
            {
                password = value;
            }
        }

        // get set for timeout
        public int Timeout
        {
            get
            {
                return timeout;
            }
            set
            {
                timeout = value;
            }
        }

        // get set for config file path
        public string WebConfigFilePath
        {
            get
            {
                return webConfigFilePath;
            }
            set
            {
                webConfigFilePath = value;
            }
        }

        // get set for aplication log directory
        public string ApplicationLogDirectory
        {
            get
            {
                return applicationLogDirectory;
            }
            set
            {
                applicationLogDirectory = value;
            }
        }

        // get set for full text path
        public string DriverPath
        {
            get
            {
                return driverPath;
            }
            set
            {
                driverPath = value;
            }
        }
       
        public string OutlookPath
        {
            get
            {
                return outlookPath;
            }
            set
            {
                outlookPath = value;
            }
        }

        public string ProjectDirectory
        {
            get
            {
                return projectDirectory;
            }
            set
            {
                projectDirectory = value;
            }
        }

        public string TestDataDirectory
        {
            get
            {
                return testDataDirectory;
            }
            set
            {
                testDataDirectory = value;
            }
        }

        public string TestScriptDirectory
        {
            get
            {
                return testScriptDirectory;
            }
            set
            {
                testScriptDirectory = value;
            }
        }

        public string GenericDataDirectory
        {
            get
            {
                return genericDataDirectory;
            }
            set
            {
                genericDataDirectory = value;
            }
        }

        public string PreReqDataDirectory
        {
            get
            {
                return preReqDataDirectory;
            }
            set
            {
                preReqDataDirectory = value;
            }
        }
        
        public string CustomResultDirectory
        {
            get
            {
                return customResultDirectory;
            }
            set
            {
                customResultDirectory = value;
            }
        }

        public string HtmlReportDirectory
        {
            get
            {
                return htmlReportDirectory;
            }
            set
            {
                htmlReportDirectory = value;
            }
        }

        public string CsvLogDirectory
        {
            get
            {
                return csvLogDirectory;
            }
            set
            {
                csvLogDirectory = value;
            }
        }

        public string EmailTo
        {
            get
            {
                return emailTo;
            }
            set
            {
                emailTo = value;
            }
        }

        public string EmailFrom
        {
            get
            {
                return emailFrom;
            }
            set
            {
                emailFrom = value;
            }
        }

        public string EmailSubject
        {
            get
            {
                return emailSubject;
            }
            set
            {
                emailSubject = value;
            }
        }

        public string EmailMessage
        {
            get
            {
                return emailMessage;
            }
            set
            {
                emailMessage = value;
            }
        }

        public string EmailSMTPServer
        {
            get
            {
                return emailSMTPServer;
            }
            set
            {
                emailSMTPServer = value;
            }
        }

        public string EmailUserName
        {
            get
            {
                return emailUserName;
            }
            set
            {
                emailUserName = value;
            }
        }

        public string EmailPassword
        {
            get
            {
                return emailPassword;
            }
            set
            {
                emailPassword = value;
            }
        }

        public string CurrentDirectory
        {
            get
            {
                return currentDirectory;
            }
            set
            {
                currentDirectory = value;
            }
        }

        public string LogFileDirectory
        {
            get
            {
                return logFileDirectory;
            }
            set
            {
                logFileDirectory = value;
            }
        }

        public string ReportFileDirectory
        {
            get
            {
                return reportFileDirectory;
            }
            set
            {
                reportFileDirectory = value;
            }
        }

        public string TestScriptName
        {
            get
            {
                return testScriptName;
            }
            set
            {
                testScriptName = value;
            }
        }

        public void LoadConfigData()
        {
            try
            {
                // get current directory path
                CurrentDirectory = GetProjectLocation();

                // Prepare config file path
                string configFilePath = this.currentDirectory + "\\" + "Automation\\Test Data Index\\Generic Data\\ConfigData.xml";
                driverPath = this.currentDirectory + "Automation\\Test Data Index\\AutoITScripts\\";
                Server = GetValuesFromXML("TestEnvironment", "Server", configFilePath);
                BrowserId = GetValuesFromXML("TestEnvironment", "BrowserId", configFilePath);
                ApplicationURL = GetValuesFromXML("TestEnvironment", "ApplicationURL", configFilePath);
                UserName = GetValuesFromXML("TestEnvironment", "UserName", configFilePath);
                Password = GetValuesFromXML("TestEnvironment", "Password", configFilePath);
                Timeout = Int32.Parse(GetValuesFromXML("TestEnvironment", "TimeoutValue", configFilePath));
                WebConfigFilePath = GetValuesFromXML("TestEnvironment", "WebConfigFilePath", configFilePath);
                ApplicationLogDirectory = GetValuesFromXML("TestEnvironment", "ApplicationLogDirectory", configFilePath);
                
                OutlookPath = GetValuesFromXML("TestEnvironment", "OutlookPath", configFilePath);
                ProjectDirectory = GetValuesFromXML("TestResults", "ProjectDirectory", configFilePath);
                TestDataDirectory = GetValuesFromXML("TestResults", "TestDataDirectory", configFilePath);
                TestScriptDirectory = GetValuesFromXML("TestResults", "TestScriptDirectory", configFilePath);
                GenericDataDirectory = GetValuesFromXML("TestResults", "GenericDataDirectory", configFilePath);
                PreReqDataDirectory = GetValuesFromXML("TestResults", "PreReqDataDirectory", configFilePath);
                CustomResultDirectory = GetValuesFromXML("TestResults", "CustomResultDirectory", configFilePath);
                HtmlReportDirectory = GetValuesFromXML("TestResults", "HtmlReportDirectory", configFilePath);
                CsvLogDirectory = GetValuesFromXML("TestResults", "CsvLogDirectory", configFilePath);
                EmailTo = GetValuesFromXML("Notification", "EmailTo", configFilePath);
                EmailFrom = GetValuesFromXML("Notification", "EmailFrom", configFilePath);
                EmailSubject = GetValuesFromXML("Notification", "EmailSubject", configFilePath);
                EmailMessage = GetValuesFromXML("Notification", "EmailMessage", configFilePath);
                EmailSMTPServer = GetValuesFromXML("Notification", "EmailSMTPServer", configFilePath);
                EmailUserName = GetValuesFromXML("Notification", "EmailUserName", configFilePath);
                EmailPassword = GetValuesFromXML("Notification", "EmailPassword", configFilePath);
                logFileDirectory = this.currentDirectory + this.customResultDirectory + "\\" + this.csvLogDirectory + "\\" + this.csvLogDirectory + "_" + DateTime.Now.Year.ToString() + "_" + DateTime.Now.Month.ToString() + "_" + DateTime.Now.Day.ToString() + "_" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString();
                reportFileDirectory = this.currentDirectory + this.customResultDirectory + "\\" + this.htmlReportDirectory + @"\TestRun_" + DateTime.Now.Year.ToString() + "_" + DateTime.Now.Month.ToString() + "_" + DateTime.Now.Day.ToString() + "_" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString();

            }
            catch (Exception exception)
            {
                // Print exception on console
                Console.WriteLine("Exception in LoadConfigData " + exception.Message);
            }

        }

        #region Function to get project loation
        /// <summary>
        /// Function to get project location
        /// </summary>
        /// <returns>project location  path</returns>
        /// <author>rajan.bansod</author> 
        /// <ModifiedBy>rajan.bansod</ModifiedBy>
        /// <Date>12-Aug-2011</Date>
        public string GetProjectLocation()
        {
            string projectPath = null;
            try
            {
                // get result directory 
                projectPath = System.IO.Directory.GetCurrentDirectory();
                // split the path from 'TestResults' folder name which is auto-generated
                string[] projectPathArray = Regex.Split(projectPath, "TestResults");
                // return project path
                //return @"C:\QA\Automation\RefactoredCode\ComputeNextAutomationRefactored\";
                return projectPathArray[0];
            }
            catch (Exception exception)
            {
                // Print exception on console
                Console.WriteLine(exception.Message);
                //WriteLogs(TCID, STEP_NO, "GetProjectLocation " + exception.Message, FAIL,null);
                return projectPath;
            }
        }
        #endregion

        #region Function to get value from XML file
        /// <summary>
        ///Function to get values from XML file
        /// </summary>
        /// <returns>return node value</returns>
        /// <Date>03-Nov-2011</Date>
        public string GetValuesFromXML(string nodeName, string attributeName, string filePath)
        {
            string xmlNode = null;
            try
            {
                string xmlFilePath = null;

                // Prepare file path if it is null
                if (null == filePath)
                {
                    xmlFilePath = this.currentDirectory + this.projectDirectory + "\\" + this.testDataDirectory + "\\" + this.testScriptName + "\\" + this.testScriptName + ".xml";
                }
                else
                {
                    xmlFilePath = filePath;
                }

                // assign attribute as a normal to file
                System.IO.File.SetAttributes(xmlFilePath, System.IO.FileAttributes.Normal);
                XmlDocument doc = new XmlDocument();
                doc.Load(xmlFilePath);
                XmlElement root = doc.DocumentElement;
                XmlNodeList nodes = root.SelectNodes("//" + nodeName);
                if (nodes.Count < 1)
                {
                    Console.WriteLine("Xml Node not found");
                    return xmlNode;
                }
                //string xmlNode = nodes.ToString();
                foreach (XmlNode node in nodes)
                {
                    xmlNode = node[attributeName].InnerText;
                    return xmlNode;
                }
                return null;
            }
            catch (Exception exception)
            {
                // Print exception on console
                Console.WriteLine("Exception Occur In Read Xml File" + exception.Message);
                return "Exception Occur In Read Xml File" + exception.Message;
            }
        }
        #endregion

    }
}


