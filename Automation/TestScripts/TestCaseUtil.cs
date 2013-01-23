using System;
using System.Drawing.Imaging;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml;
using Automation.Development.Browsers;
using OpenQA.Selenium;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using System.Linq;
using System.Collections;


namespace Automation.TestScripts
{
    public class TestCaseUtil
    {
        protected string testScriptName = String.Empty;
        protected string server = String.Empty;
        protected string browserId = String.Empty;
        protected string applicationURL = String.Empty;
        protected string userName = String.Empty;
        protected string password = String.Empty;
        protected int timeout = 0;
        protected string webConfigFilePath = String.Empty;
        protected string applicationLogDirectory = String.Empty;
        protected string driverPath = String.Empty;
        protected string outlookPath = String.Empty;
        protected string projectDirectory = String.Empty;
        protected string testDataDirectory = String.Empty;
        protected string testScriptDirectory = String.Empty;
        protected string genericDataDirectory = String.Empty;
        protected string preReqDataDirectory = String.Empty;
        protected string customResultDirectory = String.Empty;
        protected string htmlReportDirectory = String.Empty;
        protected string csvLogDirectory = String.Empty;
        protected string emailTo = String.Empty;
        protected string emailFrom = String.Empty;
        protected string emailSubject = String.Empty;
        protected string emailMessage = String.Empty;
        protected string emailSMTPServer = String.Empty;
        protected string emailUserName = String.Empty;
        protected string emailPassword = String.Empty;
        protected string logFileDirectory = String.Empty;
        protected string reportFileDirectory = String.Empty;
        protected string currentDirectory = String.Empty;
        
        public TestCaseUtil()
        {

        }

        public void GetConfigData()
        {
            try
            {
                Configuration oConfig = new Configuration();
                this.currentDirectory = oConfig.CurrentDirectory;
                this.server = oConfig.Server;
                this.browserId = oConfig.BrowserId;
                this.applicationURL = oConfig.ApplicationURL;
                this.userName = oConfig.UserName;
                this.password = oConfig.Password;
                this.timeout = oConfig.Timeout;
                this.webConfigFilePath = oConfig.WebConfigFilePath;
                this.applicationLogDirectory = oConfig.ApplicationLogDirectory;
                this.driverPath = oConfig.DriverPath;
                this.outlookPath = oConfig.OutlookPath;
                this.projectDirectory = oConfig.ProjectDirectory;
                this.testDataDirectory = oConfig.TestDataDirectory;
                this.testScriptDirectory = oConfig.TestScriptDirectory;
                this.genericDataDirectory = oConfig.GenericDataDirectory;
                this.preReqDataDirectory = oConfig.PreReqDataDirectory;
                this.customResultDirectory = oConfig.CustomResultDirectory;
                this.htmlReportDirectory = oConfig.HtmlReportDirectory;
                this.csvLogDirectory = oConfig.CsvLogDirectory;
                this.emailTo = oConfig.EmailTo;
                this.emailFrom = oConfig.EmailFrom;
                this.emailSubject = oConfig.EmailSubject;
                this.emailMessage = oConfig.EmailMessage;
                this.emailSMTPServer = oConfig.EmailSMTPServer;
                this.emailUserName = oConfig.EmailUserName;
                this.emailPassword = oConfig.EmailPassword;
                this.logFileDirectory = oConfig.LogFileDirectory;
                this.reportFileDirectory = oConfig.ReportFileDirectory;
            }
            catch (Exception exception)
            {
                Console.WriteLine("Exception in LoadConfigData " + exception.Message);
            }
        }

        #region Function to get script name
        /// <summary>
        /// Function to get current test script name
        /// </summary>
        /// <returns>script name</returns>        
        protected string GetTestScriptName()
        {
            string scriptName = "";
            try
            {
                // get class name 
                scriptName = this.ToString();

                // replace folder names from scriptName
                scriptName = scriptName.Replace("Automation.", "");
                scriptName = scriptName.Replace("TestScripts.", "");

                // return scriptName
                return scriptName;
            }
            catch (Exception exception)
            {
                // write log
                Console.WriteLine(exception.Message);
                return scriptName;
            }
        }
        #endregion

        protected string PrepareTestDataFilePath(string testScriptName)
        {
            return this.currentDirectory + this.projectDirectory + "\\" + this.testDataDirectory + "\\" + testScriptName + "\\" + testScriptName + ".xml";
        }

        protected string PrepareDocumentPath(string testScriptName, string fileName)
        {
            return this.currentDirectory + this.projectDirectory + "\\" + this.testDataDirectory + "\\" + testScriptName + "\\" + fileName;
        }

        protected string PrepareDocumentPath(string fileName)
        {
            return this.currentDirectory + this.projectDirectory + "\\" + this.testDataDirectory + "\\" + this.preReqDataDirectory + "\\" + fileName;
        }

        protected string PrepareConfigureDataFilePath()
        {
            return this.currentDirectory + this.projectDirectory + "\\" + this.testDataDirectory + "\\" + this.genericDataDirectory;
        }

        protected string[] PrepareProcessFilePath()
        {
            string[] processPath = new string[4];
            processPath[0] = this.currentDirectory + this.projectDirectory + "\\" + this.testDataDirectory + "\\AutoITScripts\\QTAgentHandler.exe";
            processPath[1] = this.currentDirectory + this.projectDirectory + "\\" + this.testDataDirectory + "\\AutoITScripts\\QTAgentApplicationErrorHandler.exe";
            processPath[2] = this.currentDirectory + this.projectDirectory + "\\" + this.testDataDirectory + "\\AutoITScripts\\UploadFileScriptError.exe";
            processPath[3] = this.currentDirectory + this.projectDirectory + "\\" + this.testDataDirectory + "\\AutoITScripts\\InternetExplorerStopped.exe";
            return processPath;
        }

        protected void PrepareLogDirectoryPath(string configFilesLocation)
        {
            string runTimeConfigFilePath = configFilesLocation + "\\RunTime.xml";
            string customResultDirectoryStored = GetValuesFromXML("TestDataConfig", "CustomResultDirectory", runTimeConfigFilePath);
            if (null != customResultDirectoryStored)
            {
                string htmlReportDirectoryStored = GetValuesFromXML("TestDataConfig", "HtmlReportDirectory", runTimeConfigFilePath);
                this.logFileDirectory = customResultDirectoryStored;
                this.reportFileDirectory = htmlReportDirectoryStored;
            }
            else
            {
                // Set log directory details to xml file
                SetValuesToXML("TestDataConfig", "CustomResultDirectory", this.logFileDirectory);
                SetValuesToXML("TestDataConfig", "HtmlReportDirectory", this.reportFileDirectory);
                SetValuesToXML("TestDataConfig", "ExecutionStartDateTime", DateTime.Now.ToString());

                // Delete all files under a temporary files directory
                DeleteTempFiles();
            }
        }
                
        #region Function to Write to CSV Log file
        /// <summary>
        /// Function to write logs to a csv file
        /// </summary>
        /// <param name="testID">Test case XPATH</param>
        /// <param name="stepNum">Step Num</param>
        /// <param name="stepDescription">Error description</param>
        /// <param name="stepResult">Result of step</param>
        /// <author>rajan.bansod</author> 
        /// <ModifiedBy>rajan.bansod</ModifiedBy>
        /// <Date>12-Aug-2011</Date>
        public void WriteLogs(string testID, int stepNum, string stepDescription, string stepResult, Browser browser)
        {
            StreamWriter logWriter = null;
            try
            {
                // check if log directory exist or not
                string filePath = this.logFileDirectory;
                //Console.WriteLine("filePath " + filePath);
                try
                {
                    if (!Directory.Exists(filePath))
                    {
                        // create directory
                        Directory.CreateDirectory(filePath);
                    }
                }
                catch (Exception exception)
                {
                    Console.WriteLine(exception.Message);
                }

                if ("FAIL".Equals(stepResult))
                {
                    // check if html report directory exist or not
                    try
                    {
                        if (!Directory.Exists(this.reportFileDirectory))
                        {
                            // create directory
                            Directory.CreateDirectory(this.reportFileDirectory);
                        }
                    }
                    catch (Exception exception)
                    {
                        Console.WriteLine(exception.Message);
                    }

                    // Code to take screen shot
                    CaptureScreenShot(testID + "_" + stepNum.ToString(), this.reportFileDirectory + "\\" + testID);
                    if (browser != null)
                    {
                        TakeScreenShot(testID + "_" + stepNum.ToString(), this.reportFileDirectory + "\\" + testID, browser);
                    }
                }
                // convert description string in to readable string, remove ',' from string
                stepDescription = ConvertToLogDescription(stepDescription);
                // append csv file name to directory path
                filePath = filePath + @"\" + testID + ".csv";
                //Console.WriteLine("log file filePath " + filePath);
                // check if log file exist or not
                if (File.Exists(filePath))
                {
                    // append data to the log file
                    logWriter = File.AppendText(filePath);
                    logWriter.WriteLine(testID + "," + stepNum.ToString() + "," + stepDescription + "," + stepResult + "," + DateTime.Now);
                    // close logWriter
                    logWriter.Close();
                    // write detials to console for visual studio result
                    Console.WriteLine(testID + "," + stepNum.ToString() + "," + stepDescription + "," + stepResult + "," + DateTime.Now);
                }
                else
                {
                    // add header to the log file
                    logWriter = File.AppendText(filePath);
                    logWriter.WriteLine("TCID,STEPNO,DESCRIPTION,RESULT,DATE_TIME");
                    // append data to the log file
                    logWriter.WriteLine(testID + "," + stepNum.ToString() + "," + stepDescription + "," + stepResult + "," + DateTime.Now);
                    // close logWriter
                    logWriter.Close();
                    // write detials to console for visual studio result
                    Console.WriteLine(testID + "," + stepNum.ToString() + "," + stepDescription + "," + stepResult + "," + DateTime.Now);
                }
                GenerateLog(testID, stepNum, stepDescription, stepResult);
            }
            catch (Exception exception)
            {
                Console.WriteLine("WriteLog " + exception.Message);
            }
        }
        #endregion

        #region Function to generate Log file
        /// <summary>
        /// Function to generate logs to a csv file
        /// </summary>
        /// <param name="testID">Test case XPATH</param>
        /// <param name="stepNum">Step Num</param>
        /// <param name="stepDescription">Error description</param>
        /// <param name="stepResult">Result of step</param>
        /// <Date>12-Aug-2011</Date>
        private void GenerateLog(string testID, int stepNum, string stepDescription, string stepResult)
        {
            StreamWriter logWriter = null;
            try
            {
                // append csv file name to directory path
                string filePath = this.logFileDirectory + @"\Logs.csv";
                // check if log file exist or not
                if (File.Exists(filePath))
                {
                    // append data to the log file
                    logWriter = File.AppendText(filePath);
                    logWriter.WriteLine(testID + "," + stepNum.ToString() + "," + stepDescription + "," + stepResult + "," + DateTime.Now);
                    // close logWriter
                    logWriter.Close();
                }
                else
                {
                    // add header to the log file
                    logWriter = File.AppendText(filePath);
                    logWriter.WriteLine("TCID,STEPNO,DESCRIPTION,RESULT,DATE_TIME");
                    // append data to the log file
                    logWriter.WriteLine(testID + "," + stepNum.ToString() + "," + stepDescription + "," + stepResult + "," + DateTime.Now);
                    // close logWriter
                    logWriter.Close();
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine("Exception in generate log " + exception.Message);
            }
        }
        #endregion

        #region Function to Delete all files under a temporary files directory
        /// <summary>
        /// Function to delete temporary files
        /// </summary>
        /// <Date>04-April-2012</Date>
        protected void DeleteTempFiles()
        {
            try
            {
                // Get temp folder path
                string sPath = System.IO.Path.GetTempPath();
                // Store file path to the string array
                string[] files = System.IO.Directory.GetFiles(sPath);
                // Iterte through all array elemnts
                foreach (string file in files)
                {
                    try
                    {
                        // set permissions to normal of each file
                        System.IO.File.SetAttributes(file, System.IO.FileAttributes.Normal);
                        // delete file
                        System.IO.File.Delete(file);
                    }
                    catch { }
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine("Exception in delete temp files " + exception.Message);
            }
        }
        #endregion 

        #region Function to take ScreenShot
        /// <summary>
        /// Function to take ScreenShot
        /// </summary>
        /// <param name="argument">test step no</param>
        /// <returns>bool</returns>
        /// <author>rajan.bansod</author> 
        /// <ModifiedBy>rajan.bansod</ModifiedBy>
        /// <Date>06-Sept-2011</Date>
        private bool TakeScreenShot(string argument, string currentResultDirectory, Browser browser)
        {
            try
            {
                if (!Directory.Exists(currentResultDirectory))
                {
                    // create directory
                    Directory.CreateDirectory(currentResultDirectory);
                }
                string filePath = currentResultDirectory + "\\" + argument + ".png";

                IWebDriver ITakesScreenshotDriver = browser.Driver;
                Screenshot ScreenShotObject = ((ITakesScreenshot)ITakesScreenshotDriver).GetScreenshot();
                string screenshot = ScreenShotObject.AsBase64EncodedString;
                byte[] screenshotAsByteArray = ScreenShotObject.AsByteArray;
                ScreenShotObject.SaveAsFile(filePath, ImageFormat.Png);
                return true;
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
                return false;
            }
        }
        #endregion

        #region Function to take ScreenShot
        /// <summary>
        /// Function to take ScreenShot
        /// </summary>
        /// <param name="argument">test step no</param>
        /// <returns>bool</returns>
        /// <author>rajan.bansod</author> 
        /// <ModifiedBy>rajan.bansod</ModifiedBy>
        /// <Date>06-Sept-2011</Date>
        private bool CaptureScreenShot(string argument, string currentResultDirectory)
        {
            Bitmap bitmap = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
            Graphics graphics = Graphics.FromImage(bitmap as Image);
            try
            {
                if (!Directory.Exists(currentResultDirectory))
                {
                    // create directory
                    Directory.CreateDirectory(currentResultDirectory);
                }
                string filePath = currentResultDirectory + "\\" + argument + "_Captured.png";
                graphics.CopyFromScreen(0, 0, 0, 0, bitmap.Size);
                bitmap.Save(filePath, ImageFormat.Png);
                return true;
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
                return false;
            }
            finally
            {
                bitmap.Dispose();
                graphics.Dispose();
            }
        }
        #endregion

        #region function for convert exception object to string [ depricated: use exception.Message ]
        /// <summary>
        /// Function to convert exception object to readable string
        /// </summary>
        /// <param name="exceptionObject">exception object</param>
        /// <returns>exception string</returns>
        /// <Date>09-Nov-2011</Date>
        public string ConvertToLogDescription(string exceptionString)
        {
            // Declare objects
            string[] stringArray = new string[10];

            try
            {
                // Verify if object is null
                if (exceptionString != null)
                {
                    stringArray = Regex.Split(exceptionString, "\r\n");
                    exceptionString = stringArray[0].ToString().Trim();
                    exceptionString = Regex.Replace(exceptionString, @"\r\n?|\n", " ");
                    exceptionString = exceptionString.Replace(",", " ");
                }
                return exceptionString;
            }
            catch (Exception exception)
            {
                Console.WriteLine("ConvertToLogDescription " + exception.Message);
                return "Failed to log exceptoin.";
            }
        }
        #endregion

        #region Function to set values in XML file
        /// <summary>
        /// Function to set values in XML file
        /// </summary>
        /// <returns>script name</returns>
        /// <Date>02-Nov-2011</Date>
        public bool SetValuesToXML(string nodeName, string attributeName, string attributeValue)
        {
            try
            {
                string value = null;
                string fileFolderPath = PrepareConfigureDataFilePath();
                string filePath = fileFolderPath + "\\" + "RunTime.xml";
                //Console.WriteLine("filePath " + filePath);
                bool isFileExists = File.Exists(filePath);
                if (isFileExists)
                {
                    System.IO.File.SetAttributes(filePath, System.IO.FileAttributes.Normal);
                    value = GetValuesFromXML(nodeName, attributeName, filePath);
                    if (value == attributeValue)
                    {
                        return true;
                    }
                    else if (value != null && value != attributeValue)
                    {
                        bool flag = false;
                        XmlDocument XMLEdit = new XmlDocument();
                        XMLEdit.Load(filePath);

                        XmlElement XMLEditNode = XMLEdit.DocumentElement;

                        foreach (XmlNode node in XMLEditNode) // Loop through XML file
                        {
                            if (node.Name == nodeName) // Check for the nodeName field information
                            {
                                XmlNode childNode = node.FirstChild;
                                XmlNode lastNode = node.LastChild;
                                if (childNode.Name == attributeName)
                                {
                                    childNode.InnerText = attributeValue;
                                    flag = true;
                                    break;
                                }
                                else
                                {
                                    do
                                    {
                                        childNode = childNode.NextSibling;
                                        if (childNode.Name == attributeName)
                                        {
                                            childNode.InnerText = attributeValue;
                                            flag = true;
                                            break;
                                        }
                                    }
                                    while (childNode == lastNode);
                                }
                            }
                        }
                        if (flag == true)
                        {
                            XMLEdit.Save(filePath);
                            return true;
                        }
                        else
                        {
                            return false;
                        }
                    }
                    else
                    {
                        bool flag = false;
                        XmlDocument xmlDoc = new XmlDocument();
                        xmlDoc.Load(filePath);
                        XmlNode root = xmlDoc.DocumentElement;
                        XmlNode childNode = root.FirstChild;

                        if (childNode.Name != nodeName)
                        {
                            XmlNode lastNode = root.LastChild;
                            childNode = childNode.NextSibling;
                            do
                            {
                                if (childNode.Name == nodeName)
                                {
                                    flag = true;
                                    break;
                                }
                            }
                            while (childNode != lastNode);
                        }
                        else
                        {
                            flag = true;
                        }
                        if (flag != true)
                        {
                            Console.WriteLine("Node " + nodeName + "notFound");
                            return false;
                        }
                        else
                        {
                            //Create a new node.
                            XmlElement element = xmlDoc.CreateElement(attributeName);
                            element.InnerText = attributeValue;

                            //Add the node to the document.  
                            childNode.AppendChild(element);
                            xmlDoc.Save(filePath);
                            return true;
                        }
                    }
                }
                else
                {
                    System.IO.Directory.CreateDirectory(fileFolderPath);
                    // Create the xml document containe
                    XmlDocument doc = new XmlDocument();// Create the XML Declaration, and append it to XML document
                    XmlDeclaration declaration = doc.CreateXmlDeclaration("1.0", null, null);
                    doc.AppendChild(declaration);// Create the root element
                    XmlElement root = doc.CreateElement("RunTimeValues");
                    doc.AppendChild(root);

                    // Create Nodes For RunTimeValues            
                    XmlElement reportFolder = doc.CreateElement(nodeName);
                    XmlElement folderPath = doc.CreateElement(attributeName);
                    folderPath.InnerText = attributeValue;
                    reportFolder.AppendChild(folderPath);
                    root.AppendChild(reportFolder);
                    doc.Save(filePath);
                    return true;
                }
            }
            catch (Exception exception)
            {
                // write log
                Console.WriteLine(exception.Message);
                return false;
            }
        }
        #endregion

        #region Function to set values in XML file
        /// <summary>
        /// Function to set values in XML file
        /// </summary>
        /// <returns>script name</returns>
        /// <Date>02-Nov-2011</Date>
        public bool SetValuesToXML(string nodeName, string attributeName, string attributeValue, string fileFolderPath)
        {
            try
            {
                string value = null;
                //string fileFolderPath = PrepareConfigureDataFilePath();
                string filePath = fileFolderPath + "\\" + "RunTime.xml";
                //Console.WriteLine("filePath " + filePath);
                bool isFileExists = File.Exists(filePath);
                if (isFileExists)
                {
                    System.IO.File.SetAttributes(filePath, System.IO.FileAttributes.Normal);
                    value = GetValuesFromXML(nodeName, attributeName, filePath);
                    if (value == attributeValue)
                    {
                        return true;
                    }
                    else if (value != null && value != attributeValue)
                    {
                        bool flag = false;
                        XmlDocument XMLEdit = new XmlDocument();
                        XMLEdit.Load(filePath);

                        XmlElement XMLEditNode = XMLEdit.DocumentElement;

                        foreach (XmlNode node in XMLEditNode) // Loop through XML file
                        {
                            if (node.Name == nodeName) // Check for the nodeName field information
                            {
                                XmlNode childNode = node.FirstChild;
                                XmlNode lastNode = node.LastChild;
                                if (childNode.Name == attributeName)
                                {
                                    childNode.InnerText = attributeValue;
                                    flag = true;
                                    break;
                                }
                                else
                                {
                                    do
                                    {
                                        childNode = childNode.NextSibling;
                                        if (childNode.Name == attributeName)
                                        {
                                            childNode.InnerText = attributeValue;
                                            flag = true;
                                            break;
                                        }
                                    }
                                    while (!childNode.Equals(lastNode));
                                }
                            }
                        }
                        if (flag == true)
                        {
                            XMLEdit.Save(filePath);
                            return true;
                        }
                        else
                        {
                            return false;
                        }
                    }
                    else
                    {
                        bool flag = false;
                        XmlDocument xmlDoc = new XmlDocument();
                        xmlDoc.Load(filePath);
                        XmlNode root = xmlDoc.DocumentElement;
                        XmlNode childNode = root.FirstChild;

                        if (childNode.Name != nodeName)
                        {
                            XmlNode lastNode = root.LastChild;
                            childNode = childNode.NextSibling;
                            do
                            {
                                if (childNode.Name == nodeName)
                                {
                                    flag = true;
                                    break;
                                }
                            }
                            while (childNode != lastNode);
                        }
                        else
                        {
                            flag = true;
                        }
                        if (flag != true)
                        {
                            Console.WriteLine("Node " + nodeName + "notFound");
                            return false;
                        }
                        else
                        {
                            //Create a new node.
                            XmlElement element = xmlDoc.CreateElement(attributeName);
                            element.InnerText = attributeValue;

                            //Add the node to the document.  
                            childNode.AppendChild(element);
                            xmlDoc.Save(filePath);
                            return true;
                        }
                    }
                }
                else
                {
                    System.IO.Directory.CreateDirectory(fileFolderPath);
                    // Create the xml document containe
                    XmlDocument doc = new XmlDocument();// Create the XML Declaration, and append it to XML document
                    XmlDeclaration declaration = doc.CreateXmlDeclaration("1.0", null, null);
                    doc.AppendChild(declaration);// Create the root element
                    XmlElement root = doc.CreateElement("RunTimeValues");
                    doc.AppendChild(root);

                    // Create Nodes For RunTimeValues            
                    XmlElement reportFolder = doc.CreateElement(nodeName);
                    XmlElement folderPath = doc.CreateElement(attributeName);
                    folderPath.InnerText = attributeValue;
                    reportFolder.AppendChild(folderPath);
                    root.AppendChild(reportFolder);
                    doc.Save(filePath);
                    return true;
                }
            }
            catch (Exception exception)
            {
                // write log
                Console.WriteLine(exception.Message);
                return false;
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
                if (null == filePath)
                {
                    xmlFilePath = this.currentDirectory + this.projectDirectory + "\\" + this.testDataDirectory + "\\" + this.testScriptName + "\\" + this.testScriptName + ".xml";
                }
                else
                {
                    xmlFilePath = filePath;
                }
                //Console.WriteLine("xmlFilePath " + xmlFilePath);

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
            catch (Exception)
            {
                return null;
            }

        }
        #endregion

        #region DeleteConfigDetailsFile
        /// <summary>
        /// 
        /// </summary>
        public void DeleteConfigDetailsFile()
        {
            try
            {
                string fileFolderPath = PrepareConfigureDataFilePath();

                string fileName = fileFolderPath + "\\" + "RunTime.xml";

                // check if log file exist
                if (File.Exists(fileName))
                {
                    try
                    {
                        System.IO.File.SetAttributes(fileName, System.IO.FileAttributes.Normal);
                        // delete log file
                        File.Delete(fileName);
                    }
                    catch (Exception exception)
                    {
                        Console.WriteLine("Failed to delete log file => " + exception.Message);
                    }
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine("Failed to delete log file => " + exception.Message);
            }
            finally
            {
            }
        }
        #endregion

        #region function for kill all open IE browsers
        /// <summary>
        /// Function to Kill All Already Open Internet Explorer instances
        /// </summary>
        /// <returns>string With pass value or Exception</returns>
        public void KillAlreadyOpenBrowsers()
        {
            try
            {
                // get ie process names to array
                Process[] localByName = Process.GetProcessesByName("iexplore");

                foreach (Process item in localByName)
                {
                    // kill the process
                    item.Kill();
                }
            }
            catch (Exception)
            {
            }
        }
        #endregion
        
        #region Function for write log, kill all open IE browsers, close the browser object gracefully and return the exception
        /// <summary>
        /// Function for write log, kill all open IE browsers, close the browser object gracefully and return the exception
        /// </summary>
        /// <returns>Exception string</returns>
        public void ExceptionCleanUp(string testScriptName, int stepNo, string exception, Browser browser)
        {
            try
            {
                // write log for failed step and to take screen-shot
                WriteLogs(testScriptName, stepNo, "Execution of script terminated. " + exception, "FAIL", browser);

                // Logout script
                //Automation.Development.Pages.SearchPage searchPage = new Automation.Development.Pages.SearchPage(browser);
                //searchPage.ReleaseAllKey();
                //searchPage.LogOut();
                
                // close the browser
                browser.Close();
                // kill all open IE browsers
                KillAlreadyOpenBrowsers();
                // set browser object to null
                browser = null;
            }
            catch (Exception)
            {
                // close the browser
                browser.Close();
                // kill all open IE browsers
                KillAlreadyOpenBrowsers();
                // set browser object to null
                browser = null;
            }
            // throw exception forcefully to fail the test case 
            Assert.Fail();
        }
        #endregion

        #region Function for kill all open IE browsers, close the browser object gracefully and return the exception
        /// <summary>
        /// Function for kill all open IE browsers, close the browser object gracefully and return the exception
        /// </summary>
        /// <returns>Exception string</returns>
        public string ExceptionCleanUp(string errorMessage, Browser browser)
        {
            try
            {
                //Automation.Development.Pages.SearchPage searchPage = new Automation.Development.Pages.SearchPage(browser);
                //searchPage.LogOut();
                //// close the browser
                //browser.Close();
                //// kill all open IE browsers
                //KillAlreadyOpenBrowsers();
                // set browser object to null
                browser = null;
            }
            catch (Exception)
            {
                // close the browser
                browser.Close();
                // kill all open IE browsers
                KillAlreadyOpenBrowsers();
                // set browser object to null
                browser = null;
            }
            // throw exception forcefully to fail the test case 
            Console.WriteLine("Execution of clean up script terminated. " + errorMessage);
            return null;
        }
        #endregion

        #region Function to return arraylist of aborted,not executed and error test cases
        /// <summary>
        ///  Function to return arraylist of aborted,not executed and error test cases
        /// </summary>
        /// <param name="testCaseStatus">test case status</param>
        /// <returns></returns>
        public ArrayList TestCasesDetails(string testCaseStatus)
        {
            try
            {
                //decleration of array list
                ArrayList testcaseList = new ArrayList();
                //getting path for TestRunID
                string pathUptoRMDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\VSEQT\\QTController\\rm";
                //finding all subdirectories in rm folder
                string[] subDirectories = Directory.GetDirectories(pathUptoRMDir);
                int length = subDirectories.Length;
                //getting name of current directory
                string directoryName = subDirectories[length - 1];
                string[] splitDirName = directoryName.Split('\\');
                int splitLen = splitDirName.Length;
                //current test run id
                int testRunID = Int32.Parse(splitDirName[splitLen - 1]);
                //connecting project location
                TfsTeamProjectCollection tfs = new TfsTeamProjectCollection(new Uri("[3:00:12 PM] shiv kumar Yadav: http://67.40.18.251:8080/tfs"));
                //getting service for test management
                ITestManagementService testManagement = (ITestManagementService)tfs.GetService(typeof(ITestManagementService));
                //finding automation project by name
                ITestManagementTeamProject testProject = testManagement.GetTeamProject("ComputeNext");
                //getting all details of test run
                ITestRun testRun = testProject.TestRuns.Find(testRunID);
                ITestCaseResultCollection testCases;
                //getting all aborted test cases
                if (testCaseStatus.Equals("Aborted"))
                {
                    testCases = testRun.QueryResultsByOutcome(TestOutcome.Failed);
                }
                //getting all not executed test cases
                else if(testCaseStatus.Equals("NotExecuted"))
                {
                    testCases = testRun.QueryResultsByOutcome(TestOutcome.NotExecuted);
                }
                //getting all error test cases
                else if (testCaseStatus.Equals("Error"))
                {
                    testCases = testRun.QueryResultsByOutcome(TestOutcome.Error);
                }
                else
                {
                    return null;
                }
                //adding test cases in array list
                foreach(ITestCaseResult testCaseID in testCases)
                {
                    testcaseList.Add(testCaseID.TestCaseId);        
                }
                return testcaseList;
            }
            catch(Exception)
            {
                return null;
            }
        }
        #endregion
    }
}


