using System;
using System.Collections;
using System.IO;
using System.Xml;
using System.Threading;


namespace Automation.TestScripts
{
    public class ApplicationLog
    {
        #region Constructeurs
        public ApplicationLog()
        {
        }

        public ApplicationLog(string dbgExeFileLocation, string reportFileDirectory, string testScriptName)
        {
            this.DbgExeFileLocation = dbgExeFileLocation;
            this.ReportFileDirectory = reportFileDirectory;
            this.TestScriptName = testScriptName;
        }
        #endregion Constructeurs

        string _sDbgExeFileLocation;
        ///
        /// Report File Directory
        ///
        public string DbgExeFileLocation
        {
            get { return _sDbgExeFileLocation; }
            set { _sDbgExeFileLocation = value; }
        }

        string _sReportFileDirectory;
        ///
        /// Report File Directory
        ///
        public string ReportFileDirectory
        {
            get { return _sReportFileDirectory; }
            set { _sReportFileDirectory = value; }
        }

        string _sTestScriptName;
        ///
        /// Report File Directory
        ///
        public string TestScriptName
        {
            get { return _sTestScriptName; }
            set { _sTestScriptName = value; }
        }

        public void StartDebugViewer()
        {
            // Prepare file path for debugview exe
            this.DbgExeFileLocation = this.DbgExeFileLocation + "\\Dbgview.exe";
            // Prepare log file path for debug view
            this.ReportFileDirectory = this.ReportFileDirectory + "\\" + this.TestScriptName;
            string logFileName = this.ReportFileDirectory + "\\" + "DebugView.log";
            //Console.WriteLine("logFileName " + logFileName);
            string strCommandParameters = "/t /l " + logFileName;
            try
            {
                // check if log file exist
                if (!Directory.Exists(this.ReportFileDirectory))
                {
                    // create directory
                    Directory.CreateDirectory(this.ReportFileDirectory);
                }
                try
                {
                    //Create process
                    System.Diagnostics.Process pProcess = new System.Diagnostics.Process();

                    //strCommand is path and file name of command to run
                    pProcess.StartInfo.FileName = this.DbgExeFileLocation;

                    //strCommandParameters are parameters to pass to program
                    pProcess.StartInfo.Arguments = strCommandParameters;

                    pProcess.StartInfo.UseShellExecute = false;

                    //Set output of program to be written to process output stream
                    pProcess.StartInfo.RedirectStandardOutput = true;

                    //Optional
                    //pProcess.StartInfo.WorkingDirectory = strWorkingDirectory;

                    //Start the process
                    pProcess.Start();

                    //Get program output
                    StreamReader strOutput = pProcess.StandardOutput;

                    //Wait for process to finish
                    pProcess.WaitForExit(2000);
                    pProcess.Close();

                }
                catch (Exception exception)
                {
                    // Print exception on console
                    Console.WriteLine("Failed to start debug viewer => " + exception.Message);
                }
            }
            catch (Exception exception)
            {
                // Print exception on console
                Console.WriteLine("Failed to start debug viewer => " + exception.Message);
            }
            finally
            {
            }
        }

        public void StopDebugViewer()
        {
            try
            {
                // get ie process names to array
                System.Diagnostics.Process[] localByName = System.Diagnostics.Process.GetProcessesByName("Dbgview");
                // Iterate through all process having name dbgview
                foreach (System.Diagnostics.Process item in localByName)
                {
                    // kill the process
                    item.Kill();
                }
                // Add wait after killing all files
                Thread.Sleep(2000);
            }
            catch (Exception exception)
            {
                // Print exception on console
                Console.WriteLine("Failed to stop debug viewer => " + exception.Message);
            }
            finally
            {
            }
        }

        public bool VerifyDebugLogFiles(string testReportFileDirectory, string testScriptName)
        {
            try
            {
                // generate log file path
                testReportFileDirectory = testReportFileDirectory + "\\" + testScriptName;
                string applicationLogFile = testReportFileDirectory + "\\" + "DebugView.log";
                // check if log file exist
                if (!File.Exists(applicationLogFile))
                {
                    // If debug viewer log file in not present, return true as there is no exception
                    return true;
                }
                else 
                {
                    StreamReader file = null;
                    string textPresentInLogFile = null;
                    // array of keywords to find in log file
                    string[] keywords = new string[3] { "exceptn", "exception", "error" };
                    try
                    {
                        file = new StreamReader(applicationLogFile);
                        // read log file line by line
                        while ((textPresentInLogFile = file.ReadLine()) != null)
                        {
                            try
                            {
                                // check all keywords in each line
                                for (int count = 0; count < keywords.Length; count++)
                                {
                                    if (textPresentInLogFile.ToString().ToLower().Contains(keywords[count]))
                                    {
                                        // return false if exception is present
                                        return false;
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                    finally
                    {
                        // close the file object
                        if (file != null)
                            file.Close();
                    }
                }                                
                return true;
            }
            catch (Exception exception)
            {
                // Print exception on console
                Console.WriteLine("Failed to verify debug file => " + exception.Message);
                return false;
            }
            finally
            {
            }
        }

        /// <summary>
        /// Function to verify application logs
        /// </summary>
        /// <param name="applicationLogDirectory"></param>
        /// <param name="logFileNames"></param>
        /// <returns></returns>
        public bool VerifyApplicationLogFiles(string applicationLogDirectory, string logFileNames)
        {
            try
            {
                string[] fileNames = logFileNames.Split(',');
                for (int outerLoop = 0; outerLoop < fileNames.Length; outerLoop++)
                {
                    // generate log file path
                    string applicationLogFile = applicationLogDirectory + "\\" + fileNames[outerLoop];

                    // check if log file exist
                    if (File.Exists(applicationLogFile))
                    {
                        StreamReader file = null;
                        string textPresentInLogFile = null;
                        string[] keywords = new string[3] { "exceptn", "exception", "error" };
                        try
                        {
                            file = new StreamReader(applicationLogFile);
                            // read log file line by line
                            while ((textPresentInLogFile = file.ReadLine()) != null)
                            {
                                // check all keywords in each line
                                for (int count = 0; count < keywords.Length; count++)
                                {
                                    if (textPresentInLogFile.ToLower().Contains(keywords[count]))
                                    {
                                        Console.WriteLine(textPresentInLogFile);
                                        return false;
                                    }
                                }
                            }
                        }
                        finally
                        {
                            // close file object
                            if (file != null)
                                file.Close();
                        }
                    }
                }
                return true;
            }
            catch (Exception exception)
            {
                // Print exception on console
                Console.WriteLine("Failed to verify log file => " + exception.Message);
                return false;
            }
            finally
            {
            }
        }
                
        #region depricated function
        public bool VerifyApplicationLogFiles_OLD(string applicationLogDirectory, string logFileNames)
        {
            try
            {
                string[] fileNames = logFileNames.Split(',');
                for (int outerLoop = 0; outerLoop < fileNames.Length; outerLoop++)
                {
                    // generate log file path
                    string applicationLogFile = applicationLogDirectory + "\\" + fileNames[outerLoop];

                    // check if log file exist
                    if (!File.Exists(applicationLogFile))
                    {
                        break;
                    }

                    // read file contents in a veriable
                    string textPresentInLogFile = System.IO.File.ReadAllText(applicationLogFile);

                    // Verify if the keywords present in file contents.
                    string[] keywords = new string[3] { "exceptn", "exception", "error" };
                    for (int count = 0; count < keywords.Length; count++)
                    {
                        if (textPresentInLogFile.ToLower().Contains(keywords[count]))
                        {
                            return false;
                        }
                    }
                    // empty the string
                    textPresentInLogFile = null;
                }
                return true;
            }
            catch (Exception exception)
            {
                Console.WriteLine("Failed to verify log file => " + exception.Message);
                return false;
            }
            finally
            {
            }
        }
        #endregion

        public void CopyApplicationLogFiles(string applicationLogDirectory, string logFileNames, string testReportFileDirectory, string testScriptName)
        {
            try
            {
                testReportFileDirectory = testReportFileDirectory + "\\" + testScriptName;
                string[] fileNames = logFileNames.Split(',');
                for (int outerLoop = 0; outerLoop < fileNames.Length; outerLoop++)
                {
                    // generate log file path
                    string applicationLogFile = applicationLogDirectory + "\\" + fileNames[outerLoop];
                    string destinationFilePath = testReportFileDirectory + "\\" + fileNames[outerLoop];

                    // check if log file exist
                    if (File.Exists(applicationLogFile))
                    {
                        if (!Directory.Exists(testReportFileDirectory))
                        {
                            // create directory
                            Directory.CreateDirectory(testReportFileDirectory);
                        }
                        try
                        {
                            // delete log file
                            File.Copy(applicationLogFile, destinationFilePath);
                        }
                        catch (Exception exception)
                        {
                            // Print exception on console
                            Console.WriteLine("Failed to copy log file => " + exception.Message);
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                // Print exception on console
                Console.WriteLine("Failed to copy log file => " + exception.Message);
            }
            finally
            {
            }
        }

        public void DeleteApplicationLogFiles(string applicationLogDirectory, string logFileNames)
        {
            try
            {
                string[] fileNames = logFileNames.Split(',');
                for (int outerLoop = 0; outerLoop < fileNames.Length; outerLoop++)
                {
                    // generate log file path
                    string applicationLogFile = applicationLogDirectory + "\\" + fileNames[outerLoop];

                    // check if log file exist
                    if (File.Exists(applicationLogFile))
                    {
                        try
                        {
                            // delete log file
                            File.Delete(applicationLogFile);
                        }
                        catch (Exception exception)
                        {
                            Console.WriteLine("Failed to delete log file => " + exception.Message);
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                // Print exception on console
                Console.WriteLine("Failed to delete log file => " + exception.Message);
            }
            finally
            {
            }
        }
    }
}


