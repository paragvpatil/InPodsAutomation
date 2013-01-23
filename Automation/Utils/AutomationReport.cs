using System;
using System.Collections;
using System.Data.OleDb;
using System.IO;


namespace Automation.TestScripts
{
    public class AutomationReport
    {
        #region Constructeurs
        public AutomationReport()
        {
        }

        public AutomationReport(string logFileDirectory, string reportFileDirectory, string serverName, DateTime startTimeOfExecution, DateTime endTimeOfExecution)
        {
            this.LogFileDirectory = logFileDirectory;
            this.ReportFileDirectory = reportFileDirectory;
            this.ServerName = serverName;
            this.StartTimeOfExecution = startTimeOfExecution;
            this.EndTimeOfExecution = endTimeOfExecution;
        }
        #endregion Constructeurs

        DateTime _sStartTimeOfExecution;
        ///
        /// Start Time Of Execution
        ///
        public DateTime StartTimeOfExecution
        {
            get { return _sStartTimeOfExecution; }
            set { _sStartTimeOfExecution = value; }
        }
        DateTime _sEndTimeOfExecution;
        ///
        /// End Time Of Execution
        ///
        public DateTime EndTimeOfExecution
        {
            get { return _sEndTimeOfExecution; }
            set { _sEndTimeOfExecution = value; }
        }

        string _sLogFileDirectory;
        ///
        /// Log File Directory
        ///
        public string LogFileDirectory
        {
            get { return _sLogFileDirectory; }
            set { _sLogFileDirectory = value; }
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

        string _sServerName;
        ///
        /// Server Name
        ///
        public string ServerName
        {
            get { return _sServerName; }
            set { _sServerName = value; }
        }


        #region Function to write individual test case .htm file
        /// <summary>
        /// Function to write individual test case .htm file
        /// </summary>
        /// <param name="testCaseId">Test case Id</param>
        /// <param name="testData">Test data in htm format</param>
        /// <author>rajan.bansod</author> 
        /// <ModifiedBy>rajan.bansod</ModifiedBy>
        /// <Date>12-Aug-2011</Date>
        private void CreateTestCaseFile(string testCaseId, string testData, string htmlHeaderData, string testRunPath)
        {
            StreamWriter testcaseWriter = null;
            try
            {
                testRunPath = testRunPath + @"\" + testCaseId;
                // check if directory exist or not
                if (!Directory.Exists(testRunPath))
                {
                    // if directory is not exist, create one
                    Directory.CreateDirectory(testRunPath);
                }
                
                // Prepare main html file path
                FileInfo fileInfo = new FileInfo(testRunPath + @"\" + testCaseId + ".htm");
                // Append data to html log file
                string sHTML = "<HTML><frameset rows=\"12%,*\" framespacing=\"0\" frameborder=\"no\" border=\"0\">";
                sHTML = sHTML + "<frame src=\"" + testCaseId + "_Header.htm\"/>";
                sHTML = sHTML + "<frameset framespacing=\"0\" frameborder=\"no\" border=\"0\">";
                sHTML = sHTML + "<frame src=\"" + testCaseId + "_Details.htm\" name=\"targetframe\"/>";
                sHTML = sHTML + "</frameset></frameset></HTML>";
                testcaseWriter = fileInfo.AppendText();
                testcaseWriter.WriteLine(sHTML);
                testcaseWriter.Flush();

                // Prepare header file path
                FileInfo fileInfoHeader = new FileInfo(testRunPath + @"\" + testCaseId + "_Header.htm");
                // Append data to html log file
                testcaseWriter = fileInfoHeader.AppendText();
                testcaseWriter.WriteLine(htmlHeaderData);
                testcaseWriter.Flush();

                // Prepare detailed steps file path
                FileInfo fileInfoDetails = new FileInfo(testRunPath + @"\" + testCaseId + "_Details.htm");
                // Append data to html log file
                testcaseWriter = fileInfoDetails.AppendText();
                testcaseWriter.WriteLine(testData);
                testcaseWriter.Flush();
            }
            catch (Exception exception)
            {
                // Print exception on console
                Console.WriteLine("Test case file writer. Error in writing test case files " + exception.Message);
            }
            finally
            {
                testcaseWriter.Close();
            }
        }
        #endregion

        #region Function to create a header report
        /// <summary>
        /// Function to create a header report
        /// </summary>
        /// <param name="totalTestCases">Total number of test cases</param>
        /// <param name="PassedTestCases">Total Passed test cases</param>
        /// <param name="FailedTestCases">Total failed test cases</param>
        /// <param name="totalSteps">Total steps</param>
        /// <param name="FailedSteps">Total failed steps</param>
        /// <param name="PassedSteps">Total Passed steps</param>
        /// <author>rajan.bansod</author> 
        /// <ModifiedBy>rajan.bansod</ModifiedBy>
        /// <Date>12-Aug-2011</Date>
        private void CreateHeaderReport(int totalTestCases, int passedTestCases, int failedTestCases, int infoTestCases, int totalSteps, int failedSteps, int passedSteps, int infoSteps, string testRunPath, string serverName)
        {
            StreamWriter headerReportWriter = null;
            TimeSpan timeSpan = new TimeSpan();
            try
            {
                // find out the differance
                timeSpan = this.EndTimeOfExecution.Subtract(this.StartTimeOfExecution);
            }
            catch (Exception) { }

            try
            {
                testRunPath = testRunPath + @"\HtmlFiles";
                if (!Directory.Exists(testRunPath))
                {
                    Directory.CreateDirectory(testRunPath);
                }
                FileInfo fileInfo = new FileInfo(testRunPath + @"\ReportHeader.htm");
                headerReportWriter = fileInfo.AppendText();
                headerReportWriter.WriteLine("<HTML><BODY bgcolor=\"#C5D8FF\" STYLE=\"background-repeat: no-repeat;\"><TABLE width=\"100%\">");
                string report = "<TR><TD>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp";
                report = report + "</BR></BR> &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp <A href=\"dashboard.htm\" style=\"text-decoration:none\" target=\"targetframe\" >DASHBOARD<A></TD>";
                report = report + "<TD><A href=\"TestResultReport.htm\" style=\"text-decoration:none\" target=\"resultframe\"><font size=\"2\" face=\"Calibri\" color=black> Test cases executed:&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp" + totalTestCases + "</A></BR>";
                report = report + "<A href=\"TestPassedResultReport.htm\" style=\"text-decoration:none\" target=\"resultframe\"><font size=\"2\" face=\"Calibri\" color=green>Total passed test cases:&nbsp&nbsp" + passedTestCases + "</font></A></BR>";
                report = report + "<A href=\"TestFailedResultReport.htm\" style=\"text-decoration:none\" target=\"resultframe\"><font size=\"2\" face=\"Calibri\" color=red>Total failed test cases:&nbsp&nbsp&nbsp&nbsp" + failedTestCases + "</font></A></BR>";
                report = report + "<font size=\"2\" face=\"Calibri\" color=blue>Total info test cases:&nbsp&nbsp&nbsp&nbsp" + infoTestCases + "</font></BR>";
                report = report + "</TD><TD><font size=\"2\" face=\"Calibri\"> Total Steps : " + totalSteps + "</font></BR>";
                report = report + "<font size=\"2\" face=\"Calibri\">Total passed steps : " + passedSteps + "</font></BR>";
                report = report + "<font size=\"2\" face=\"Calibri\">Total failed steps : " + failedSteps + "</font></BR>";
                report = report + "<font size=\"2\" face=\"Calibri\">Total info steps : " + infoSteps + "</font></BR>";
                report = report + "</TD><TD><font size=\"2\" face=\"Calibri\" color=green><STRONG>Server Name  :  <a href=\"https://" + serverName + "/ComputeNext/\"  style=\"text-decoration:none\" target=\"_blank\">" + serverName + "</a></STRONG></font></BR>";
                report = report + "<font size=\"2\" face=\"Calibri\">Execution start time: " + this.StartTimeOfExecution.ToString() + "</font></BR>";
                report = report + "<font size=\"2\" face=\"Calibri\">Execution  end  time : " + this.EndTimeOfExecution.ToString() + "</font></BR>";
                if (0 == timeSpan.Days)
                {
                    report = report + "<font size=\"2\" face=\"Calibri\">Total execution time: " + timeSpan.Hours + " Hrs. " + timeSpan.Minutes + " Mins. " + timeSpan.Seconds + " Secs. " + "</font></BR>";
                }
                else
                {
                    report = report + "<font size=\"2\" face=\"Calibri\">Total execution time: " + timeSpan.Days + " Days. " + timeSpan.Hours + " Hrs. " + timeSpan.Minutes + " Mins. " + timeSpan.Seconds + " Secs. " + "</font></BR>";
                }
                report = report + "</TD></TR>";
                headerReportWriter.WriteLine(report);
                headerReportWriter.WriteLine("</TABLE></BODY></HTML>");
            }
            catch (Exception exception)
            {
                // Print exception on console
                Console.WriteLine("CreateHeaderReport " + "Writing Header report. Error in writing header report " + exception.Message);
            }
            finally
            {
                headerReportWriter.Close();
                GC.Collect();
            }
        }
        #endregion

        #region Function to create a Embedded Mail Contents
        /// <summary>
        /// Function to create a header report
        /// </summary>
        /// <param name="totalTestCases">Total number of test cases</param>
        /// <param name="PassedTestCases">Total Passed test cases</param>
        /// <param name="FailedTestCases">Total failed test cases</param>
        /// <param name="totalSteps">Total steps</param>
        /// <param name="FailedSteps">Total failed steps</param>
        /// <param name="PassedSteps">Total Passed steps</param>
        /// <author>rajan.bansod</author> 
        /// <ModifiedBy>rajan.bansod</ModifiedBy>
        /// <Date>12-Aug-2011</Date>
        private string CreateEmbeddedMailContents(int totalTestCases, int passedTestCases, int failedTestCases, int infoTestCases, string serverName, string testRunPath)
        {
            TimeSpan timeSpan = new TimeSpan();
            string embeddedMailContents = null;
            try
            {
                // find out the differance
                timeSpan = this.EndTimeOfExecution.Subtract(this.StartTimeOfExecution);
            }
            catch (Exception) { }

            try
            {
                embeddedMailContents = "<TABLE width=\"100%\" border=\"1\"><TR>";
                embeddedMailContents = embeddedMailContents + "<TD><font size=\"3\" face=\"Calibri\" color=black> Test cases executed: " + totalTestCases + "</A><br />";
                embeddedMailContents = embeddedMailContents + "<font size=\"3\" face=\"Calibri\" color=green>Total passed test cases: " + passedTestCases + "</font></A><br />";
                embeddedMailContents = embeddedMailContents + "<font size=\"3\" face=\"Calibri\" color=red>Total failed test cases: " + failedTestCases + "</font></A><br />";
                embeddedMailContents = embeddedMailContents + "<font size=\"3\" face=\"Calibri\" color=blue>Total info test cases: " + infoTestCases + "</font><br />";
                embeddedMailContents = embeddedMailContents + "</TD><TD><font size=\"3\" face=\"Calibri\" color=green><STRONG>Server Name  :  <a href=\"https://" + serverName + "/ComputeNext/\"  style=\"text-decoration:none\" target=\"_blank\">" + serverName + "</a></STRONG></font><br />";
                embeddedMailContents = embeddedMailContents + "<font size=\"3\" face=\"Calibri\">Execution start time: " + this.StartTimeOfExecution.ToString() + "</font><br />";
                embeddedMailContents = embeddedMailContents + "<font size=\"3\" face=\"Calibri\">Execution  end  time : " + this.EndTimeOfExecution.ToString() + "</font><br />";
                if (0 == timeSpan.Days)
                {
                    embeddedMailContents = embeddedMailContents + "<font size=\"3\" face=\"Calibri\">Total execution time: " + timeSpan.Hours + " Hrs. " + timeSpan.Minutes + " Mins. " + timeSpan.Seconds + " Secs. " + "</font><br />";
                }
                else
                {
                    embeddedMailContents = embeddedMailContents + "<font size=\"3\" face=\"Calibri\">Total execution time: " + timeSpan.Days + " Days. " + timeSpan.Hours + " Hrs. " + timeSpan.Minutes + " Mins. " + timeSpan.Seconds + " Secs. " + "</font><br />";
                }
                embeddedMailContents = embeddedMailContents + "</TD></TR></TABLE>";

                embeddedMailContents = embeddedMailContents + "<br /><font size=\"3\" face=\"Calibri\" color=black> <strong> Please find the detailed execution status on server location: </font><font size=\"3\" face=\"Calibri\" color=green>" + serverName + " </strong><br />";
                return embeddedMailContents = embeddedMailContents + "<font size=\"3\" face=\"Calibri\" color=black>" + testRunPath + @"\FinalTestReport.htm</font>";

                //embeddedMailContents = "<head><script type=\"text/javascript\" src=\"https://www.google.com/jsapi\"></script><script type=\"text/javascript\">google.load(\"visualization\", \"1\", {packages:[\"corechart\"]});google.setOnLoadCallback(drawChart);function drawChart() {var data = google.visualization.arrayToDataTable([['Test Status', 'Results'],['Passed', " + passedTestCases + "],  ['Failed'," + failedTestCases + "]]);";
                //embeddedMailContents = embeddedMailContents + "var options = {title: 'ComputeNext Test Automation Summery',colors:['#2E8B57','#DC143C']};var chart = new google.visualization.PieChart(document.  getElementById('chart_div'));chart.draw(data, options);}</script></head>";
                //embeddedMailContents = "<body><div id=\"chart_div\" style=\"width: 900px; height: 500px;\"></div></body>";

            }
            catch (Exception exception)
            {
                // Print exception on console
                Console.WriteLine("EmbeddedMailContents. " + "Error in writing Embedded Mail Contents " + exception.Message);
                return "Unable to add contents to the mail. Please check the logs for more details.";
            }
        }
        #endregion

        #region Function to write report for test case results in html format
        /// <summary>
        /// Function to write report for test case results in html format
        /// </summary>
        /// <param name="testCaseList">Array List of test case Ids</param>
        /// <author>rajan.bansod</author> 
        /// <ModifiedBy>rajan.bansod</ModifiedBy>
        /// <Date>12-Aug-2011</Date>
        private void CreateTestResultReport(ArrayList testCaseList, string testRunPath)
        {
            StreamWriter testResultWriter = null;
            try
            {
                testRunPath = testRunPath + @"\HtmlFiles";
                if (!Directory.Exists(testRunPath))
                {
                    Directory.CreateDirectory(testRunPath);
                }

                FileInfo fileInfo = new FileInfo(testRunPath + @"\TestResultReport.htm");
                testResultWriter = fileInfo.AppendText();
                testResultWriter.WriteLine("<HTML><BODY background=\"../../../imgscr/MainBody.png style=\"background-repeat: no-repeat;\"><TABLE width=\"100%\">");

                string report = "";
                for (int count = 0; count < testCaseList.Count; count++)
                {
                    string[] testCase = testCaseList[count].ToString().Split(',');
                    if (testCase[1] == "PASS")
                    {
                        report = report + "<TR><TD><button style=\"background-color:#4E9258\" onClick=\"parent.targetframe.location='../" + testCase[0] + "\\\\DebugView.log'\" onmouseover=\"this.style.background='#4AA02C'\" onMouseout=\"this.style.background='#4E9258'\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\"><STRONG>!</STRONG></button></TD><TD><button style=\"background-color:#4E9258\" onClick=\"parent.targetframe.location='../" + testCase[0] + "\\\\" + testCase[0] + ".htm'\" onmouseover=\"this.style.background='#4AA02C'\" onMouseout=\"this.style.background='#4E9258'\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\"><STRONG>" + testCase[0] + "</STRONG></button>"
                        //report = report + "<TR><TD width=\"10%\"><A href=\"../" + testCase[0] + "\\DebugView.log\" target=\"targetframe\"><img src=\"../../../imgscr/PassedIcon.png\"></A></TD><TD width=\"50%\" background=\"../../../imgscr/Pass.png\" style=\"background-repeat: no-repeat; padding:  2px 5px 2px 13px\"><A href=\"../" + testCase[0] + "\\" + testCase[0] + ".htm\" style=\"text-decoration:none\" target=\"targetframe\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\">" + testCase[0]
                                        //+ "</TD><TD width=\"50%\" background=\"../../../imgscr/Result.png\" style=\"background-repeat: no-repeat; padding:  2px 5px 2px 20px\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\">" + testCase[1]
                                        + "</TD></TR>";
                    }
                    else if (testCase[1] == "FAIL")
                    {
                        report = report + "<TR><TD><button style=\"background-color:#4E9258\" onClick=\"parent.targetframe.location='../" + testCase[0] + "\\\\DebugView.log'\" onmouseover=\"this.style.background='#4AA02C'\" onMouseout=\"this.style.background='#4E9258'\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\"><STRONG>!</STRONG></button></TD><TD><button style=\"background-color:#C24641\" onClick=\"parent.targetframe.location='../" + testCase[0] + "\\\\" + testCase[0] + ".htm'\" onmouseover=\"this.style.background='#E77471'\" onMouseout=\"this.style.background='#C24641'\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\"><STRONG>" + testCase[0] + "</STRONG></button>"
                        //report = report + "<TR><TD width=\"10%\"><A href=\"../" + testCase[0] + "\\DebugView.log\" target=\"targetframe\"><img src=\"../../../imgscr/PassedIcon.png\"></A></TD><TD width=\"50%\" background=\"../../../imgscr/Fail.png\" style=\"background-repeat: no-repeat; padding:  2px 5px 2px 13px\"><A href=\"../" + testCase[0] + "\\" + testCase[0] + ".htm\" style=\"text-decoration:none\" target=\"targetframe\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\">" + testCase[0]
                                        //+ "</TD><TD width=\"50%\" background=\"../../../imgscr/Result.png\" style=\"background-repeat: no-repeat; padding:  2px 5px 2px 20px\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\">FAIL"
                                        + "</TD></TR>";
                    }
                    else if (testCase[1] == "INFO_PASS")
                    {
                        report = report + "<TR><TD><button style=\"background-color:#C24641\" onClick=\"parent.targetframe.location='../" + testCase[0] + "\\\\DebugView.log'\" onmouseover=\"this.style.background='#E77471'\" onMouseout=\"this.style.background='#C24641'\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\"><STRONG>!</STRONG></button></TD><TD><button style=\"background-color:#4E9258\" onClick=\"parent.targetframe.location='../" + testCase[0] + "\\\\" + testCase[0] + ".htm'\" onmouseover=\"this.style.background='#4AA02C'\" onMouseout=\"this.style.background='#4E9258'\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\"><STRONG>" + testCase[0] + "</STRONG></button>"
                        //report = report + "<TR><TD width=\"10%\"><A href=\"../" + testCase[0] + "\\DebugView.log\" target=\"targetframe\"><img src=\"../../../imgscr/FailedIcon.png\"></A></TD><TD width=\"50%\" background=\"../../../imgscr/Pass.png\" style=\"background-repeat: no-repeat; padding:  2px 5px 2px 13px\"><A href=\"../" + testCase[0] + "\\" + testCase[0] + ".htm\" style=\"text-decoration:none\" target=\"targetframe\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\">" + testCase[0]
                                        //+ "</TD><TD width=\"50%\" background=\"../../../imgscr/Result.png\" style=\"background-repeat: no-repeat; padding:  2px 5px 2px 20px\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\">PASS"
                                        + "</TD></TR>";
                    }
                    else
                    {
                        report = report + "<TR><TD><button style=\"background-color:#C24641\" onClick=\"parent.targetframe.location='../" + testCase[0] + "\\\\DebugView.log'\" onmouseover=\"this.style.background='#E77471'\" onMouseout=\"this.style.background='#C24641'\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\"><STRONG>!</STRONG></button></TD><TD><button style=\"background-color:#C24641\" onClick=\"parent.targetframe.location='../" + testCase[0] + "\\\\" + testCase[0] + ".htm'\" onmouseover=\"this.style.background='#E77471'\" onMouseout=\"this.style.background='#C24641'\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\"><STRONG>" + testCase[0] + "</STRONG></button>"
                        //report = report + "<TR><TD width=\"10%\"><A href=\"../" + testCase[0] + "\\DebugView.log\" target=\"targetframe\"><img src=\"../../../imgscr/FailedIcon.png\"></A></TD><TD width=\"50%\" background=\"../../../imgscr/Fail.png\" style=\"background-repeat: no-repeat; padding:  2px 5px 2px 13px\"><A href=\"../" + testCase[0] + "\\" + testCase[0] + ".htm\" style=\"text-decoration:none\" target=\"targetframe\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\">" + testCase[0]
                                        //+ "</TD><TD width=\"50%\" background=\"../../../imgscr/Result.png\" style=\"background-repeat: no-repeat; padding:  2px 5px 2px 20px\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\">FAIL"
                                        + "</TD></TR>";
                    }
                }
                testResultWriter.WriteLine(report);
                testResultWriter.WriteLine("</TABLE></BODY></HTML>");
            }
            catch (Exception exception)
            {
                // Print exception on console
                Console.WriteLine("CreateTestResultReport " + "Test Result Writer. Error in writing test result" + exception.Message);
            }
            finally
            {
                testResultWriter.Close();
            }
        }
        #endregion

        #region Function to write report for test case results in html format
        /// <summary>
        /// Function to write report for test case results in html format
        /// </summary>
        /// <param name="testCaseList">Array List of test case Ids</param>
        /// <author>rajan.bansod</author> 
        /// <ModifiedBy>rajan.bansod</ModifiedBy>
        /// <Date>12-Aug-2011</Date>
        private void CreatePassedTestResultReport(ArrayList testCaseList, string testRunPath)
        {
            StreamWriter testResultWriter = null;
            try
            {
                testRunPath = testRunPath + @"\HtmlFiles";
                if (!Directory.Exists(testRunPath))
                {
                    Directory.CreateDirectory(testRunPath);
                }

                FileInfo fileInfo = new FileInfo(testRunPath + @"\TestPassedResultReport.htm");
                testResultWriter = fileInfo.AppendText();
                testResultWriter.WriteLine("<HTML><BODY background=\"../../../imgscr/MainBody.png style=\"background-repeat: no-repeat;\"><TABLE width=\"100%\">");

                string report = "";
                for (int count = 0; count < testCaseList.Count; count++)
                {
                    string[] testCase = testCaseList[count].ToString().Split(',');
                    if (testCase[1] == "PASS")
                    {
                        report = report + "<TR><TD><button style=\"background-color:#4E9258\" onClick=\"parent.targetframe.location='../" + testCase[0] + "\\\\DebugView.log'\" onmouseover=\"this.style.background='#4AA02C'\" onMouseout=\"this.style.background='#4E9258'\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\"><STRONG>!</STRONG></button></TD><TD><button style=\"background-color:#4E9258\" onClick=\"parent.targetframe.location='../" + testCase[0] + "\\\\" + testCase[0] + ".htm'\" onmouseover=\"this.style.background='#4AA02C'\" onMouseout=\"this.style.background='#4E9258'\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\"><STRONG>" + testCase[0] + "</STRONG></button>"
                        //report = report + "<TR><TD><A href=\"" + testCase[0] + "\\" + testCase[0] + ".htm\" target=\"targetframe\">" + testCase[0] + "</TD><TD>" + testCase[1] + "</TD></TR>";
                        //report = report + "<TR><TD width=\"10%\"><A href=\"../" + testCase[0] + "\\DebugView.log\" target=\"targetframe\"><img src=\"../../../imgscr/PassedIcon.png\"></A></TD><TD width=\"50%\" background=\"../../../imgscr/Pass.png\" style=\"background-repeat: no-repeat; padding:  2px 5px 2px 13px\"><A href=\"../" + testCase[0] + "\\" + testCase[0] + ".htm\" style=\"text-decoration:none\" target=\"targetframe\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\">" + testCase[0]
                                        //+ "</TD><TD width=\"50%\" background=\"../../../imgscr/Result.png\" style=\"background-repeat: no-repeat; padding:  2px 5px 2px 20px\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\">" + testCase[1]
                                        + "</TD></TR>";
                    }
                    else if (testCase[1] == "INFO_PASS")
                    {
                        report = report + "<TR><TD><button style=\"background-color:#C24641\" onClick=\"parent.targetframe.location='../" + testCase[0] + "\\\\DebugView.log'\" onmouseover=\"this.style.background='#E77471'\" onMouseout=\"this.style.background='#C24641'\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\"><STRONG>!</STRONG></button></TD><TD><button style=\"background-color:#4E9258\" onClick=\"parent.targetframe.location='../" + testCase[0] + "\\\\" + testCase[0] + ".htm'\" onmouseover=\"this.style.background='#4AA02C'\" onMouseout=\"this.style.background='#4E9258'\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\"><STRONG>" + testCase[0] + "</STRONG></button>"
                        //report = report + "<TR><TD width=\"10%\"><A href=\"../" + testCase[0] + "\\DebugView.log\" target=\"targetframe\"><img src=\"../../../imgscr/FailedIcon.png\"></A></TD><TD width=\"50%\" background=\"../../../imgscr/Pass.png\" style=\"background-repeat: no-repeat; padding:  2px 5px 2px 13px\"><A href=\"../" + testCase[0] + "\\" + testCase[0] + ".htm\" style=\"text-decoration:none\" target=\"targetframe\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\">" + testCase[0]
                                        //+ "</TD><TD width=\"50%\" background=\"../../../imgscr/Result.png\" style=\"background-repeat: no-repeat; padding:  2px 5px 2px 20px\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\">PASS"
                                        + "</TD></TR>";
                    }
                }
                testResultWriter.WriteLine(report);
                testResultWriter.WriteLine("</TABLE></BODY></HTML>");
            }
            catch (Exception exception)
            {
                // Print exception on console
                Console.WriteLine("CreatePassedTestResultReport " + "Test Result Writer. Error in writing test result" + exception.Message);
            }
            finally
            {
                testResultWriter.Close();
            }
        }
        #endregion

        #region Function to write report for test case results in html format
        /// <summary>
        /// Function to write report for test case results in html format
        /// </summary>
        /// <param name="testCaseList">Array List of test case Ids</param>
        /// <author>rajan.bansod</author> 
        /// <ModifiedBy>rajan.bansod</ModifiedBy>
        /// <Date>12-Aug-2011</Date>
        private void CreateFailedTestResultReport(ArrayList testCaseList, string testRunPath)
        {
            StreamWriter testResultWriter = null;
            try
            {
                testRunPath = testRunPath + @"\HtmlFiles";
                if (!Directory.Exists(testRunPath))
                {
                    Directory.CreateDirectory(testRunPath);
                }

                FileInfo fileInfo = new FileInfo(testRunPath + @"\TestFailedResultReport.htm");
                testResultWriter = fileInfo.AppendText();
                testResultWriter.WriteLine("<HTML><BODY background=\"../../../imgscr/MainBody.png style=\"background-repeat: no-repeat;\"><TABLE width=\"100%\">");

                string report = "";
                for (int count = 0; count < testCaseList.Count; count++)
                {
                    string[] testCase = testCaseList[count].ToString().Split(',');
                    if (testCase[1] == "PASS")
                    {
                    }
                    else if (testCase[1] == "FAIL")
                    {
                        report = report + "<TR><TD><button style=\"background-color:#4E9258\" onClick=\"parent.targetframe.location='../" + testCase[0] + "\\\\DebugView.log'\" onmouseover=\"this.style.background='#4AA02C'\" onMouseout=\"this.style.background='#4E9258'\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\"><STRONG>!</STRONG></button></TD><TD><button style=\"background-color:#C24641\" onClick=\"parent.targetframe.location='../" + testCase[0] + "\\\\" + testCase[0] + ".htm'\" onmouseover=\"this.style.background='#E77471'\" onMouseout=\"this.style.background='#C24641'\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\"><STRONG>" + testCase[0] + "</STRONG></button>"
                        //report = report + "<TR><TD width=\"10%\"><A href=\"../" + testCase[0] + "\\DebugView.log\" target=\"targetframe\"><img src=\"../../../imgscr/PassedIcon.png\"></A></TD><TD width=\"50%\" background=\"../../../imgscr/Fail.png\" style=\"background-repeat: no-repeat; padding:  2px 5px 2px 13px\"><A href=\"../" + testCase[0] + "\\" + testCase[0] + ".htm\" style=\"text-decoration:none\" target=\"targetframe\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\">" + testCase[0]
                                        //+ "</TD><TD width=\"50%\" background=\"../../../imgscr/Result.png\" style=\"background-repeat: no-repeat; padding:  2px 5px 2px 20px\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\">FAIL"
                                        + "</TD></TR>";
                    }
                    else if (testCase[1] == "INFO_PASS")
                    {
                    }
                    else
                    {
                        report = report + "<TR><TD><button style=\"background-color:#C24641\" onClick=\"parent.targetframe.location='../" + testCase[0] + "\\\\DebugView.log'\" onmouseover=\"this.style.background='#E77471'\" onMouseout=\"this.style.background='#C24641'\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\"><STRONG>!</STRONG></button></TD><TD><button style=\"background-color:#C24641\" onClick=\"parent.targetframe.location='../" + testCase[0] + "\\\\" + testCase[0] + ".htm'\" onmouseover=\"this.style.background='#E77471'\" onMouseout=\"this.style.background='#C24641'\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\"><STRONG>" + testCase[0] + "</STRONG></button>"
                        //report = report + "<TR><TD width=\"10%\"><A href=\"../" + testCase[0] + "\\DebugView.log\" target=\"targetframe\"><img src=\"../../../imgscr/FailedIcon.png\"></A></TD><TD width=\"50%\" background=\"../../../imgscr/Fail.png\" style=\"background-repeat: no-repeat; padding:  2px 5px 2px 13px\"><A href=\"../" + testCase[0] + "\\" + testCase[0] + ".htm\" style=\"text-decoration:none\" target=\"targetframe\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\">" + testCase[0]
                                        //+ "</TD><TD width=\"50%\" background=\"../../../imgscr/Result.png\" style=\"background-repeat: no-repeat; padding:  2px 5px 2px 20px\"><font size=\"2\" face=\"Calibri\" align=\"center\" color=\"WHITE\">FAIL"
                                        + "</TD></TR>";
                    }
                }
                testResultWriter.WriteLine(report);
                testResultWriter.WriteLine("</TABLE></BODY></HTML>");
            }
            catch (Exception exception)
            {
                // Print exception on console
                Console.WriteLine("CreateFailedTestResultReport " + "Test Result Writer. Error in writing test result" + exception.Message);
            }
            finally
            {
                testResultWriter.Close();
            }
        }
        #endregion

        #region Function to generate dashboard html file
        /// <summary>
        /// Function to generate final html report
        /// </summary>
        /// <author>rajan.bansod</author> 
        /// <ModifiedBy>rajan.bansod</ModifiedBy>
        /// <Date>12-Aug-2011</Date>
        private void CreateDashboard(string testRunPath, string totalTestCases,  string passedTestCases, string failedTestCases)
        {
            StreamWriter reportWriter = null;
            try
            {
                testRunPath = testRunPath + @"\HtmlFiles";
                if (!Directory.Exists(testRunPath))
                {
                    Directory.CreateDirectory(testRunPath);
                }

                FileInfo fileInfo = new FileInfo(testRunPath + @"\Dashboard.htm");
                reportWriter = fileInfo.AppendText();

                string report = "<HTML>";

                report = report + "<HEAD><script type=\"text/javascript\" src=\"https://www.google.com/jsapi\"></script><script type=\"text/javascript\">google.load(\"visualization\", \"1\", {packages:[\"corechart\"]});google.setOnLoadCallback(drawChart);function drawChart() {var data = google.visualization.arrayToDataTable([['Test Status', 'Results'],['Passed', " + passedTestCases + "],  ['Failed'," + failedTestCases + "]]);";
                report = report + "var options = {title: 'ComputeNext Test Automation Summery',colors:['#2E8B57','#DC143C'],chartArea:{left:50,top:50,width:\"60%\",height:\"75%\"}};var chart = new google.visualization.PieChart(document.  getElementById('chart_div'));";
                report = report + "function open_win(url){window.open(url,'resultframe');}";
                report = report + "function selectHandler() {var selectedItem = chart.getSelection()[0];if (selectedItem) {var testResult = data.getValue(selectedItem.row, 0);	if (testResult == 'Passed')	{open_win('TestPassedResultReport.htm');}else if (testResult == 'Failed'){open_win('TestFailedResultReport.htm');}else{	open_win('TestResultReport.htm');}  } } google.visualization.events.addListener(chart, 'select', selectHandler);";
                report = report + "chart.draw(data, options);}</script></HEAD>";
                
                //report = report + "<body><div id=\"chart_div\" style=\"width: 900px; height: 500px;\"></div></body>";
                report = report + "<BODY STYLE=\"background-repeat: no-repeat;\"><TABLE width=\"100%\"><TR><TD><center><div id=\"chart_div\" style=\"width: 900px; height: 500px;\"></div></center></TD>";
                //report = report + "<TD><A href=\"TestResultReport.htm\" style=\"text-decoration:none\" target=\"resultframe\"><font face=\"Calibri\" color=black size=6> Test cases executed:&nbsp&nbsp&nbsp&nbsp&nbsp" + totalTestCases + "</BR></font></A>";
                //report = report + "<A href=\"TestPassedResultReport.htm\" style=\"text-decoration:none\" target=\"resultframe\"><font face=\"Calibri\" color=green size=6>Total passed test cases:&nbsp&nbsp" + passedTestCases + "</font></BR></A>";
                //report = report + "<A href=\"TestFailedResultReport.htm\" style=\"text-decoration:none\" target=\"resultframe\"><font face=\"Calibri\" color=red size=6>Total failed test cases:&nbsp&nbsp&nbsp&nbsp" + failedTestCases + "</font></BR></TD>";
                report = report + "</TR></TABLE></BODY></HTML>";

                reportWriter.WriteLine(report);
            }
            catch (Exception exception)
            {
                // Print exception on console
                Console.WriteLine("CreateFinalReport " + "Final Report writer. Error in generating final report" + exception.Message);
            }
            finally
            {
                reportWriter.Close();
            }
        }
        #endregion

        #region Function to generate blank html file
        /// <summary>
        /// Function to generate final html report
        /// </summary>
        /// <author>rajan.bansod</author> 
        /// <ModifiedBy>rajan.bansod</ModifiedBy>
        /// <Date>12-Aug-2011</Date>
        private void CreateBlankFile(string testRunPath)
        {
            StreamWriter reportWriter = null;
            try
            {
                testRunPath = testRunPath + @"\HtmlFiles";
                if (!Directory.Exists(testRunPath))
                {
                    Directory.CreateDirectory(testRunPath);
                }

                FileInfo fileInfo = new FileInfo(testRunPath + @"\Blank.htm");
                reportWriter = fileInfo.AppendText();

                //string report = "<HTML><frameset rows=\"16%,*\" framespacing=\"0\" frameborder=\"no\" border=\"0\"><frame src=\"ReportHeader.htm\" /><frameset cols=\"20%,*\" framespacing=\"0\" frameborder=\"no\" border=\"0\"><frame src=\"TestResultReport.htm\" /><frame name=\"targetframe\"/></frameset></frameset></HTML>";
                string report = "<HTML><BODY STYLE=\"background-repeat: no-repeat;\"><TABLE width=\"100%\"></TABLE></BODY></HTML>";
                
                reportWriter.WriteLine(report);
            }
            catch (Exception exception)
            {
                // Print exception on console
                Console.WriteLine("CreateFinalReport " + "Final Report writer. Error in generating final report" + exception.Message);
            }
            finally
            {
                reportWriter.Close();
            }
        }
        #endregion
                
        #region Function to generate final html report
        /// <summary>
        /// Function to generate final html report
        /// </summary>
        /// <author>rajan.bansod</author> 
        /// <ModifiedBy>rajan.bansod</ModifiedBy>
        /// <Date>12-Aug-2011</Date>
        private void CreateFinalReport(string testRunPath)
        {
            StreamWriter reportWriter = null;
            try
            {
                if (!Directory.Exists(testRunPath))
                {
                    Directory.CreateDirectory(testRunPath);
                }

                FileInfo fileInfo = new FileInfo(testRunPath + @"\FinalTestReport.htm");
                reportWriter = fileInfo.AppendText();

                //string report = "<HTML><frameset rows=\"16%,*\" framespacing=\"0\" frameborder=\"no\" border=\"0\"><frame src=\"ReportHeader.htm\" /><frameset cols=\"20%,*\" framespacing=\"0\" frameborder=\"no\" border=\"0\"><frame src=\"TestResultReport.htm\" /><frame name=\"targetframe\"/></frameset></frameset></HTML>";
                string report = "<HTML><frameset rows=\"16%,*\" framespacing=\"0\" frameborder=\"no\" border=\"0\"><frame src=\"HtmlFiles/ReportHeader.htm\" />";
                report = report + "<frameset cols=\"15%,*\" framespacing=\"0\" frameborder=\"no\" border=\"0\"><frame src=\"HtmlFiles/blank.htm\" name=\"resultframe\"/>";
                report = report + "<frame src=\"HtmlFiles/Dashboard.htm\" name=\"targetframe\"/></frameset></frameset></HTML>";

                reportWriter.WriteLine(report);
            }
            catch (Exception exception)
            {
                // Print exception on console
                Console.WriteLine("CreateFinalReport " + "Final Report writer. Error in generating final report" + exception.Message);
            }
            finally
            {
                reportWriter.Close();
            }
        }
        #endregion

        #region Function to compile html report from csv log files
        /// <summary>
        /// Function to complie report from csv log files
        /// </summary>
        /// <author>rajan.bansod</author> 
        /// <ModifiedBy>rajan.bansod</ModifiedBy>
        /// <Date>12-Aug-2011</Date>
        public string CompileReportFromCSV()
        {
            string embeddedMailContents = null;
            try
            {
                // declare and initialse all variables
                int totalTestCases = 0;
                int totalPassedTestCases = 0;
                int totalFailedTestCases = 0;
                int totalInfoTestCases = 0;
                int totalSteps = 0;
                int totalFailedSteps = 0;
                int totalPassedSteps = 0;
                int totalInfoSteps = 0;
                int count = 0;
                ArrayList testCaseIds = new ArrayList();
                ArrayList testResult = new ArrayList();
                string testCaseData = "";

                //DateTime startTimeOfExecution = new DateTime();
                //DateTime endTimeOfExecution = new DateTime();
                DateTime startTimeOfExecutionTC = new DateTime();
                DateTime endTimeOfExecutionTC = new DateTime();

                // connaction string to read csv file
                string connString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + this.LogFileDirectory + "\\" + @";Extended Properties=""text;HDR=No;FMT=Delimited(,)""";
                OleDbConnection connection = new OleDbConnection(connString);
                OleDbCommand command = new OleDbCommand();

                try
                {
                    // open connection for csv file
                    command.Connection = connection;
                    connection.Open();
                }
                catch (Exception exception)
                {
                    Console.WriteLine("Open Connection exception: " + exception.Message);
                }

                // query to fetch data from first column of csv file
                string query = "SELECT DISTINCT F1 FROM Logs.csv";
                command.CommandText = query;
                OleDbDataReader testDataReader = command.ExecuteReader();
                //oConnection.Close();
                while (testDataReader.Read())
                {
                    if (testDataReader.GetValue(0).Equals("TCID"))
                    {
                        continue;
                    }
                    testCaseIds.Add(testDataReader.GetValue(0).ToString());
                }

                count = 0;
                testDataReader.Close();

                do
                {
                    try
                    {
                        // query to fetch start time data of perticuler test case from fift column of csv file
                        string queryStartTimeTC = "SELECT MIN(F5) FROM Logs.csv WHERE F1='" + testCaseIds[count].ToString() + "'";
                        command.CommandText = queryStartTimeTC;
                        object startTimeTC = command.ExecuteScalar();
                        startTimeOfExecutionTC = Convert.ToDateTime(startTimeTC.ToString());
                    }
                    catch (Exception) { }

                    try
                    {
                        // query to fetch end time data of perticuler test case from fift column of csv file
                        string queryEndTimeTC = "SELECT MAX(F5) FROM Logs.csv WHERE F1='" + testCaseIds[count].ToString() + "'";
                        command.CommandText = queryEndTimeTC;
                        object endTimeTC = command.ExecuteScalar();
                        endTimeOfExecutionTC = Convert.ToDateTime(endTimeTC.ToString());
                    }
                    catch (Exception) { }

                    // Time differance between start time of test case and end time of test case.
                    TimeSpan timeSpanTC = endTimeOfExecutionTC.Subtract(startTimeOfExecutionTC);

                    // query to fetch data from first column of csv file
                    query = "SELECT * FROM Logs.csv WHERE F1='" + testCaseIds[count].ToString() + "'";
                    command.CommandText = query;
                    testDataReader = command.ExecuteReader();
                    string sHTML = "<HTML><HEAD><STYLE type=\"text/css\"> ps {color:green} fl {color:red}</STYLE></HEAD><BODY background=\"../../../imgscr/MainBody.png\" style=\"background-repeat: no-repeat;\">";
                    string sHTMLHeader = "<TABLE width=\"100%\" border=1><TR bgcolor=\"#A5DBEB\"><TD> <center><font size=\"2\" face=\"Calibri\"><a href=\"" + testCaseIds[count].ToString() + ".htm\" target=\"targetframe\" style=\"text-decoration:none\"><STRONG>" + testCaseIds[count].ToString() + "</STRONG></font></center></TD></TR></TABLE>";
                    sHTMLHeader = sHTMLHeader + "<TABLE width=\"100%\" border=1><TR bgcolor=\"#DBF0F7\"><TD><font size=\"2\" face=\"Calibri\"> <center> Start Time: " + startTimeOfExecutionTC.ToString() + "</center></TD><TD> <font size=\"2\" face=\"Calibri\"><center> Finish Time: " + endTimeOfExecutionTC.ToString() + "</center></TD><TD> <font size=\"2\" face=\"Calibri\"><center> Execution Time: " + timeSpanTC.Hours + " Hrs, " + timeSpanTC.Minutes + " Mins, " + timeSpanTC.Seconds + " Secs " + "</center></font></TD></TR></TABLE>";
                    string sTestData = "";
                    bool OddEvenStepFlag = false;
                    while (testDataReader.Read())
                    {
                        if (OddEvenStepFlag)
                        {
                            sTestData = sTestData + "<TR bgcolor=\"#B8E2EF\">";
                            OddEvenStepFlag = false;
                        }
                        else
                        {
                            sTestData = sTestData + "<TR bgcolor=\"#DBF0F7\">";
                            OddEvenStepFlag = true;
                        }

                        if (testDataReader.GetValue(3).ToString().ToUpper().Equals("PASS"))
                        {
                            //sTestData = sTestData + "<TD><ps><font color=green>" + testDataReader.GetValue(0) + "</font></ps></TD>";
                            sTestData = sTestData + "<TD width=\"4%\" align=\"center\"><ps><font size=\"2\" face=\"Calibri\" color=green>" + testDataReader.GetValue(1) + "</font></ps></TD>";
                            sTestData = sTestData + "<TD width=\"70%\" align=\"left\"><ps><font size=\"2\" face=\"Calibri\" color=green>" + testDataReader.GetValue(2) + "</font></ps></TD>";
                            sTestData = sTestData + "<TD width=\"6%\" align=\"center\"><ps><font size=\"2\" face=\"Calibri\" color=green>" + testDataReader.GetValue(3) + "</font></ps></TD>";
                            sTestData = sTestData + "<TD width=\"20%\" align=\"center\"><ps><font size=\"2\" face=\"Calibri\" color=green>" + testDataReader.GetValue(4) + "</font></ps></TD>";
                            sTestData = sTestData + "</TR>";
                        }
                        if (testDataReader.GetValue(3).ToString().ToUpper().Equals("FAIL"))
                        {
                            //sTestData = sTestData + "<TD><fl><font color=red>" + testDataReader.GetValue(0) + "</font></fl></TD>";
                            sTestData = sTestData + "<TD width=\"4%\" align=\"center\"><fl><font size=\"2\" face=\"Calibri\" color=red><A HREF=\"./" + testDataReader.GetValue(0) + "_" + testDataReader.GetValue(1) + ".png\" style=\"text-decoration:none\">" + testDataReader.GetValue(1) + "</A></font></fl></TD>";
                            sTestData = sTestData + "<TD width=\"70%\" align=\"left\"><fl><font size=\"2\" face=\"Calibri\" color=red>" + testDataReader.GetValue(2) + "</font></fl></TD>";
                            sTestData = sTestData + "<TD width=\"6%\" align=\"center\"><fl><font size=\"2\" face=\"Calibri\" color=red><A HREF=\"./" + testDataReader.GetValue(0) + "_" + testDataReader.GetValue(1) + "_Captured.png\" style=\"text-decoration:none\">" + testDataReader.GetValue(3) + "</A></font></fl></TD>";
                            sTestData = sTestData + "<TD width=\"20%\" align=\"center\"><fl><font size=\"2\" face=\"Calibri\" color=red>" + testDataReader.GetValue(4) + "</font></fl></TD>";
                            sTestData = sTestData + "</TR>";
                        }
                        if (testDataReader.GetValue(3).ToString().ToUpper().Equals("INFO"))
                        {
                            //sTestData = sTestData + "<TD><fl><font color=red>" + testDataReader.GetValue(0) + "</font></fl></TD>";
                            sTestData = sTestData + "<TD width=\"4%\" align=\"center\"><fl><font size=\"2\" face=\"Calibri\" color=blue>" + testDataReader.GetValue(1) + "</font></ps></TD>";
                            sTestData = sTestData + "<TD width=\"70%\" align=\"left\"><fl><font size=\"2\" face=\"Calibri\" color=blue>" + testDataReader.GetValue(2) + "</font></fl></TD>";
                            sTestData = sTestData + "<TD width=\"6%\" align=\"center\"><fl><font size=\"2\" face=\"Calibri\" color=blue>" + testDataReader.GetValue(3) + "</font></fl></TD>";
                            sTestData = sTestData + "<TD width=\"20%\" align=\"center\"><fl><font size=\"2\" face=\"Calibri\" color=blue>" + testDataReader.GetValue(4) + "</font></fl></TD>";
                            sTestData = sTestData + "</TR>";
                        }
                    }
                    testCaseData = sHTML + "<TABLE width=\"100%\">" + sTestData + "</TABLE></BODY></HTML>";
                    sHTMLHeader = sHTML + sHTMLHeader + "</BODY></HTML>";
                    try
                    {
                        CreateTestCaseFile(testCaseIds[count].ToString(), testCaseData, sHTMLHeader, this.ReportFileDirectory);
                    }
                    catch (Exception exception)
                    {
                        Console.WriteLine("Failed to generate html log file for test case " + testCaseIds[count].ToString() + " and the exception is " + exception.Message);
                    }
                    count++;
                    testDataReader.Close();
                } while (count < testCaseIds.Count);

                //try
                //{
                //    // query to fetch start time data from fift column of csv file
                //    string queryStartTime = "SELECT top 1 F5 FROM Logs.csv WHERE F1 <> 'TCID'";
                //    command.CommandText = queryStartTime;
                //    object startTime = command.ExecuteScalar();
                //    startTimeOfExecution = Convert.ToDateTime(startTime.ToString());
                //}
                //catch (Exception) { }

//                try
//                {
//                    // query to fetch end time data from fift column of csv file
//                    string queryEndTime = "SELECT MAX(F5) FROM Logs.csv WHERE F1 <> 'TCID'";
////                    string queryTotalCount = "SELECT COUNT(*) FROM Logs.csv WHERE F1 <> 'TCID'";
////                    command.CommandText = queryTotalCount;
////                    object totalCount = command.ExecuteScalar();
////                    string recordCount = totalCount.ToString();
////                    int prevRecord = int.Parse(totalCount.ToString())-1;
////                    string prevRecordCount = prevRecord.ToString();
//////                    string queryEndTime = "SELECT top " + recordCount + " F5 FROM Logs.csv WHERE F1 <> 'TCID'";
//////                    string queryEndTime = "SELECT F5 FROM Logs.csv WHERE F5 = IDENT_CURRENT('Logs.csv')"; 
////                    string queryEndTime = "SELECT F5 FROM Logs.csv limit " + prevRecordCount + "," + recordCount; 
//                    command.CommandText = queryEndTime;
//                    object endTime = command.ExecuteScalar();
//                    endTimeOfExecution = Convert.ToDateTime(endTime.ToString());
//                }
//                catch (Exception) { }

                // query to fetch data from fourth column of csv file
                query = "SELECT COUNT(*) FROM Logs.csv WHERE F4='FAIL'";
                command.CommandText = query;
                object failedSteps = command.ExecuteScalar();
                totalFailedSteps = Convert.ToInt16(failedSteps.ToString());
                failedSteps = null;

                // query to fetch data from fourth column of csv file
                query = "SELECT COUNT(*) FROM Logs.csv WHERE F4='PASS'";
                command.CommandText = query;
                object passedSteps = command.ExecuteScalar();
                totalPassedSteps = Convert.ToInt16(passedSteps.ToString());
                passedSteps = null;

                // query to fetch data from fourth column of csv file
                query = "SELECT COUNT(*) FROM Logs.csv WHERE F4='INFO'";
                command.CommandText = query;
                object infoSteps = command.ExecuteScalar();
                totalInfoSteps = Convert.ToInt16(infoSteps.ToString());
                infoSteps = null;

                // query to fetch count of data from csv file
                query = "SELECT COUNT(*) FROM Logs.csv";
                command.CommandText = query;
                object allSteps = command.ExecuteScalar();
                totalSteps = Convert.ToInt16(allSteps.ToString()) - 1;
                allSteps = null;

                totalTestCases = testCaseIds.Count;

                count = 0;
                do
                {
                    int pass = 0;
                    int fail = 0;
                    int info = 0;
                    query = "SELECT F1,F4 FROM Logs.csv WHERE F1='" + testCaseIds[count] + "'";
                    command.CommandText = query;
                    if (!testDataReader.IsClosed)
                    {
                        testDataReader.Close();
                    }
                    testDataReader = command.ExecuteReader();
                    while (testDataReader.Read())
                    {
                        if (testDataReader.GetValue(1).ToString().ToUpper().Equals("PASS"))
                        {
                            pass = pass + 1;
                        }
                        if (testDataReader.GetValue(1).ToString().ToUpper().Equals("FAIL"))
                        {
                            fail = fail + 1;
                        }
                        if (testDataReader.GetValue(1).ToString().ToUpper().Equals("INFO"))
                        {
                            info = info + 1;
                        }
                    }
                    if (fail == 0 && info == 0)
                    {
                        totalPassedTestCases = totalPassedTestCases + 1;
                        testResult.Add(testCaseIds[count].ToString() + ",PASS");
                    }
                    else if (fail == 0 && info != 0)
                    {
                        totalInfoTestCases = totalInfoTestCases + 1;
                        totalPassedTestCases = totalPassedTestCases + 1;
                        testResult.Add(testCaseIds[count].ToString() + ",INFO_PASS");
                    }
                    else if (fail != 0 && info != 0)
                    {
                        totalInfoTestCases = totalInfoTestCases + 1;
                        totalFailedTestCases = totalFailedTestCases + 1;
                        testResult.Add(testCaseIds[count].ToString() + ",INFO_FAIL");
                    }
                    else if (fail != 0 && info == 0)
                    {
                        totalFailedTestCases = totalFailedTestCases + 1;
                        testResult.Add(testCaseIds[count].ToString() + ",FAIL");
                    }
                    else
                    {
                        totalFailedTestCases = totalFailedTestCases + 1;
                        testResult.Add(testCaseIds[count].ToString() + ",FAIL");
                    }
                    count++;
                    testDataReader.Close();
                } while (count < testCaseIds.Count);
                connection.Close();

                try
                {
                    // Generate Graph for current test report.
                    GraphicalReport graphicalReport = new GraphicalReport(totalTestCases.ToString(), totalPassedTestCases.ToString(), totalFailedTestCases.ToString(), this.ReportFileDirectory);
                    graphicalReport.GeneratePieGraph();
                }
                catch (Exception exception)
                {
                    // Print exception on console
                    Console.WriteLine("CompileReportFromCSV => GenerateGraph: " + exception.Message);
                }
                // Generate Custome HTML report
                CreateBlankFile(this.ReportFileDirectory);
                CreateDashboard(this.ReportFileDirectory, totalTestCases.ToString(), totalPassedTestCases.ToString(), totalFailedTestCases.ToString());
                CreateHeaderReport(totalTestCases, totalPassedTestCases, totalFailedTestCases, totalInfoTestCases, totalSteps, totalFailedSteps, totalPassedSteps, totalInfoSteps, this.ReportFileDirectory, this.ServerName);
                CreateTestResultReport(testResult, this.ReportFileDirectory);
                CreatePassedTestResultReport(testResult, this.ReportFileDirectory);
                CreateFailedTestResultReport(testResult, this.ReportFileDirectory);
                CreateFinalReport(this.ReportFileDirectory);
                embeddedMailContents = CreateEmbeddedMailContents(totalTestCases, totalPassedTestCases, totalFailedTestCases, totalInfoTestCases, this.ServerName, this.ReportFileDirectory);
                return embeddedMailContents;
            }
            catch (Exception exception)
            {
                // Print exception on console
                Console.WriteLine("CompileReportFromCSV " + exception.Message);
                return embeddedMailContents;
            }
            finally
            {
                GC.Collect();
            }
        }
        #endregion       
    }
}


