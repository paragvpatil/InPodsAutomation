using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Net;

namespace Automation.TestScripts
{
    public class GraphicalReport
    {
        #region Constructeurs
        public GraphicalReport()
        {
        }

        public GraphicalReport(string totalTestCases, string passedTestCases, string failedTestCases, string imgFilePath)
        {
            this.TotalTestCases = totalTestCases;
            this.PassedTestCases = passedTestCases;
            this.FailedTestCases = failedTestCases;
            this.ImgFilePath = imgFilePath;
        }
        #endregion Constructeurs

        string _sTotalTestCases;
        ///
        /// Report File Directory
        ///
        public string TotalTestCases
        {
            get { return _sTotalTestCases; }
            set { _sTotalTestCases = value; }
        }

        string _sPassedTestCases;
        ///
        /// Report File Directory
        ///
        public string PassedTestCases
        {
            get { return _sPassedTestCases; }
            set { _sPassedTestCases = value; }
        }

        string _sFailedTestCases;
        ///
        /// Report File Directory
        ///
        public string FailedTestCases
        {
            get { return _sFailedTestCases; }
            set { _sFailedTestCases = value; }
        }

        string _sImgFilePath;
        ///
        /// Report File Directory
        ///
        public string ImgFilePath
        {
            get { return _sImgFilePath; }
            set { _sImgFilePath = value; }
        }

        #region Function to generate graph file using result counts
        /// <summary>
        /// Function to generate graph file using result counts
        /// </summary>
        /// <param name="passedTestCases">no of passed test cases</param>
        /// <param name="failedTestCases">no of failed test cases</param>
        /// <param name="imgFilePath">report directory path</param>
        public void GeneratePieGraph()
        {
            try
            {
                // Verify if directory is present or not, if not, create directory 
                this.ImgFilePath = this.ImgFilePath + @"\HtmlFiles";
                if (!Directory.Exists(this.ImgFilePath))
                {
                    Directory.CreateDirectory(this.ImgFilePath);
                }

                // Generate URL
                string chartUrl = "http://chart.apis.google.com/chart?chs=500x250&cht=p&chco=007800,BD0000&chd=t:" + this.PassedTestCases + "," + this.FailedTestCases + "&chdl=Passed+Test+Cases|Failed+Test+Cases&chl=Passed+" + this.PassedTestCases + "|Failed+" + this.FailedTestCases + "&chtt=ComputeNext+Test+Automation+Summery&chts=675050,11.5";
                byte[] chartBytes = null;
                // Access URL and download its contents
                WebClient client = new WebClient();
                chartBytes = client.DownloadData(chartUrl);

                using (var memStream = new MemoryStream())
                {
                    memStream.Write(chartBytes, 0, chartBytes.Length);
                }
                
                // Save image in given location
                File.WriteAllBytes(this.ImgFilePath + @"\excel_chart_export.bmp", chartBytes);
            }
            catch (Exception exception)
            {
                // Print exception on console
                Console.WriteLine("GenerateGraph: " + exception.Message);
            }
        }
        #endregion

        #region Function to generate graph file using result counts
        /// <summary>
        /// Function to generate graph file using result counts
        /// </summary>
        /// <param name="passedTestCases">no of passed test cases</param>
        /// <param name="failedTestCases">no of failed test cases</param>
        /// <param name="imgFilePath">report directory path</param>
        public void GenerateGraph()
        {
            try
            {
                this.ImgFilePath = this.ImgFilePath + @"\HtmlFiles";
                if (!Directory.Exists(this.ImgFilePath))
                {
                    Directory.CreateDirectory(this.ImgFilePath);
                }

                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                //add data to excel file
                xlWorkSheet.Cells[1, 1] = "";
                xlWorkSheet.Cells[1, 2] = "Result";

                xlWorkSheet.Cells[2, 1] = " Executed Test Cases: " + this.TotalTestCases;
                xlWorkSheet.Cells[2, 2] = 0;

                xlWorkSheet.Cells[3, 1] = "Failed Test Cases: " + this.FailedTestCases;
                xlWorkSheet.Cells[3, 2] = this.FailedTestCases;

                xlWorkSheet.Cells[4, 1] = "Passed Test Cases: " + this.PassedTestCases;
                xlWorkSheet.Cells[4, 2] = this.PassedTestCases;

                Excel.Range chartRange;
                
                Excel.ChartObjects xlCharts = (Excel.ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);

                Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 80, 300, 250);
                Excel.Chart chartPage = myChart.Chart;
                               

                chartRange = xlWorkSheet.get_Range("A1", "B4");
                chartPage.SetSourceData(chartRange, misValue);
                                
                chartPage.ChartType = Excel.XlChartType.xl3DPie;

                //export chart as picture file
                chartPage.Export(this.ImgFilePath + @"\excel_chart_export.bmp", "BMP", misValue);

                xlWorkBook.SaveAs(this.ImgFilePath + "\\csharp.net-informations.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

                DeleteFile(this.ImgFilePath + "\\csharp.net-informations.xls");
                //MessageBox.Show("Excel file created , you can find the file " + imgFilePath + @"\excel_chart_export.bmp");
            }
            catch (Exception exception)
            {
                // Print exception on console
                Console.WriteLine("GenerateGraph: " + exception.Message);
            }
        }
        #endregion

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Exception Occured while releasing object " + ex.ToString());
                //MessageBox.Show();
            }
            finally
            {
                GC.Collect();
            }
        }

        public void DeleteFile(string fileName)
        {
            try
            {
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
                        Console.WriteLine("Failed to delete file => " + exception.Message);
                    }
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine("Failed to delete file => " + exception.Message);
            }
            finally
            {
            }
        }


    }
}
