using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using OpenQA.Selenium.Chrome;

namespace Automation.Development.Browsers
{
    public class BrowserFactory
    {
        private static BrowserFactory _instance = null;

        private BrowserFactory() { }

        public static BrowserFactory Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new BrowserFactory();
                }
                return _instance;
            }
        }

        /// <summary>
        /// This function create browser instance
        /// </summary>
        /// <param name="key">Name of browser</param>
        /// <param name="scriptName">Test script name</param>
        /// <param name="fileName">Runtime.xml</param>
        /// <returns></returns>
        public Browser GetBrowser(string key, string scriptName, string fileName,string driverPath)
        {
            switch (key)
            {
                case "Chrome":
                    return new Chrome(driverPath);
                case "FireFox":
                    return new FireFox();
                case "IE":
                    InternetExplorer internetExplorer = null;
                    try
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
                        catch { }
                        internetExplorer = new InternetExplorer(driverPath);                       


                    }
                    catch (Exception)
                    {
                        
                        GC.Collect();
                        System.Diagnostics.Process[] localByName = System.Diagnostics.Process.GetProcessesByName("QTAgent32");
                        Console.WriteLine(" Killed process name: " + localByName[0]);
                        localByName[0].Kill();
                        System.Threading.Thread.Sleep(3000);
                        internetExplorer = new InternetExplorer(driverPath);
                    }
                    return internetExplorer;
                default:
                    throw new ArgumentException("unknown browser type: " + key);
            }
        }
    }
}
