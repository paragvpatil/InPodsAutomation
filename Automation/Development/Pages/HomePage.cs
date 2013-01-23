using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;
using Automation.Development.Browsers;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System.Collections.Generic;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Interactions.Internal;
using Automation.TestScripts;

namespace Automation.Development.Pages
{
    public class HomePage : ObjectRepository
    {  

        public static int timeoutWait = 10;

        public Browser Browser { get; private set; }

        protected HomePage() { }

        public HomePage(Browser browser)
        {
           
            if (browser == null)
            {
                throw new ArgumentNullException("browser");
            }
            this.Browser = browser;
           


        }
        ObjectRepository objectRepository;
        public void LocateControls()
        {
            string pageName = "HomePage";
            string objectRepositoryFilePath = PrepareObjectRepositoryFilePath(pageName);
            objectRepository = new ObjectRepository(objectRepositoryFilePath);            

        }

        #region Function to find select control by using selector
        /// <summary>
        /// Function to find select control by using selector
        /// </summary>
        /// <param name="selector">name of selector</param>
        /// <returns>IWebElement</returns>
        public IWebElement FindControlBySelector(string selector)
        {
            return this.Browser.FindControlBySelector(selector);
        }
        #endregion

        #region Function to find control by selector
        /// <summary>
        /// Function to find control by selector
        /// </summary>
        /// <param name="selector">name of selector</param>
        /// <returns>IWebElement</returns>
        public SelectElement FindSelectControlBySelector(string selector)
        {
            return this.Browser.FindSelectControlBySelector(selector);
        }
        #endregion

        #region Function to find control by using xpath
        /// <summary>
        /// Function to find control by using xpath
        /// </summary>
        /// <param name="key">xpath of element</param>
        /// <returns>IWebElement</returns>
        public IWebElement FindControlByXPath(string key)
        {
            return this.Browser.FindControlByXpath(key);
        }
        #endregion

        #region Function to find control by id
        /// <summary>
        /// Function to find control by id
        /// </summary>
        /// <param name="key">id of element</param>
        /// <returns>IWebElement</returns>
        public IWebElement FindControlById(string key)
        {
            return this.Browser.FindControlById(key);
        }
        #endregion

        #region Function to find control by classname
        /// <summary>
        /// Function to find control by classname
        /// </summary>
        /// <param name="className">classname of element</param>
        /// <returns>IWebElement</returns>
        public IWebElement FindControlByClass(string className)
        {
            return this.Browser.FindControlByClass(className);
        }
        #endregion

        #region Function to find control by linktext
        /// <summary>
        /// Function to find control by linktext
        /// </summary>
        /// <param name="key">linktext of element</param>
        /// <returns>IWebElement</returns>
        public IWebElement FindControlByLinkText(string key)
        {
            return this.Browser.FindControlByLinkText(key);
        }
        #endregion

        #region Function to find control by class
        /// <summary>
        /// Function to find control by class
        /// </summary>
        /// <param name="className">class of element</param>
        /// <returns>collection of IWebElement</returns>
        public ReadOnlyCollection<IWebElement> FindControlsByClass(string className)
        {
            return this.Browser.FindControlsByClass(className);
        }
        #endregion

        #region Function to find control by xpath
        /// <summary>
        /// Function to find control by xpath
        /// </summary>
        /// <param name="key">xpath of element</param>
        /// <returns>collection of IWebElement</returns>
        public ReadOnlyCollection<IWebElement> FindControlsByXPath(string key)
        {
            return this.Browser.FindControlsByXpath(key);
        }
        #endregion

        #region Function to find control by tagname
        /// <summary>
        /// Function to find control by tagname
        /// </summary>
        /// <param name="key">tagname</param>
        /// <returns>collection of IWebElement</returns>
        public ReadOnlyCollection<IWebElement> FindControlsByTagName(string key)
        {
            return this.Browser.FindControlsByTagName(key);
        }
        #endregion

        #region Function to find control by name
        /// <summary>
        /// Function to find control by name
        /// </summary>
        /// <param name="key">name of element</param>
        /// <returns>IWebElement</returns>
        public IWebElement FindControlByName(string key)
        {
            return this.Browser.FindControlByName(key);
        }
        #endregion

        #region Function to find control by partial linktext
        /// <summary>
        /// Function to find control by partial linktext
        /// </summary>
        /// <param name="key">partiallinktext of element</param>
        /// <returns>IWebElement</returns>
        public IWebElement FindControlByPartialLinkText(string key)
        {
            return this.Browser.FindControlByPartialLinkText(key);
        }
        #endregion

        #region Function to find select control by xpath
        /// <summary>
        /// Function to find select control by xpath
        /// </summary>
        /// <param name="key">xpath of element</param>
        /// <returns>selectelement</returns>
        public SelectElement FindSelectControlByXpath(string key)
        {
            return this.Browser.FindSelectControlByXpath(key);
        }
        #endregion

        #region Function to find select control by class
        /// <summary>
        /// Function to find select control by class
        /// </summary>
        /// <param name="key">classname of element</param>
        /// <returns>selectelement</returns>
        public SelectElement FindSelectControlByClass(string className)
        {
            return this.Browser.FindSelectControlByClass(className);
        }
        #endregion

        #region Function to find select control by id
        /// <summary>
        /// Function to find select control by id
        /// </summary>
        /// <param name="key">id of element</param>
        /// <returns>selectelement</returns>
        public SelectElement FindSelectControlById(string id)
        {
            return this.Browser.FindSelectControlById(id);
        }
        #endregion

        #region Function to switch on alert window
        /// <summary>
        /// Function to switch on alert window
        /// </summary>
        /// <returns>IAlert</returns>
        public IAlert SwitchToAlert()
        {
            return this.Browser.SwitchToAlert();
        }
        #endregion

        #region Function to switch on window with index
        /// <summary>
        /// Function to switch on window with index
        /// </summary>
        /// <returns></returns>
        public void SwitchToWindowWithIndex(int windowHandleIndex)
        {
            this.Browser.SwitchToWindowWithIndex(windowHandleIndex);
        }
        #endregion

        #region Function to switch on window using windowtitle
        /// <summary>
        /// Function to switch on window using windowtitle
        /// </summary>
        /// <param name="windowTitle">window title name</param>
        /// <returns>boolean</returns>
        public bool SwitchToWindow(string windowTitle)
        {
            try
            {
                int flagWindowFound = 0;
                int count = 0;
                string currentHandle = null;
                ReadOnlyCollection<string> windowHandles = null;
                do
                {
                    windowHandles = this.Browser.Driver.WindowHandles;
                    foreach (string handle in windowHandles)
                    {
                        if (this.Browser.Driver.SwitchTo().Window(handle).Title.Equals(windowTitle))
                        {
                            flagWindowFound = 1;
                            currentHandle = handle;
                            break;
                        }
                    }
                    if (1 == flagWindowFound)
                    {
                        this.Browser.Driver.SwitchTo().Window(currentHandle);
                        break;
                    }

                    ImplicitlyWait(500);
                    count++;
                }
                while (timeoutWait != count);
                if (timeoutWait == count)
                {
                    return false;
                }
                return true;
            }
            catch (Exception exception)
            {
                Console.WriteLine("Exception in switch window " + exception.Message);
                return false;
            }
        }
        #endregion

        #region Function to switch from onr window to the another using windowtitle
        /// <summary>
        /// Function to switch from onr window to the another using windowtitle
        /// </summary>
        /// <param name="windowTitleOne">first window title name</param>
        /// <param name="windowTitleTwo">second window title name</param>
        /// <returns>boolean</returns>
        public bool SwitchToWindow(string windowTitleOne, string windowTitleTwo)
        {
            try
            {
                ReadOnlyCollection<string> windowHandles = null;
                do
                {
                    windowHandles = this.Browser.Driver.WindowHandles;
                    this.Browser.Driver.SwitchTo().Window(this.Browser.Driver.WindowHandles[0]);
                    windowHandles = this.Browser.Driver.WindowHandles;
                    foreach (string handle in windowHandles)
                    {
                        if (this.Browser.Driver.SwitchTo().Window(handle).Title == windowTitleOne || this.Browser.Driver.SwitchTo().Window(handle).Title == windowTitleTwo)
                        {
                            this.Browser.Driver.SwitchTo().Window(handle);
                            break;
                        }
                    }
                }
                while (!this.Browser.Driver.SwitchTo().Window(this.Browser.Driver.CurrentWindowHandle).Title.Equals(windowTitleOne) || this.Browser.Driver.SwitchTo().Window(this.Browser.Driver.CurrentWindowHandle).Title.Equals(windowTitleTwo));
                return true;
            }
            catch (Exception exception)
            {
                Console.WriteLine("Exception in switch window " + exception.Message);
                return false;
            }
        }
        #endregion

        #region Function to switch between Frames
        /// <summary>
        /// WebDriver switches between Frame
        /// Note: GlobalVariables.PASS the Frame id to switch
        /// </summary>
        /// <param name="frameID">e.g. ID_FRAME_EMAIL</param>
        /// <returns></returns>
        public bool SwitchToFrame(string frameID, string elementType)
        {

            int flagFrameFound = 0;
            int count = 0;
            int frameCount = 0;
            string typeOfElement = null;
            elementType = elementType.ToUpper();
            try
            {
                ReadOnlyCollection<IWebElement> frameHandles = null;
                switch (elementType)
                {
                    case "ID":
                        typeOfElement = elementType;
                        break;
                    case "SRC":
                        typeOfElement = elementType;
                        break;
                    default:
                        throw new ArgumentException("unknown element type. ");
                }
                do
                {
                    IWebElement htmlElement = this.FindControlByXPath("//html");
                    frameHandles = this.FindControlsByTagName("iframe");
                    for (frameCount = 0; frameCount < frameHandles.Count; frameCount++)
                    {
                        if (frameHandles[frameCount].GetAttribute(typeOfElement.ToLower()).ToLower().Contains(frameID.ToLower()))
                        {
                            flagFrameFound = 1;
                            break;
                        }
                    }
                    if (1 == flagFrameFound)
                    {
                        this.Browser.Driver.SwitchTo().Frame(frameCount);
                        break;
                    }
                    ImplicitlyWait(500);
                    count++;
                }
                while (timeoutWait != count);
                if (timeoutWait == count)
                {
                    return false;
                }
                return true;
            }
            catch (Exception exception)
            {
                Console.WriteLine("Exception in switch frame " + exception.Message);
                return false;
            }
        }
        #endregion

        #region Function to wait for an object
        /// <summary>
        /// Function to wait for an object
        /// </summary>
        /// <param name="elementType"></param>
        /// <param name="elementValue"></param>
        /// <returns> bool </returns>
        /// <Date>02-Sept-2011</Date>
        public bool WaitForElement(string elementType, string elementValue)
        {
            int count = 0;
            do
            {
                try
                {
                    IWebElement element = null;
                    switch (elementType)
                    {
                        case "XPATH": element = this.FindControlByXPath(elementValue);
                            break;
                        case "ID": element = this.FindControlById(elementValue);
                            break;
                        case "NAME": element = this.FindControlByName(elementValue);
                            break;
                        case "CLASS": element = this.FindControlByClass(elementValue);
                            break;
                    }

                    // Verify if element present or not
                    if (null != element)
                    {
                        if (element.Displayed)
                        {
                            return true;
                        }
                    }
                }
                catch (Exception) { }
                count++;
                ImplicitlyWait(500);
            }
            while (count < timeoutWait);
            return false;
        }
        #endregion

        #region Function to wait for an object
        /// <summary>
        /// Function to wait for an object
        /// </summary>
        /// <param name="elementType"></param>
        /// <param name="elementValue"></param>
        /// <returns> bool </returns>
        /// <Date>02-Sept-2011</Date>
        public bool WaitForElement(string elementType, string elementValue, int userDefinedTime)
        {
            int count = 0;
            userDefinedTime = userDefinedTime * 2;
            do
            {
                try
                {
                    IWebElement element = null;
                    switch (elementType)
                    {
                        case "XPATH": element = this.FindControlByXPath(elementValue);
                            break;
                        case "ID": element = this.FindControlById(elementValue);
                            break;
                        case "NAME": element = this.FindControlByName(elementValue);
                            break;
                        case "CLASS": element = this.FindControlByClass(elementValue);
                            break;
                    }

                    // Verify if element present or not
                    if (null != element)
                    {
                        if (element.Displayed)
                        {
                            return true;
                        }
                    }
                }
                catch (Exception)
                {
                }
                count++;
                ImplicitlyWait(500);
            }
            while (count < userDefinedTime);
            return false;
        }
        #endregion

       
        #region Function to logout user
        /// <summary>
        /// Function to logout user
        /// </summary>
        /// <returns>boolean</returns>
        public bool LogOut()
        {
            // code to logout. Common to all pages.
            try
            {
                LocateControls();
                IWebElement logoutDiv = this.FindControlByXPath((string)objectRepository.ObjectRepositoryTable["AccountDiv"]);
                Actions action = new Actions(this.Browser.Driver);
                action.MoveToElement(logoutDiv).Perform();
                this.WaitForElementToGetEnabled("XPATH", (string)objectRepository.ObjectRepositoryTable["LogOutLink"]);
                IWebElement logOutLinkControl = this.FindControlByXPath((string)objectRepository.ObjectRepositoryTable["LogOutLink"]);
                logOutLinkControl.Click();
                ImplicitlyWait(2000);
                this.Browser.Driver.Manage().Cookies.DeleteAllCookies();
                this.Browser.Close();
                ImplicitlyWait(500);

                KillAlreadyOpenBrowsers();
                ImplicitlyWait(2000);
                string[] processName = { "QTAgentHandler", "QTAgentApplicationErrorHandler", "UploadFileScriptError" };
                for (int count = 0; count <= 2; count++)
                {
                    // Get array of process name
                    Process[] localByName = Process.GetProcessesByName(processName[count]);
                    foreach (Process item in localByName)
                    {
                        // kill the process
                        item.Kill();
                    }
                }
                return true;
            }
            catch (Exception exception)
            {
                Console.WriteLine("Exception in Logout " + exception.Message);
            }
            return true;
        }
        #endregion

       

        #region Function for Implicitly wait
        /// <summary>
        /// Implicitly wait
        /// </summary>
        /// <Date>22-Aug-2011</Date>
        public void ImplicitlyWait(int timeInMiliSecond)
        {
            try
            {
                Thread.Sleep(timeInMiliSecond);
            }
            catch (Exception)
            {
            }
        }
        #endregion

        #region function to get XPATH count
        /// <summary>
        /// Function to get XPATH count
        /// </summary>
        /// <param name="locator">xpath of the element</param>
        /// <returns>count of XPATH</returns>
        /// <Date>12-Aug-2011</Date>
        public int GetTableCount(string locator)
        {
            return this.Browser.FindControlsByXpath(locator).Count();
        }
        #endregion

        #region Mouse Events
        /// <summary>
        /// Mouse Events
        /// </summary>
        /// <param name="dwFlags"></param>
        /// <param name="dx"></param>
        /// <param name="dy"></param>
        /// <param name="dwData"></param>
        /// <param name="dwExtraInfo"></param>
        [DllImport("user32.dll")]
        static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);
        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "SetCursorPos")]
        internal extern static int SetCursorPos(int x, int y);
        #endregion

        #region Mouse Event Flags
        /// <summary>
        /// Mouse Event Flags
        /// </summary>
        [Flags]
        public enum MouseEventFlags
        {
            LEFTDOWN = 0x00000002,
            LEFTUP = 0x00000004,
            MIDDLEDOWN = 0x00000020,
            MIDDLEUP = 0x00000040,
            MOVE = 0x00000001,
            ABSOLUTE = 0x00008000,
            RIGHTDOWN = 0x00000008,
            RIGHTUP = 0x00000010,
            WHEEL = 0x00000800,
            XDOWN = 0x00000080,
            XUP = 0x00000100
        }
        #endregion

        #region Move the mouse from one location to another
        /// <summary>
        /// Move the mouse from one location to another
        /// </summary>
        /// <param name="xCordinate"></param>
        /// <param name="yCordinate"></param>
        public void MouseMove(int xCordinate, int yCordinate)
        {
            SetCursorPos(xCordinate, yCordinate);
            mouse_event((int)(MouseEventFlags.MOVE), xCordinate, yCordinate, 0, 0);
        }
        #endregion

        #region Move the mouse wheel up and down
        /// <summary>
        /// Move the mouse wheel up and down
        /// Note - Roll the mouse wheel. Delta of 120 wheels up once normally, -120 wheels down once normally
        /// </summary>
        /// <param name="delta">mouse move up or down</param>
        public void MouseScrollUpDown(int delta)
        {
            //PressKey(0x12);
            //PressKey(0x10);
            mouse_event((int)(MouseEventFlags.WHEEL), 0, 0, delta, 0);
            //ReleaseKey(0x10);
            //ReleaseKey(0x12);
        }
        #endregion

        #region Function to press left button of mouse
        /// <summary>
        /// Function to press left button of mouse
        /// </summary>
        /// <param name="xCordinate">xCordinate or position</param>
        /// <param name="yCordinate">yCordinate or position</param>
        public void LeftClick(int xCordinate, int yCordinate)
        {
            SetCursorPos(xCordinate, yCordinate);
            mouse_event((int)(MouseEventFlags.LEFTDOWN), xCordinate, yCordinate, 0, 0);
            mouse_event((int)(MouseEventFlags.LEFTUP), xCordinate, yCordinate, 0, 0);
        }
        #endregion

        #region Function to press right button of mouse
        /// <summary>
        /// Function to press right button of mouse
        /// </summary>
        /// <param name="xCordinate">xCordinate or position</param>
        /// <param name="yCordinate">yCordinate or position</param>
        public void RightClick(int xCordinate, int yCordinate)
        {
            SetCursorPos(xCordinate, yCordinate);
            Cursor.Position = new System.Drawing.Point(xCordinate, yCordinate);
            mouse_event((int)(MouseEventFlags.RIGHTDOWN), 0, 0, 0, 0);
            mouse_event((int)(MouseEventFlags.RIGHTUP), 0, 0, 0, 0);
        }
        #endregion
        
        #region Function to get the instance of browser
        /// <summary>
        /// Function to get the instance of browser
        /// </summary>
        /// <returns>instance of the browser</returns>
        public Browser getBrowser()
        {
            try
            {
                return this.Browser;
            }
            catch (Exception exception)
            {
                throw new Exception("Exception in clicking preferences button" + exception.Message);
            }
        }
        #endregion

        #region Function for Press Key
        /// <summary>
        /// Function for Press Key
        /// </summary>
        /// <param name="buttonVirtualKey">struct System.byte</param>
        /// <returns></returns>
        /// <Date>14-Sept-2011</Date>        
        [DllImport("user32.dll", SetLastError = true)]
        static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);
        public bool PressKey(byte buttonVirtualKey)
        {
            try
            {
                keybd_event(buttonVirtualKey, 0, 0, 0);
                return true;
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
                return false;
            }
        }
        #endregion

        #region Function for Release Key
        /// <summary>
        /// Function for Release Key
        /// </summary>
        /// <param name="buttonVirtualKey">struct System.byte</param>
        /// <returns></returns>
        /// <author>rajan.bansod</author> 
        /// <ModifiedBy>rajan.bansod</ModifiedBy>
        /// <Date>14-Sept-2011</Date>
        public bool ReleaseKey(byte buttonVirtualKey)
        {
            try
            {
                keybd_event(buttonVirtualKey, 0, 0x0002, 0);
                return true;
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
                return false;
            }
        }
        #endregion

        
        #region Function to find the control br link text
        /// <summary>
        /// Function to find the control br link text
        /// </summary>
        /// <param name="linkText">link text of the element</param>
        /// <returns>collection of element</returns>
        public ReadOnlyCollection<IWebElement> FindControlsByLinkText(string linkText)
        {
            return this.Browser.FindControlsByLinkText(linkText);

        }
        #endregion

        #region Function to maximize the window
        /// <summary>
        /// Function to maximize the window
        /// </summary>
        /// <param name="url">url of application</param>
        /// <returns></returns>
        /// <Date>14-Sept-2011</Date>
        public void MaximizeWindow(string url)
        {
            try
            {
                ((IJavaScriptExecutor)this.Browser.Driver).ExecuteScript("window.moveTo(0,0);");
                ((IJavaScriptExecutor)this.Browser.Driver).ExecuteScript("window.resizeTo(screen.width, screen.height-25);");
                ((IJavaScriptExecutor)this.Browser.Driver).ExecuteScript("window.focus();");
                ((IJavaScriptExecutor)this.Browser.Driver).ExecuteScript("window.focus();");
            }
            catch (Exception exception)
            {
                throw new Exception("Failed to maximize the window" + exception.Message);
            }
        }
        #endregion

        #region Function to enable NumLock key
        /// <summary>
        /// Function to enable NumLock key
        /// </summary>
        /// <param></param>
        /// <returns></returns>
        /// <Date>20-Sept-2011</Date>
        //[DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true, CallingConvention = CallingConvention.Winapi)]
        //[DllImport("user32.dll", SetLastError = true)]
        [DllImport("user32.dll")]
        public static extern short GetKeyState(int nVirtKey);
        public bool EnableNumLock()
        {
            try
            {
                //check status of numlock key
                short status = GetKeyState(0x90);
                if (status == 0)
                {
                    PressKey(0x90);
                    ReleaseKey(0x90);
                }
                return true;
            }
            catch (Exception exception)
            {
                throw new Exception("Failed to enable the num lock key: " + exception.Message);
            }
        }
        #endregion

        #region Function to disable NumLock key
        /// <summary>
        /// Function to disable NumLock key
        /// </summary>
        /// <param></param>
        /// <returns>boolean</returns>
        public bool DisableNumLock()
        {
            try
            {
                //status of num lock key
                short status = GetKeyState(0x90);
                if (status == 1)
                {
                    PressKey(0x90);
                    ReleaseKey(0x90);
                }
                return true;
            }
            catch (Exception exception)
            {
                throw new Exception("Failed disable numlock key: " + exception.Message);
            }
        }
        #endregion

        #region Function to double click on element
        /// <summary>
        /// Function to double click on element
        /// </summary>
        /// <param name="xCordinate">xcordinate value</param>
        /// <param name="yCordinate">ycordinate value</param>
        /// <returns>void</returns>
        public void DoubleClick(int xCordinate, int yCordinate)
        {
            int MOUSEEVENTF_LEFTDOWN = 0x02;
            int MOUSEEVENTF_LEFTUP = 0x04;
            SetCursorPos(xCordinate + 10, yCordinate + 10);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, xCordinate, yCordinate, 0, 0);
            this.ImplicitlyWait(150);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, xCordinate, yCordinate, 0, 0);
        }
        #endregion

        #region Select area using mouse
        /// <summary>
        /// Function to select area using mouse
        /// </summary>
        /// <param name="xCordinate">xcordinate value</param>
        /// <param name="yCordinate">ycordinate value</param>
        /// <returns></returns>
        /// <Date>19-Sept-2011</Date> 
        public void MouseSelectArea(int xCordinate, int yCordinate)
        {
            SetCursorPos(xCordinate, yCordinate);
            ImplicitlyWait(3000);
            mouse_event((int)(MouseEventFlags.LEFTDOWN), xCordinate, yCordinate, 0, 0);
            MouseMove(xCordinate + 20, yCordinate + 20);
            mouse_event((int)(MouseEventFlags.LEFTUP), xCordinate + 20, yCordinate + 20, 0, 0);
        }
        #endregion

        #region function for kill open processes
        /// <summary>
        /// Function to Kill All Already Open Internet Explorer instances
        /// </summary>
        /// <returns>string value or Exception</returns>
        /// <Date>12-Aug-2011</Date>
        public void KillProcess(string processName)
        {
            try
            {
                // get ie process names to array
                Process[] localByName = Process.GetProcessesByName(processName);
                foreach (Process item in localByName)
                {
                    // kill the process
                    item.Kill();
                }
            }
            catch (Exception exception)
            {
                throw new Exception("Exception occur to kill process." + exception.Message);
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

        #region Function to disable caps lock key
        /// <summary>
        /// Function to disable caps lock key
        /// </summary>
        /// <param></param>
        /// <returns>boolean</returns>
        /// <Date>20-Sept-2011</Date>
        public bool DisableCapsLock()
        {
            try
            {
                //check status of caps lock key
                short status = GetKeyState(0x14);
                if (status == 1)
                {
                    PressKey(0x14);
                    ReleaseKey(0x14);
                }
                return true;
            }
            catch (Exception exception)
            {
                throw new Exception("Failed to enable the num lock key: " + exception.Message);
            }
        }
        #endregion

       
        #region Function to check pop up alert for 15 seconds
        /// <summary>
        /// Function to check pop up alert for 15 seconds
        /// </summary>
        /// <param></param>
        /// <returns>boolean</returns>
        /// <Date>30-Sept-2011</Date>
        public void AcceptAlert(bool isAccept)
        {
            try
            {
                int count = 0;
                do
                {
                    try
                    {
                        IAlert alert = this.Browser.Driver.SwitchTo().Alert();
                        if (null != alert)
                        {
                            if (isAccept)
                            {
                                alert.Accept();
                            }
                            else
                            {
                                alert.Dismiss();
                            }
                            ImplicitlyWait(2000);
                            break;
                        }
                    }
                    catch
                    {
                        count++;
                        ImplicitlyWait(500);
                    }
                } while (count < 10);
            }
            catch (Exception)
            {
                throw new Exception("Failed to check alert");
            }
        }
        #endregion

        #region Function to determines if string array is sorted from A -> Z
        /// <summary>
        /// Determines if string array is sorted from A -> Z
        /// </summary>
        /// <returns>boolean</returns>
        public bool IsAscSorted(string[] arr)
        {
            for (int i = 1; i < arr.Length; i++)
            {
                if (arr[i - 1].CompareTo(arr[i]) > 0) // If previous is bigger, return false
                {
                    return false;
                }
            }
            return true;
        }
        #endregion


        #region function to determines if string array is sorted from Z -> A
        /// <summary>
        /// Determines if string array is sorted from Z -> A
        /// </summary>
        /// <returns>boolean</returns>
        public bool IsSortedDescending(string[] arr)
        {
            for (int i = arr.Length - 2; i >= 0; i--)
            {
                if (arr[i].CompareTo(arr[i + 1]) < 0) // If previous is smaller, return false
                {
                    return false;
                }
            }
            return true;
        }
        #endregion

        #region function to determines if second string array is sorted from Z -> A
        /// <summary>
        /// Determines if string array is sorted from Z -> A
        /// </summary>
        /// <returns>boolean</returns>
        public bool IsSortedSecondArrayDescending(string[] arrFirst, string[] arrSecond)
        {
            //loop for first array
            for (int upCount = 0; upCount < arrFirst.Length; upCount++)
            {
                //loop for first and second array
                for (int fCount = upCount + 1; fCount < arrFirst.Length; fCount++)
                {
                    //check first array contains to equal value
                    if (arrFirst[upCount].Equals(arrFirst[fCount]))
                    {
                        if (arrSecond[upCount].CompareTo(arrSecond[fCount]) < 0)// If previous is smaller, return false
                        {
                            return false;
                        }
                    }
                }
            }
            return true;
        }
        #endregion

        #region Function to maximize the window
        /// <summary>
        /// Function to maximize the window
        /// </summary>
        /// <param></param>
        /// <returns></returns>
        /// <Date>14-Sept-2011</Date>
        public void MaximizeWindowWithAltD()
        {
            try
            {
                PressKey(0x5B);
                PressKey(0x44);
                ReleaseKey(0x5B);
                ReleaseKey(0x44);
                this.ImplicitlyWait(1000);
                string browserName  = this.Browser.Driver.GetType().ToString();
                if (browserName == "OpenQA.Selenium.IE.InternetExplorerDriver")
                {
                    ((IJavaScriptExecutor)this.Browser.Driver).ExecuteScript("window.focus();");
                    ((IJavaScriptExecutor)this.Browser.Driver).ExecuteScript("window.moveTo(0,0);");
                    ((IJavaScriptExecutor)this.Browser.Driver).ExecuteScript("window.focus();");
                    ((IJavaScriptExecutor)this.Browser.Driver).ExecuteScript("window.resizeTo(screen.width, screen.height-25);");
                    ((IJavaScriptExecutor)this.Browser.Driver).ExecuteScript("window.focus();");
                    ((IJavaScriptExecutor)this.Browser.Driver).ExecuteScript("window.focus();");
                }
                else
                {
                    ((IJavaScriptExecutor)this.Browser.Driver).ExecuteScript("window.focus();");
                    this.Browser.Driver.Manage().Window.Maximize();
                    
                }               
            }
            catch (Exception exception)
            {
                throw new Exception("Failed to maximize the window" + exception.Message);
            }
        }
        #endregion

        #region Release all keyboard key
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        /// 
        public bool ReleaseAllKey()
        {
            try
            {
                byte[] key = { 0x10, 0x11, 0x09, 0x12, 0x0D };
                short keyStatus;
                for (int keyCount = 0; keyCount < key.Length; keyCount++)
                {
                    keyStatus = GetKeyState(key[keyCount]);
                    if (0 != keyStatus)
                    {
                        ReleaseKey(key[keyCount]);
                    }
                }
                return true;
            }
            catch (Exception exception)
            {
                throw new Exception("Failed to release key board key." + exception.Message);
            }
        }
        #endregion        

       
        #region Function to wait for an object to get enabled
        /// <summary>
        /// Function to wait for an object to get enabled
        /// </summary>
        /// <param name="elementType"></param>
        /// <param name="elementValue"></param>
        public bool WaitForElementToGetEnabled(string elementType, string elementValue)
        {
            int count = 0;
            do
            {
                try
                {
                    IWebElement element = null;
                    switch (elementType)
                    {
                        case "XPATH": element = this.FindControlByXPath(elementValue);
                            break;
                        case "ID": element = this.FindControlById(elementValue);
                            break;
                        case "NAME": element = this.FindControlByName(elementValue);
                            break;
                        case "CLASS": element = this.FindControlByClass(elementValue);
                            break;
                    }

                    // Verify if element present or not
                    if (null != element)
                    {
                        if (element.Enabled)
                        {
                            return true;
                        }
                    }
                }
                catch (Exception) { }
                count++;
                ImplicitlyWait(500);
            }
            while (count < timeoutWait);
            return false;
        }
        #endregion        

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public WorkspacePage clickWorkSpaceButton()
        {
            LocateControls();
            IWebElement workspaceButtonControl = this.FindControlByXPath((string) objectRepository.ObjectRepositoryTable["WorkspaceButton"]);
            workspaceButtonControl.Click();            
            Actions action = new Actions(this.Browser.Driver);
            try
            {

                bool isConfigureButtonEnabled = this.WaitForElementToGetEnabled("XPATH", (string)objectRepository.ObjectRepositoryTable["ConfigureButton"]);
                if (isConfigureButtonEnabled)
                {
                    IWebElement configureButtonControl = this.FindControlByXPath((string)objectRepository.ObjectRepositoryTable["ConfigureButton"]);
                    action.MoveToElement(configureButtonControl).Click().Perform();                   

                }
            }catch(Exception){}
            ImplicitlyWait(5000);
            return new WorkspacePage(this.Browser);
        }
    }
}

