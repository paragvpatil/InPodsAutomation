using System;
using Automation.Development.Browsers;
using Automation.Development.Pages;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Automation.TestScripts;

namespace Automation.TestScripts
{
    [TestClass]
    public class BaseClass : TestCaseUtil
    {

        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }
        
        /// <summary>
        /// Test Class initialize 
        /// </summary>
        [ClassInitialize()]
        protected void ClassInitialize()
        {
            
        }

        /// <summary>
        /// Test Class cleanup
        /// </summary>
        [ClassCleanup()]
        protected void ClassCleanup()
        {
 
        }

    }
}