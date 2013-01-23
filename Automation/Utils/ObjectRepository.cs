using System;
using System.Collections;
using System.IO;
using System.Xml;
using System.Security.AccessControl;
using System.Collections.Generic;


namespace Automation.TestScripts
{
    public class ObjectRepository
    {
        private IDictionary _objectRepositoryTable = new Dictionary<string, string>();
        
        public IDictionary ObjectRepositoryTable
        {
            get
            {
                return _objectRepositoryTable;
            }
        }

        public ObjectRepository()
        {
            
        }
        public ObjectRepository(string objectRepositoryFilePath)
        {
            LoadObjectRepository(objectRepositoryFilePath);
        }

            protected string currentDirectory = String.Empty;
            protected string projectDirectory = String.Empty;
            protected string testDataDirectory = String.Empty;

            public void LoadObjectRepository(string objectRepositoryFilePath)
                {
                    System.IO.File.SetAttributes(objectRepositoryFilePath, System.IO.FileAttributes.Normal);
                    FileStream fStream = new FileStream(objectRepositoryFilePath, FileMode.Open);
                    XmlTextReader reader = new XmlTextReader(fStream);
                    try
                    {
                        string name = null;
                        reader.WhitespaceHandling = WhitespaceHandling.None;
                        while (reader.Read())
                        {

                            if (reader.NodeType == XmlNodeType.Element)
                            {
                                name = reader.Name;

                            }
                            else if (reader.NodeType == XmlNodeType.Text)
                            {
                                _objectRepositoryTable.Add(name, reader.Value);
                            }

                        }
                        fStream.Close();
                        Console.Read();
                    }
                    catch (Exception exception)
                    {
                        fStream.Close();
                        Console.Read();
                        Console.WriteLine("Exception in Loading ObjectRepository" + exception.Message);
                    }
            }

        // this function will be usefull for reading xml (recursive call of function)
        public static void ClearAttributes(string currentDir)
        {
            if (Directory.Exists(currentDir))
            {
                string[] subDirs = Directory.GetDirectories(currentDir);
                foreach (string dir in subDirs)
                    ClearAttributes(dir);
                string[] files = files = Directory.GetFiles(currentDir);
                foreach (string file in files)
                    File.SetAttributes(file, FileAttributes.Normal);
            }
        }

        #region Function to get script name
        /// <summary>
        /// Function to get current test script name
        /// </summary>
        /// <returns>script name</returns>        
        protected string GetPageName()
        {
            string pageName = "";
            try
            {
                // get class name 
                pageName = this.ToString();

                // replace folder names from scriptName
                pageName = pageName.Replace("Automation.", "");
                pageName = pageName.Replace("Development.", "");
                pageName = pageName.Replace("Pages.", "");

                // return scriptName
                return pageName;
            }
            catch (Exception exception)
            {
                // write log
                Console.WriteLine(exception.Message);
                return pageName;
            }
        }
        #endregion
        /// <summary>
        /// 
        /// </summary>
        /// <param name="testScriptName"></param>
        /// <returns></returns>
        protected string PrepareObjectRepositoryFilePath(string pageName)
        {
            Configuration oConfig = new Configuration();
            this.currentDirectory = oConfig.CurrentDirectory;
            this.projectDirectory = oConfig.ProjectDirectory;
           
            return this.currentDirectory + this.projectDirectory + "\\ObjectRepository\\" + pageName + ".xml";
        }
    }
}


