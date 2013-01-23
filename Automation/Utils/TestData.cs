using System;
using System.Collections;
using System.IO;
using System.Xml;
using System.Security.AccessControl;
using System.Collections.Generic;


namespace Automation.TestScripts
{
    public class TestData 
    {
        private IDictionary _testDataTable = new Dictionary<string ,string>();

        public IDictionary TestDataTable
        {
            get
            {
                return _testDataTable;
            }
        }


        public TestData(string testDataFilePath) 
        {
            LoadTestData(testDataFilePath);
        }


        public void LoadTestData(string testDataFilePath)
        {
            System.IO.File.SetAttributes(testDataFilePath, System.IO.FileAttributes.Normal);
            FileStream fStream = new FileStream(testDataFilePath, FileMode.Open);
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
                        _testDataTable.Add(name, reader.Value);                       
                    }

                }
                fStream.Close();
                Console.Read();
            }
            catch (Exception exception)
            {
                fStream.Close();
                Console.Read();
                Console.WriteLine("Exception in Loading Test Data" + exception.Message);
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
    }
}


