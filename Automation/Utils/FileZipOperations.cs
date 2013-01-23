using System;
using System.Collections;
using System.IO;
using ICSharpCode.SharpZipLib.Zip;


namespace Automation.TestScripts
{
    public class FileZipOperations
    {
        #region Constructeurs
        public FileZipOperations()
        {
        }

        public FileZipOperations(string inputFolderPath, string outputPathAndFile, string password)
        {
            this.InputFolderPath = inputFolderPath;
            this.OutputPathAndFile = outputPathAndFile;
            this.Password = password;
        }
        #endregion Constructeurs

        string _sInputFolderPath;
        ///
        /// Report File Directory
        ///
        public string InputFolderPath
        {
            get { return _sInputFolderPath; }
            set { _sInputFolderPath = value; }
        }

        string _sOutputPathAndFile;
        ///
        /// Report File Directory
        ///
        public string OutputPathAndFile
        {
            get { return _sOutputPathAndFile; }
            set { _sOutputPathAndFile = value; }
        }

        string _sPassword;
        ///
        /// Report File Directory
        ///
        public string Password
        {
            get { return _sPassword; }
            set { _sPassword = value; }
        }

        #region Function to unzip the files
        /// <summary>
        /// Function to unzip the files
        /// </summary>
        /// <param name="zipPathAndFile"></param>
        /// <param name="outputFolder"></param>
        /// <param name="password"></param>
        /// <param name="deleteZipFile"></param>
        public void UnZipFiles(string zipPathAndFile, string outputFolder, string password, bool deleteZipFile)
        {
            try
            {
                ZipInputStream s = new ZipInputStream(File.OpenRead(zipPathAndFile));
                if (password != null && password != String.Empty)
                    s.Password = password;
                ZipEntry theEntry;
                string tmpEntry = String.Empty;
                while ((theEntry = s.GetNextEntry()) != null)
                {
                    string directoryName = outputFolder;
                    string fileName = Path.GetFileName(theEntry.Name);
                    // create directory 
                    if (directoryName != "")
                    {
                        Directory.CreateDirectory(directoryName);
                    }
                    if (fileName != String.Empty)
                    {
                        if (theEntry.Name.IndexOf(".ini") < 0)
                        {
                            string fullPath = directoryName + "\\" + theEntry.Name;
                            fullPath = fullPath.Replace("\\ ", "\\");
                            string fullDirPath = Path.GetDirectoryName(fullPath);
                            if (!Directory.Exists(fullDirPath)) Directory.CreateDirectory(fullDirPath);
                            FileStream streamWriter = File.Create(fullPath);
                            int size = 2048;
                            byte[] data = new byte[2048];
                            while (true)
                            {
                                size = s.Read(data, 0, data.Length);
                                if (size > 0)
                                {
                                    streamWriter.Write(data, 0, size);
                                }
                                else
                                {
                                    break;
                                }
                            }
                            streamWriter.Close();
                        }
                    }
                }
                s.Close();
                if (deleteZipFile)
                    File.Delete(zipPathAndFile);
            }
            catch (Exception Ex)
            {
                Console.WriteLine(Ex.Message.ToString());
            }

        }
        #endregion

        #region Function to generate file list for zip/unzip operations
        /// <summary>
        /// Function to generate file list for zip/unzip operations
        /// </summary>
        /// <param name="Dir"></param>
        /// <returns></returns>
        private ArrayList GenerateFileList(string Dir)
        {
            ArrayList files = new ArrayList();
            try
            {
                bool Empty = true;
                foreach (string file in Directory.GetFiles(Dir)) // add each file in directory
                {
                    files.Add(file);
                    Empty = false;
                }

                if (Empty)
                {
                    if (Directory.GetDirectories(Dir).Length == 0)
                    // if directory is completely empty, add it
                    {
                        files.Add(Dir + @"/");
                    }
                }

                foreach (string dirs in Directory.GetDirectories(Dir)) // recursive
                {
                    foreach (object obj in GenerateFileList(dirs))
                    {
                        files.Add(obj);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return files; // return file list
        }
        #endregion

        #region Function to zip the files
        /// <summary>
        /// Function to zip the files
        /// </summary>
        /// <param name="inputFolderPath"></param>
        /// <param name="outputPathAndFile"></param>
        /// <param name="password"></param>
        public void ZipFiles()
        {
            FileStream ostream = null;
            byte[] obuffer;
            ZipOutputStream oZipStream = null;
            ZipEntry oZipEntry = null;
            try
            {
                ArrayList ar = GenerateFileList(this.InputFolderPath); // generate file list
                int TrimLength = (Directory.GetParent(this.InputFolderPath)).ToString().Length;
                // find number of chars to remove     // from orginal file path
                TrimLength += 1; //remove '\'           
                string newinputFolderPath;
                newinputFolderPath = this.InputFolderPath.Substring(0, this.InputFolderPath.LastIndexOf('\\'));
                string outPath = this.OutputPathAndFile;
                oZipStream = new ZipOutputStream(File.Create(outPath)); // create zip stream
                if (this.Password != null && this.Password != String.Empty)
                {
                    oZipStream.Password = this.Password;
                }
                oZipStream.SetLevel(9); // maximum compression
                
                foreach (string Fil in ar) // for each file, generate a zipentry
                {
                    oZipEntry = new ZipEntry(Fil.Remove(0, TrimLength));
                    oZipStream.PutNextEntry(oZipEntry);

                    if (!Fil.EndsWith(@"/")) // if a file ends with '/' its a directory
                    {
                        ostream = File.OpenRead(Fil);
                        obuffer = new byte[ostream.Length];
                        ostream.Read(obuffer, 0, obuffer.Length);
                        oZipStream.Write(obuffer, 0, obuffer.Length);
                    }
                }
            }
            catch (Exception exception)
            {
                // Print exception on console
                Console.WriteLine("ZipFiles " + "Error in generating Zip Files" + exception.Message);
            }
            finally
            {
                obuffer = null;
                oZipEntry = null;
                ostream.Close();
                oZipStream.Finish();
                oZipStream.Close();
            }
        }
        #endregion

    }
}


