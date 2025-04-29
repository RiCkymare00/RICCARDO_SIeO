using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleAppTemp
{
    internal class FTP
    {
        string host = ConfigurationManager.AppSettings["FtpWins"].ToString();
        string UserId = ConfigurationManager.AppSettings["FtpUserIdWins"].ToString();
        string Password = ConfigurationManager.AppSettings["FtpPassWordWins"].ToString();
        string AGLEMPath = ConfigurationManager.AppSettings["PathFolderCopre"].ToString();

        public bool CreateFolder(string host, string cartellaCliente)
        {
            bool IsCreated = true;
            try
            {
                // Verifica se la cartella esiste già
                if (!FolderExists(host + cartellaCliente))
                {
                    // Se la cartella non esiste, prova a crearla
                    WebRequest request = WebRequest.Create(host + AGLEMPath);
                    request.Method = WebRequestMethods.Ftp.MakeDirectory;
                    request.Credentials = new NetworkCredential(UserId, Password);
                    using (var resp = (FtpWebResponse)request.GetResponse())
                    {
                        Console.WriteLine(resp.StatusCode);
                    }
                }
            }
            catch (Exception ex)
            {
                IsCreated = false;
            }
            return IsCreated;
        }

        // Metodo per verificare se una cartella esiste su un server FTP
        private bool FolderExists(string folderUri)
        {
            bool result = false;
            try
            {
                WebRequest request = WebRequest.Create(folderUri);
                request.Method = WebRequestMethods.Ftp.ListDirectory;
                request.Credentials = new NetworkCredential(UserId, Password);

                using (var resp = (FtpWebResponse)request.GetResponse())
                {
                    // Se non ci sono errori, la cartella esiste
                    result = true;
                }
            }
            catch (WebException ex)
            {
                // Se viene generata un'eccezione, la cartella non esiste
                result = false;
            }
            return result;
        }

        public void UploadFile(string From, string nomeFile, string To)
        {
            //ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
            try
            {
                Uri uri = new Uri($"{To}/{nomeFile}");

                FtpWebRequest request = (FtpWebRequest)WebRequest.Create(uri);
                request.Method = WebRequestMethods.Ftp.UploadFile;

                request.Credentials = new NetworkCredential(UserId, Password);
                request.EnableSsl = false;  // Abilita TLS
                request.UsePassive = false;
                request.UseBinary = true;

                byte[] fileContents = File.ReadAllBytes(From);
                request.ContentLength = fileContents.Length;

                using (Stream requestStream = request.GetRequestStream())
                {
                    requestStream.Write(fileContents, 0, fileContents.Length);
                }

                using (FtpWebResponse response = (FtpWebResponse)request.GetResponse())
                {
                    _ = new LogWriter($"Upload File Complete, status {response.StatusDescription}");
                }
            }
            catch (WebException ex)
            {
                FtpWebResponse response = (FtpWebResponse)ex.Response;
                _ = new LogWriter($"FTP UploadFile error: {response.StatusDescription}");
            }
            catch (Exception ex)
            {
                Exception currentException = ex;

                while (currentException != null)
                {
                    _ = new LogWriter("Exception FTP UploadFile: " + currentException.Message);
                    currentException = currentException.InnerException;
                }
            }
        }

        public List<string> GetAllFtpFiles(string ParentFolderpath)
        {
            try
            {
                FtpWebRequest ftpRequest = (FtpWebRequest)WebRequest.Create(ParentFolderpath);
                ftpRequest.Credentials = new NetworkCredential(UserId, Password);
                ftpRequest.Method = WebRequestMethods.Ftp.ListDirectory;
                FtpWebResponse response = (FtpWebResponse)ftpRequest.GetResponse();
                StreamReader streamReader = new StreamReader(response.GetResponseStream());

                List<string> directories = new List<string>();

                string line = streamReader.ReadLine();
                while (!string.IsNullOrEmpty(line))
                {
                    var lineArr = line.Split('/');
                    line = lineArr[lineArr.Length - 1];
                    directories.Add(line);
                    line = streamReader.ReadLine();
                }

                streamReader.Close();

                return directories;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void DownloadFileFTP(string ftpFullPath)
        {
            string percorsoLocaleDoveSalvareIFile = ConfigurationManager.AppSettings["percorsoDoveSalvareIFile"];

            using (WebClient request = new WebClient())
            {
                request.Credentials = new NetworkCredential(UserId, Password);
                byte[] fileData = request.DownloadData(ftpFullPath);

                using (FileStream file = File.Create(percorsoLocaleDoveSalvareIFile))
                {
                    file.Write(fileData, 0, fileData.Length);
                    file.Close();
                }
            }
        }

        public string DeleteFile(string fileName)
        {
            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(host + "IN/" + fileName);
            request.Method = WebRequestMethods.Ftp.DeleteFile;
            request.Credentials = new NetworkCredential(UserId, Password);

            using (FtpWebResponse response = (FtpWebResponse)request.GetResponse())
            {
                return response.StatusDescription;
            }
        }

        public void DownloadAllFilesFTP()
        {
            FtpWebRequest ftpRequest = (FtpWebRequest)WebRequest.Create(host + "IN");
            ftpRequest.Credentials = new NetworkCredential(UserId, Password);
            ftpRequest.Method = WebRequestMethods.Ftp.ListDirectory;
            FtpWebResponse response = (FtpWebResponse)ftpRequest.GetResponse();
            StreamReader streamReader = new StreamReader(response.GetResponseStream());
            List<string> directories = new List<string>();

            string line = streamReader.ReadLine();
            while (!string.IsNullOrEmpty(line))
            {
                directories.Add(line);
                line = streamReader.ReadLine();
            }
            streamReader.Close();


            using (WebClient ftpClient = new WebClient())
            {
                ftpClient.Credentials = new System.Net.NetworkCredential(UserId, Password);

                for (int i = 0; i <= directories.Count - 1; i++)
                {
                    if (directories[i].Contains("."))
                    {
                        string path = host + "IN/" + directories[i].ToString();
                        string trnsfrpth = ConfigurationManager.AppSettings["percorsoDoveSalvareIFile"] + "\\" + directories[i].ToString();
                        ftpClient.DownloadFile(path, trnsfrpth);
                    }
                }
            }
        }
    }
}
