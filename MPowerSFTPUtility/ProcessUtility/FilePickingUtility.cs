using System;
using System.Configuration;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Net.Mail;
using System.Reflection.Metadata;
using System.Text.RegularExpressions;
using ExcelDataReader;
using Renci.SshNet;
using static System.Net.WebRequestMethods;

public class FilePickingUtility
{


    public void ProcessUtilityAsyncs()
    {

        string logFolderPath = ConfigurationManager.AppSettings["ErrorLogFilePath"];
        string mappingFilePath = ConfigurationManager.AppSettings["mappingFilePath"];
        string inFolderPath = ConfigurationManager.AppSettings["inFolderPath"];
        string inBkpFolderPath = ConfigurationManager.AppSettings["inBkpFolderPath"];
        string backupFolderPath = ConfigurationManager.AppSettings["backupFolderPath"];

        string sftpHost = ConfigurationManager.AppSettings["SFTPhost"];
        int sftpPort = Convert.ToInt32(ConfigurationManager.AppSettings["SFTPport"]);
        string sftpUser = ConfigurationManager.AppSettings["SFTPusername"];
        string sftpPass = ConfigurationManager.AppSettings["SFTPpassword"];



        //Bloomberg destination
        string BloombergHost = ConfigurationManager.AppSettings["BloombergHost"];
        int BloombergPort = Convert.ToInt32(ConfigurationManager.AppSettings["BloombergPort"]);
        string BloombergUserID = ConfigurationManager.AppSettings["BloombergUserID"];
        string BloombergPassword = ConfigurationManager.AppSettings["BloombergPassword"];
        string BloombergInputPath = ConfigurationManager.AppSettings["BloombergInput"];


        //Fact Set 
        string FactSetHost = ConfigurationManager.AppSettings["FactSetHost"];
        int FactSetPort = Convert.ToInt32(ConfigurationManager.AppSettings["FactSetPort"]);
        string FactSetUserID = ConfigurationManager.AppSettings["FactSetUserID"];
        string FactSetPassword = ConfigurationManager.AppSettings["FactSetPassword"];
        string FactSetInputPath = ConfigurationManager.AppSettings["FactSetInput"];
        string FactSetBackupPath = ConfigurationManager.AppSettings["FactSetBackup"];

        string successFolderPath = ConfigurationManager.AppSettings["successFolderPath"];
        string errorFolderPath = ConfigurationManager.AppSettings["errorFolderPath"];


        DateTime startTime = DateTime.Now;
        WriteLog(logFolderPath, $"Utility started at {startTime:dd-MMM-yy HH:mm:ss}");


        try
        {
            if (!System.IO.File.Exists(mappingFilePath))
            {
                WriteLog(logFolderPath, "Configuration file not found. Utility stopped.");
                //SendEmail("Configuration file not found. Utility stopped.");
                return;
            }

            var mapping = LoadMapping(mappingFilePath, logFolderPath);
            if (!mapping.Columns.Contains("File Name Pattern") || !mapping.Columns.Contains("Folder Path"))
            {
                WriteLog(logFolderPath, "Mapping file format incorrect (missing required columns). Utility stopped.");
                //SendEmail("Mapping file format incorrect (missing required columns). Utility stopped.");

                return;
            }

            if (!Directory.Exists(inFolderPath))
            {
                WriteLog(logFolderPath, $"Input folder does not exist: {inFolderPath}");
                return;
            }

            // Now fetch files + folder files based on pattern
            var files = GetAllFilesToProcess(inFolderPath, mapping, logFolderPath);
            WriteLog(logFolderPath, $"Total files fetched for processing: {files.Count}");

            if (files.Count == 0)
            {
                WriteLog(logFolderPath, "No files found in input folder.");
            }


            string portfolioFile = null;
            string dataUploadFile = null;

            List<string> factSetFiles = new List<string>();

            foreach (var file in files)
            {
                string fileName = Path.GetFileName(file);
                FileInfo fileInfo = new FileInfo(file);

                if (fileInfo.Length == 0)
                {
                    WriteLog(logFolderPath, $"File {fileName} skipped because it is empty.");
                    continue;
                }

                WriteLog(logFolderPath, $"Processing file: {fileName}");
                bool matched = false;


                // data_upload_ & Portfoliobb_ for bloomberg file HANDLING
                if (fileName.StartsWith("Portfoliobb_", StringComparison.OrdinalIgnoreCase))
                {
                    portfolioFile = file;
                    string destFile = Path.Combine(BloombergInputPath, Path.GetFileName(file));

                    using (FileStream sourceStream = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    using (FileStream destStream = new FileStream(destFile, FileMode.Create, FileAccess.Write, FileShare.None))
                    {
                        sourceStream.CopyTo(destStream);
                    }

                    WriteLog(logFolderPath, $"Identified Portfolio Bloomberg file. Copied to {destFile}");
                }
                else if (fileName.StartsWith("data_upload_", StringComparison.OrdinalIgnoreCase))
                {
                    dataUploadFile = file;
                    string destFile = Path.Combine(BloombergInputPath, Path.GetFileName(file));

                    using (FileStream sourceStream = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    using (FileStream destStream = new FileStream(destFile, FileMode.Create, FileAccess.Write, FileShare.None))
                    {
                        sourceStream.CopyTo(destStream);
                    }

                    WriteLog(logFolderPath, $"Identified Data Upload Bloomberg file. Copied to {destFile}");
                }



                // =================== FACTSET FILE HANDLING ===================
                string[] factSetFilePrefixes = new string[]
                {
                    "Callput_", "Dues_", "Indxval_", "Portfolio_",
                    "Price_", "Security_", "Transactions_", "Weightages_"
                };

                foreach (var prefix in factSetFilePrefixes)
                {
                    var fact_file = fileName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase);
                    if (fact_file)
                    {
                        string destFile = Path.Combine(FactSetInputPath, Path.GetFileName(file));

                        using (FileStream sourceStream = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        using (FileStream destStream = new FileStream(destFile, FileMode.Create, FileAccess.Write, FileShare.None))
                        {
                            sourceStream.CopyTo(destStream);
                        }

                        factSetFiles.Add(destFile);
                        WriteLog(logFolderPath, $"Identified FactSet file: {file} copied to {destFile}");
                    }
                }






                foreach (DataRow row in mapping.Rows)
                {
                    string pattern = row["File Name Pattern"].ToString().Trim();
                    string destFolder = row["Folder Path"].ToString();
                    string GNETdestFolder = row["GNET Destination"].ToString();
                    string regexPattern = GenerateRegexFromPattern(pattern);


                    // Match file name OR match parent folder name
                    string folderName = Path.GetFileName(Path.GetDirectoryName(file));

                    if (Regex.IsMatch(fileName, regexPattern, RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(folderName, regexPattern, RegexOptions.IgnoreCase))
                    {
                        matched = true;
                        WriteLog(logFolderPath, $"Pattern matched: {fileName} => {destFolder}");

                        DateTime fileCopyStart = DateTime.Now;
                        bool isUploaded = false;
                        bool isBkp = false;

                        try
                        {



                            // Upload to SFTP
                            if (!fileName.Contains("aum_sms.txt"))
                            {
                                isUploaded = UploadToSftp(sftpHost, sftpPort, sftpUser, sftpPass, file, destFolder, logFolderPath);

                            }

                            // For GNET Destination path 
                            if (GNETdestFolder.Contains(@"\\10.81.112.31\FromMfund"))
                            {
                                string FromMfund = Path.Combine(GNETdestFolder, fileName);

                                try
                                {
                                    System.IO.File.Copy(file, FromMfund, true);
                                    WriteLog(logFolderPath, $"File uploaded successfully on  ..: {FromMfund}");
                                }
                                catch (Exception ex)
                                {
                                    WriteLog(logFolderPath, $"Error uploading file on ..: {ex.Message}");
                                     SendEmail($"Error uploading file on GNET folder {FromMfund} file is {fileName}  . Reason: {ex.Message}", "Error", null);

                                }
                            }





                            if (fileName.Contains("aum_sms.txt"))
                            {
                                isUploaded = true;
                            }

                            if (isUploaded)
                            {
                                try
                                {
                                    Directory.CreateDirectory(backupFolderPath);
                                    string backuppath = Path.Combine(backupFolderPath, fileName);
                                    System.IO.File.Copy(file, backuppath, true);
                                    WriteLog(logFolderPath, $"file sucessfullyy moved on backup folder..: {backuppath}");
                                    isBkp = true;
                                }
                                catch (Exception ex)
                                {
                                    WriteLog(logFolderPath, $"Backup failed for {fileName}: {ex.Message}");
                                    SendEmail($"Error uploading file on Backup folder {backupFolderPath} file is {fileName}  . Reason: {ex.Message}", "Error", null);

                                }
                            }



                            if (fileName.StartsWith("Portfoliobb_", StringComparison.OrdinalIgnoreCase) ||
                                fileName.StartsWith("data_upload_", StringComparison.OrdinalIgnoreCase))
                            {
                                isUploaded = false;
                                isBkp = false;
                            }

                            // For Bloomberg Uplode
                            if (!string.IsNullOrEmpty(portfolioFile) && !string.IsNullOrEmpty(dataUploadFile))
                            {

                                string zipPath = string.Empty;
                                try
                                {

                                    string portfolioFileName = Path.GetFileName(portfolioFile);
                                    string dataUploadFileName = Path.GetFileName(dataUploadFile);

                                    string portfolioFileFullPath = Path.Combine(BloombergInputPath, portfolioFileName);
                                    string dataUploadFileFullPath = Path.Combine(BloombergInputPath, dataUploadFileName);

                                    if (System.IO.File.Exists(portfolioFileFullPath) && System.IO.File.Exists(dataUploadFileFullPath))
                                    {
                                        zipPath = Path.Combine(BloombergInputPath, $"BB_{DateTime.Now:yyyyMMdd}.zip");

                                        using (var zipStream = new FileStream(zipPath, FileMode.Create))
                                        using (var archive = new System.IO.Compression.ZipArchive(zipStream, System.IO.Compression.ZipArchiveMode.Create))
                                        {
                                            archive.CreateEntryFromFile(portfolioFileFullPath, portfolioFileName);
                                            archive.CreateEntryFromFile(dataUploadFileFullPath, dataUploadFileName);
                                        }

                                        WriteLog(logFolderPath, $"ZIP created with Bloomberg files: {zipPath}");

                                        bool isZipUploaded =
                                                UploadToSftp(
                                                    BloombergHost,
                                                    BloombergPort,
                                                    BloombergUserID,
                                                    BloombergPassword,
                                                    zipPath,
                                                    "/",
                                                    logFolderPath
                                                );


                                        if (isZipUploaded)
                                        {
                                            isUploaded = true;
                                            isBkp = true;
                                            WriteLog(logFolderPath, "Bloomberg ZIP uploaded successfully.");

                                            SendEmail(
                                                "Dear Team,<br> Both Bloomberg files (Portfolio & Data Upload) have been zipped and uploaded successfully to Bloomberg SFTP.<br><br>Regards,<br>MPOWER Utility",
                                                "Bloomberg",
                                                zipPath
                                            );

                                            try
                                            {
                                                System.IO.File.Delete(portfolioFileFullPath);
                                                System.IO.File.Delete(dataUploadFileFullPath);
                                                //System.IO.File.Delete(file);
                                                string dataFile = Path.Combine(inFolderPath, dataUploadFileName);
                                                string portfFile = Path.Combine(inFolderPath, portfolioFileName);

                                                Directory.CreateDirectory(successFolderPath);
                                                System.IO.File.Copy(zipPath, Path.Combine(successFolderPath, Path.GetFileName(zipPath)), true);

                                                if (System.IO.File.Exists(dataFile))
                                                {
                                                    //System.IO.File.Copy(dataFile, Path.Combine(successFolderPath, dataUploadFileName), true);
                                                    System.IO.File.Delete(dataFile);
                                                }
                                                if (System.IO.File.Exists(portfFile))
                                                {
                                                    //System.IO.File.Copy(portfFile, Path.Combine(successFolderPath, portfFile), true);
                                                    System.IO.File.Delete(portfFile);
                                                }
                                                //WriteLog(logFolderPath, "Deleted Bloomberg input files after successful upload.");
                                            }
                                            catch (Exception delEx)
                                            {
                                                WriteLog(logFolderPath, $"Warning: Could not delete Bloomberg input files: {delEx.Message}");
                                            }
                                        }
                                        else
                                        {
                                            WriteLog(logFolderPath, "Bloomberg ZIP upload failed.");
                                            //SendEmail("Bloomberg ZIP uploaded failed.", "Error", zipPath);
                                        }
                                    }
                                    //else
                                    //{
                                    //    WriteLog(logFolderPath, "Another Bloomberg files are missing — skipping ZIP creation.");
                                    //}
                                }
                                catch (Exception ex)
                                {
                                    WriteLog(logFolderPath, $"Error creating/uploading Bloomberg ZIP: {ex.Message}");
                                    SendEmail($"Dear Team, <br>Error creating/uploading Bloomberg ZIP.<br>Reason: {ex.Message}", "Bloomberg", zipPath);
                                }
                            }


                            foreach (var prefix in factSetFilePrefixes)
                            {
                                var fact_file = fileName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase);
                                if (fact_file)
                                {

                                    isUploaded = false;
                                    isBkp = false;
                                }
                            }


                            if (factSetFiles.Count == factSetFilePrefixes.Length)
                            {
                                string zipPath = Path.Combine(FactSetBackupPath, $"FactSet_{DateTime.Now:yyyyMMdd}.zip");

                                using (var zipStream = new FileStream(zipPath, FileMode.Create))
                                using (var archive = new System.IO.Compression.ZipArchive(zipStream, System.IO.Compression.ZipArchiveMode.Create))
                                {
                                    foreach (string factfile in factSetFiles)
                                    {
                                        string factfileName = Path.GetFileName(factfile);
                                        string renamedFileName = Regex.Replace(factfileName, "_\\d{8}(?=\\.txt$)", "", RegexOptions.IgnoreCase);
                                        archive.CreateEntryFromFile(factfile, renamedFileName);
                                        WriteLog(logFolderPath, $"Added to FactSet ZIP: {factfileName}");
                                    }
                                }

                                WriteLog(logFolderPath, $"FactSet ZIP created successfully: {zipPath}");

                                // Step 3: Upload ZIP to FactSet SFTP
                                bool isFactSetZipUploaded =
                                                            UploadToSftp(
                                                            FactSetHost,
                                                            FactSetPort,
                                                            FactSetUserID,
                                                            FactSetPassword,
                                                            zipPath,
                                                            "/",
                                                            logFolderPath
                                                            );

                                if (isFactSetZipUploaded)
                                {
                                    WriteLog(logFolderPath, "FactSet ZIP uploaded successfully.");
                                    isUploaded = true;
                                    isBkp = true;
                                    SendEmail(
                                        "All FactSet files have been zipped and uploaded successfully to FactSet SFTP.<br><br>Regards,<br>MPOWER Utility",
                                        "FactSet",
                                        zipPath
                                    );


                                    foreach (string factfile in factSetFiles)
                                    {
                                        try
                                        {
                                            string intempfile = Path.Combine(inFolderPath, Path.GetFileName(factfile));
                                            System.IO.File.Delete(factfile);
                                            System.IO.File.Delete(intempfile);

                                        }
                                        catch (Exception delEx)
                                        {
                                            WriteLog(logFolderPath, $"Warning: Could not delete FactSet input file {file}: {delEx.Message}");
                                        }
                                    }

                                    Directory.CreateDirectory(successFolderPath);
                                    System.IO.File.Copy(zipPath, Path.Combine(successFolderPath, Path.GetFileName(zipPath)), true);



                                }
                                else
                                {
                                    WriteLog(logFolderPath, "FactSet ZIP upload failed.");
                                }
                            }


                            if (isUploaded && isBkp)
                            {

                                if (System.IO.File.Exists(file))
                                {
                                    Directory.CreateDirectory(successFolderPath);
                                    System.IO.File.Copy(file, Path.Combine(successFolderPath, fileName), true);
                                    WriteLog(logFolderPath, $"File copied to SUCCESS folder: {successFolderPath}");

                                    try
                                    {
                                        System.IO.File.Delete(file);
                                        WriteLog(logFolderPath, $"File Moved from the IN folder: {file}");
                                    }

                                    catch (Exception ex)
                                    {
                                        WriteLog(logFolderPath, $"Failed to Moved {file} from IN folder: {ex.Message}");
                                    }
                                }
                            }
                            else
                            {
                                Directory.CreateDirectory(errorFolderPath);
                                System.IO.File.Copy(file, Path.Combine(errorFolderPath, fileName), true);
                                WriteLog(logFolderPath, $"File copied to ERROR folder: {errorFolderPath}");
                            }
                        }
                        catch (Exception ex)
                        {


                            Directory.CreateDirectory(errorFolderPath);
                            System.IO.File.Copy(file, Path.Combine(errorFolderPath, fileName), true);
                            WriteLog(logFolderPath, $"Exception while uploading {fileName}: {ex.Message}. Copied to ERROR folder.");
                        }

                        DateTime fileCopyEnd = DateTime.Now;
                        WriteLog(logFolderPath, $"Completed processing {fileName} in {(fileCopyEnd - fileCopyStart).TotalSeconds} sec");
                        break;
                    }


                }

                if (!matched)
                {
                    WriteLog(logFolderPath, $"No pattern matched for file: {fileName}. File skipped.");
                }
            }


            if (DateTime.Now.Hour == 0 && DateTime.Now.Minute >= 50 && DateTime.Now.Minute <= 55)
            {
                var checkedFile = Directory.GetFiles(inFolderPath);
                if (checkedFile.Length > 0)
                {
                    string yestFolder = Path.Combine(inBkpFolderPath, DateTime.Now.AddDays(-1).ToString("yyyyMMdd"));
                    Directory.CreateDirectory(yestFolder);
                    try
                    {
                        foreach (var f in checkedFile)
                        {
                            string dest = Path.Combine(yestFolder, Path.GetFileName(f));
                            using (FileStream src = new FileStream(f, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                            using (FileStream dst = new FileStream(dest, FileMode.Create, FileAccess.Write))
                                src.CopyTo(dst);
                            System.IO.File.Delete(f);
                            WriteLog(logFolderPath, $"Copied & deleted leftover file: {Path.GetFileName(f)}");
                        }
                    }
                    catch (Exception e)
                    {
                        WriteLog(logFolderPath, $"In Folder Cleanup : {e.ToString()}");
                    }
                }
            }

            WriteLog(logFolderPath, $"Utility completed successfully at {DateTime.Now:dd-MMM-yy HH:mm:ss}\n");
        }
        catch (Exception ex)
        {
            WriteLog(logFolderPath, $"Unexpected error: {ex.Message}");
        }
    }

    private bool UploadToSftp(string host, int port, string username, string password, string localFile, string remoteFolder, string logPath)
    {
        string fileName = Path.GetFileName(localFile);
        WriteLog(logPath, $"File Name: {fileName}");
        WriteLog(logPath, $"Trying to connect to SFTP: {host}:{port}");

        using (var sftp = new SftpClient(host, port, username, password))
        {
            try
            {
                sftp.Connect();
                WriteLog(logPath, $"Connected to SFTP server: {host}:{port}");
            }
            catch (Exception ex)
            {
                WriteLog(logPath, $"Unable to connect to {host}:{port}. Reason: {ex.Message}");
                SendEmail($"Dear Team,<br>Unable to connect SFTP {host}:{port}. Reason: {ex.Message}<br><br>", "Error", null);
                return false;
            }

            try
            {
                //remoteFolder = remoteFolder.Replace("\\", "/");

                //if (remoteFolder.StartsWith("//"))
                //{
                //    var parts = remoteFolder.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
                //    if (parts.Length > 1)
                //    {
                //        remoteFolder = "/" + string.Join("/", parts.Skip(1));
                //    }
                //}

                //if (!remoteFolder.StartsWith("/"))
                //    remoteFolder = "/" + remoteFolder;


                //if (!sftp.Exists(remoteFolder))
                //{
                //    try
                //    {
                //        sftp.CreateDirectory(remoteFolder);
                //        WriteLog(logPath, $"Created remote folder: {remoteFolder}");
                //    }
                //    catch (Exception ex)
                //    {
                //        WriteLog(logPath, $"Unable to create remote folder {remoteFolder}. Reason: {ex.Message}");
                //        return false;
                //    }
                //}

                //string remoteFilePath = remoteFolder.TrimEnd('/') + "/" + fileName;

                remoteFolder = NormalizeRemotePath(remoteFolder);
                string remoteFilePath = remoteFolder.TrimEnd('/') + "/" + fileName;


                WriteLog(logPath, $"File : {fileName} : uploaded Location : {host}//{remoteFilePath}");
                Console.WriteLine($"File : {fileName} : uploaded Location : {host}//{remoteFilePath}");
                byte[] fileBytes = System.IO.File.ReadAllBytes(localFile);

                using (var memoryStream = new MemoryStream(fileBytes))
                {
                    sftp.UploadFile(memoryStream, remoteFilePath, true);
                }

                WriteLog(logPath, $"File uploaded successfully: {fileName}");
                return true;
            }
            catch (Exception ex)
            {
                WriteLog(logPath, $"Failed to upload file : '{fileName}' to destination :  {remoteFolder}. Reason: {ex}");
                if (fileName.StartsWith("BB_", StringComparison.OrdinalIgnoreCase))
                {
                    SendEmail($"Dear Team,<br>Failed to upload file, file name : {fileName} to destination :  {host}//{remoteFolder} : Reason" + ex.Message, "Bloomberg", localFile);
                }

                else if (fileName.StartsWith("FactSet_", StringComparison.OrdinalIgnoreCase))
                {
                    SendEmail($"Dear Team,<br>Failed to upload file,<br> file name : {fileName} to destination :  {host}{remoteFolder} : Reason" + ex.Message + "<br><br>", "FactSet", localFile);
                }
                else
                {
                    SendEmail($"Dear Team,<br>Failed to upload file, file name : {fileName} to destination :  {host}//{remoteFolder} : Reason" + ex.Message, "Error", null);
                }
                return false;
            }
            finally
            {
                if (sftp.IsConnected)
                {
                    sftp.Disconnect();
                    WriteLog(logPath, $"Disconnected from SFTP server: {host}:{port}");
                }
            }
        }
    }






    static DataTable LoadMapping(string excelPath, string logPath)
    {
        try
        {
            WriteLog(logPath, $"Opening mapping file: {excelPath}");
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using var stream = System.IO.File.Open(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = ExcelReaderFactory.CreateReader(stream);

            WriteLog(logPath, "Excel reader created successfully.");
            var result = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            });

            if (result.Tables.Count == 0 || result.Tables[0].Rows.Count == 0)
            {
                WriteLog(logPath, "Mapping file is empty or contains no valid rows.");
                throw new Exception("Mapping file empty.");
            }

            var table = result.Tables[0];
            if (!table.Columns.Contains("File Name Pattern") || !table.Columns.Contains("Folder Path"))
            {
                WriteLog(logPath, "Required columns missing in mapping file.");
                throw new Exception("Invalid mapping file format.");
            }

            /*foreach (DataRow row in table.Rows)
            {
                WriteLog(logPath, $"Loaded Mapping - Pattern: '{row["File Name Pattern"]}', Folder: '{row["Folder Path"]}' GNET Destination {row["GNET Destination"]}");
            }*/

            WriteLog(logPath, $"Mapping file loaded successfully. Total rows: {table.Rows.Count}");
            return table;


        }
        catch (Exception ex)
        {
            WriteLog(logPath, $"Mapping load failed: {ex.Message}");
            throw;
        }
    }



    private List<string> GetAllFilesToProcess(string inFolderPath, DataTable mapping, string logFolderPath)
    {
        List<string> finalFiles = new List<string>();

        try
        {
            var rootFiles = Directory.GetFiles(inFolderPath, "*.*", SearchOption.TopDirectoryOnly);
            finalFiles.AddRange(rootFiles);

            var subFolders = Directory.GetDirectories(inFolderPath);

            foreach (var folder in subFolders)
            {
                string folderName = Path.GetFileName(folder);

                foreach (DataRow row in mapping.Rows)
                {
                    string pattern = row["File Name Pattern"].ToString();
                    string regexPattern = GenerateRegexFromPattern(pattern);

                    if (Regex.IsMatch(folderName, regexPattern, RegexOptions.IgnoreCase))
                    {
                        WriteLog(logFolderPath, $"Folder matched with pattern: {folderName}");
                        var filesInFolder = Directory.GetFiles(folder, "*.*", SearchOption.AllDirectories);
                        finalFiles.AddRange(filesInFolder);
                    }
                }
            }

            WriteLog(logFolderPath, $"Total files fetched for processing: {finalFiles.Count}");
        }
        catch (Exception ex)
        {
            WriteLog(logFolderPath, $"Error fetching files: {ex.Message}");
        }

        return finalFiles;
    }


    private string GenerateRegexFromPattern(string pattern)
    {
        if (string.IsNullOrWhiteSpace(pattern))
            return string.Empty;

        pattern = pattern.Trim();

        // Escape special regex chars
        string regex = Regex.Escape(pattern);

        regex = Regex.Replace(regex, "yyyymmdd", @"\d{8}", RegexOptions.IgnoreCase);
        regex = Regex.Replace(regex, "ddmmyyyy", @"\d{8}", RegexOptions.IgnoreCase);
        regex = Regex.Replace(regex, "ddmmyy", @"\d{6}", RegexOptions.IgnoreCase);
        regex = Regex.Replace(regex, "ddmmmyyyy", @"\d{2}[A-Za-z]{3}\d{4}", RegexOptions.IgnoreCase);
        regex = Regex.Replace(regex, "ddMonyyyy", @"\d{2}[A-Za-z]{3}\d{4}", RegexOptions.IgnoreCase);
        regex = Regex.Replace(regex, "dd-Mon-yyyy", @"\d{2}-[A-Za-z]{3}-\d{4}", RegexOptions.IgnoreCase);
        regex = Regex.Replace(regex, "dd-MMM-yy", @"\d{2}-[A-Za-z]{3}-\d{2}", RegexOptions.IgnoreCase);
        regex = Regex.Replace(regex, "dd-MMM-yyyy", @"\d{2}-[A-Za-z]{3}-\d{4}", RegexOptions.IgnoreCase);

        regex = Regex.Replace(regex, @"dd\+1mmyyyy", @"\d{8}", RegexOptions.IgnoreCase);


        regex = Regex.Replace(regex, @"dd-Mon-yyyy-dd-Mon-yyyy\.xls",
        @"$1-\d{2}-[A-Za-z]{3}-\d{4}-\d{2}-[A-Za-z]{3}-\d{4}\.xls",
        RegexOptions.IgnoreCase);

        regex = regex.Replace(@"\ ", @"[ _]+");

        if (!regex.Contains(@"\."))
        {
            regex += @"(\.[A-Za-z0-9]+)?";
        }

        return "^" + regex + "$";
    }



    private string NormalizeRemotePath(string remoteFolder)
    {
        if (string.IsNullOrWhiteSpace(remoteFolder))
            return "/";

        // Replace backslashes with forward slashes
        remoteFolder = remoteFolder.Replace("\\", "/");

        // UNC path se server part hatao
        if (remoteFolder.StartsWith("//"))
        {
            var parts = remoteFolder.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length > 1)
            {
                // remove server (first part)
                remoteFolder = "/" + string.Join("/", parts.Skip(1));
            }
            else
            {
                remoteFolder = "/";
            }
        }

        // Ensure starts with "/"
        if (!remoteFolder.StartsWith("/"))
            remoteFolder = "/" + remoteFolder;

        return remoteFolder;
    }


    static void WriteLog(string logFolder, string message)
    {
        if (!Directory.Exists(logFolder))
            Directory.CreateDirectory(logFolder);

        string logFilePath = Path.Combine(logFolder, "FilePickerLog_" + DateTime.Now.ToString("dd-MMM-yyyy") + ".txt");

        using (StreamWriter sw = new StreamWriter(logFilePath, true))
        {
            sw.WriteLine($"{DateTime.Now:dd-MMM-yy HH:mm:ss}\t{message}");
        }
    }


    public static bool SendEmail(string error_msg, string type, string attachmentPath)

    {
        string logFolderPath = ConfigurationManager.AppSettings["ErrorLogFilePath"];
        bool is_success = true;

        string strBody = string.Empty;

        string strFirstLine = string.Empty;

        string strSubject = string.Empty;

        string strEmail_body = string.Empty;

        string strEmpEmail = String.Empty;

        string strEmpIDs = String.Empty;

        try

        {

            string SmtpPrimaryMailServer = ConfigurationManager.AppSettings["SmtpPrimaryMailServer"];

            string SENDER_EMAIL = ConfigurationManager.AppSettings["SenderEmail"];

            string SENDER_PASS = ConfigurationManager.AppSettings["SenderPass"];

            int PORT = Convert.ToInt32(ConfigurationManager.AppSettings["Port"]);

            bool IsSSL = Convert.ToBoolean(ConfigurationManager.AppSettings["IsSsl"]);

            string[] ToEmail = ConfigurationManager.AppSettings["ToEmail"].Split(',');

            bool is_email_stop = Convert.ToBoolean(ConfigurationManager.AppSettings["Is_Email_Stop"]);



            if (is_email_stop)

            {

                return is_success;

            }

            SmtpClient smtpClient;

            smtpClient = new SmtpClient(SmtpPrimaryMailServer, PORT);

            MailMessage message = new MailMessage();

            message.IsBodyHtml = true;

            MailAddress mailAddress = new MailAddress(SENDER_EMAIL);

            message.Sender = mailAddress;

            message.From = mailAddress;

            strEmail_body = error_msg;

            strSubject = "MPower SFTP Utility - SFTP UPLOAD ERROR > " + DateTime.Now.ToString("dd-MMM-yyyy");

            if (type.Equals("Bloomberg", StringComparison.OrdinalIgnoreCase))
            {
                ToEmail = ConfigurationManager.AppSettings["ToBloombergEmail"].Split(',');
                strSubject = "Bloomberg Files Uploaded Successfully - " + DateTime.Now.ToString("dd-MMM-yyyy");
                SENDER_EMAIL = "Bloomberg_SFTP_ADMIN@uti.co.in";

            }

            if (type.Equals("FactSet", StringComparison.OrdinalIgnoreCase))
            {
                ToEmail = "server.admin@uti.co.in".Split(',');
                strSubject = "FactSet Files Uploaded Status - " + DateTime.Now.ToString("dd-MMM-yyyy");
                SENDER_EMAIL = "FACTSET_SFTP_ADMIN@uti.co.in";

            }
            // ToEmail = "Bhaneshvar.Kshirsagar@cylsys.com".Split(',');
            string strDetailBody = String.Empty;

            string strDetails = String.Empty;

            DataSet ds = new DataSet();

            strFirstLine = "Dear Sir/Madam,";


            strDetails += "<TR>" +

                         "<TD COLSPAN=4 ><B>" + strFirstLine + "</B></TD>" +

                           "</TR>";

            strDetails += "<TR>" +

                           "<TD COLSPAN=4 ></TD>" +

                          "</TR>";

            strDetails += "<TR>" +

                            "<TD COLSPAN=4 ><B>" + strSubject + "</B></TD>" +

                           "</TR>";


            strBody = "  <html><body><table border=0 cellpadding=0 ALIGN=LEFT WIDTH=\"100%\";\">" +

                          strEmail_body +

                        "<tr><td></table></body></html>";


            foreach (string str_email in ToEmail)
            {

                if (str_email.Contains("@"))

                {

                    message.To.Add(str_email);

                }

            }
            if (System.IO.File.Exists(attachmentPath))
            {
                message.Attachments.Add(new Attachment(attachmentPath));
            }

            message.Body = strBody;

            message.Subject = strSubject;

            WriteLog(logFolderPath, "Trying To Connect SMTP");
            smtpClient.Send(message);


            WriteLog(logFolderPath, "Mail Sent successfully");


            smtpClient.Dispose();
        }

        catch (Exception ex)

        {

            is_success = false;

            WriteLog(logFolderPath, "Sending Mail Failed : " + ex);

        }

        return is_success;

    }





}
