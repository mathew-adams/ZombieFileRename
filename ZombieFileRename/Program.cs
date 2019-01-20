using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NDesk.Options;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.WindowsAPICodePack.Shell;
using Microsoft.WindowsAPICodePack.Shell.PropertySystem;

namespace ZombieFileRename
{
    class Program
    {
        private static bool log;
        private static bool Log
        {
            get { return log; }
            set { log = value; }
        }
        private static string logPath;
        private static string LogPath
        {
            get { return logPath; }
            set { logPath = value; }
        }

        public static readonly string[] keywordremove = { "720p", "720i", "1080p", "1080i", "2160p", "2160i" };

        static void Main(string[] args)
        {

            string directory = "";
            string logFile = "";
            string moveAfterRename = "";
            string newDirectoryName = "";
            string newDirectoryFull = "";
            string friendlyName = "";
            bool directoryRenamed = false;
            bool fileRenamed = false;
            bool hasMetaData = false;
            bool metaDataRenamed = false;

            StringBuilder sbLog = new StringBuilder();

            //Get parameters
            OptionSet options = new OptionSet
            {
                //"=" is for required
                //":" is for optional
                {"d=|directory", d =>{directory = d;}},
                {"lf:|log", lf =>{ logFile = lf;}},
                {"mar:|moveafterrename", mar =>{ moveAfterRename = mar;}}
            };

            try
            {
                options.Parse(args);
            }
            catch(Exception ex)
            {
                if (!String.IsNullOrEmpty(logFile))
                {
                    Log = true;
                    logPath = logFile;
                }

                if (Log)
                {
                    sbLog.Append(ex);
                    sbLog.AppendLine(Environment.NewLine);
                    WriteLog(sbLog);
                }
                Environment.Exit(-200);
            }
            
            if (!String.IsNullOrEmpty(logFile))
            {
                Log = true;
                logPath = logFile;
            }

            if (Log)
            {
                sbLog.AppendLine("------------------------------------------------------------------------------------------------------");
                sbLog.AppendLine($"Command:.............. {Environment.CommandLine}");
                sbLog.AppendLine($"Timestamp:............ {DateTime.Now.ToString()}");
                sbLog.AppendLine($"Directory:............ {directory}");
                sbLog.AppendLine($"logfile:.............. {logFile}");
                sbLog.AppendLine($"moveafterrename:...... {moveAfterRename}");
            }

            DirectoryInfo dirInfo = new DirectoryInfo(directory);
            //Get directory files
            newDirectoryName = GetSanitizedName(dirInfo.Name, out friendlyName);
            newDirectoryFull = Path.Combine(dirInfo.Parent.FullName, newDirectoryName);

            if (Directory.Exists(directory))
            {
                try
                {
                    if(Log)
                        sbLog.AppendLine("Movie folder rename begin...." + directory);

                    //Try directory move = directory rename
                    Directory.Move(directory, newDirectoryFull);
                    directoryRenamed = true;

                    if (Log)
                        sbLog.AppendLine("Movie folder rename success...." + newDirectoryFull);

                    //TODO: Rename all files that match the director name??
                    var newDirectoryInfo = new DirectoryInfo(newDirectoryFull);

                    foreach(FileInfo file in newDirectoryInfo.GetFiles())
                    {
                        //File name matches the unsanitized directoryname
                        if(file.Name.Replace(file.Extension,"").Equals(dirInfo.Name, StringComparison.CurrentCultureIgnoreCase))
                        {
                            if (Log)
                                sbLog.AppendLine("File found to match directory...." + file.FullName);

                            //Rename file with move operation
                            string fileNew = Path.Combine(newDirectoryFull, newDirectoryName) + file.Extension;
                            file.MoveTo(fileNew);
                            fileRenamed = true;

                            if (Log)
                                sbLog.AppendLine("File renamed success...." + fileNew);

                            //Change all extended properties to match filename (metadata: title, comment)
                            try
                            {
                                if (Log)
                                    sbLog.AppendLine("Try and update metadata...." + fileNew);
                                var fileMeta = ShellFile.FromFilePath(fileNew);
                                hasMetaData = !string.IsNullOrEmpty(fileMeta.Properties.System.Title.Value);
                                fileMeta.Properties.System.Title.Value = friendlyName;
                                fileMeta.Properties.System.Comment.Value = friendlyName;
                                metaDataRenamed = true;
                                if (Log)
                                    sbLog.AppendLine("Update metadata success...." + friendlyName);
                            }
                            catch(Exception ex)
                            {
                                //File metadata could not be updated
                                hasMetaData = false;
                                metaDataRenamed = false;
                                if (Log)
                                {
                                    sbLog.Append(ex);
                                    sbLog.AppendLine(Environment.NewLine);
                                }
                            }
                        }
                        else
                        {
                            if (Log)
                                sbLog.AppendLine("File did not match directory name...." + file.Name);
                        }
                    }

                    //Finally, move to destination drive if everything is expected
                    if (!String.IsNullOrEmpty(moveAfterRename))
                    {
                        if (Log)
                            sbLog.AppendLine("File rename complete, move to final directory.....");

                        string moveAfterRenameFile = Path.Combine(moveAfterRename, newDirectoryInfo.Name);
                        if (directoryRenamed && fileRenamed && hasMetaData && metaDataRenamed)
                        {
                            if (Log)
                                sbLog.AppendLine("Directory, file, and metadata renamed. Move to drive...." + moveAfterRenameFile);
                            TryMoveDirectory(Log, sbLog, newDirectoryFull, moveAfterRenameFile);
                        }
                        else if (directoryRenamed && fileRenamed && !hasMetaData)
                        {
                            if (Log)
                                sbLog.AppendLine("Directory, file renamed, no metadata. Move to driver...." + moveAfterRenameFile);
                            TryMoveDirectory(Log, sbLog, newDirectoryFull, moveAfterRenameFile);
                        }
                        else
                        {
                            if (Log)
                                sbLog.AppendLine("Conditions not met to move to driver...." + moveAfterRenameFile);
                        }
                    }
                    
                }
                catch(Exception ex)
                {
                    if (Log)
                    {
                        sbLog.Append(ex);
                        sbLog.AppendLine(Environment.NewLine);
                        WriteLog(sbLog);
                    }
                    Environment.Exit(-300);
                }
            }
            else
            {
                //Directory does't exist, log it and move on
                if (Log)
                    sbLog.AppendLine("Directory does not exist, nothing to rename.... " + directory);
            }

            //Write log
            if (Log && sbLog.Length > 0)
            {
                WriteLog(sbLog);
            }
        }

        #region "Private Functions"
        private static void WriteLog(StringBuilder sb)
        {

            //TODO: only write logs if the length > 0
            var logFile = new FileInfo(LogPath);

            if (!logFile.Directory.Exists)
            {
                logFile.Directory.Create();
            }
            using (StreamWriter swrt = new StreamWriter(LogPath, true))
            {
                swrt.Write(sb.ToString());
            }
        }

        private static void TryMoveDirectory(bool log, StringBuilder sbLog, string source, string dest)
        {
            if (log)
                sbLog.AppendLine("Try move directory to...." + dest);

            try
            {
                if (!Directory.Exists(dest))
                    Directory.CreateDirectory(dest);
                string[] files = Directory.GetFiles(source);
                foreach (string file in files)
                {
                    string name = Path.GetFileName(file);
                    string destfile = Path.Combine(dest, name);

                    if (Log)
                        sbLog.AppendLine("Copy file start:...." + destfile);
                    File.Copy(file, destfile);
                    if(Log)
                        sbLog.AppendLine("Copy file success:...." + destfile);

                    if (Log)
                        sbLog.AppendLine("Delete file start:...." + file);
                    File.Delete(file);
                    if (Log)
                        sbLog.AppendLine("Delete file success:...." + file);
                }
                string[] folders = Directory.GetDirectories(source);
                foreach (string folder in folders)
                {
                    string name = Path.GetFileName(folder);
                    string destFolder = Path.Combine(dest, name);

                    if (Log)
                        sbLog.AppendLine("Move directory start:...." + destFolder);
                    TryMoveDirectory(log, sbLog, folder, destFolder);
                    if (Log)
                        sbLog.AppendLine("Move directory success:...." + destFolder);
                }

                //All done, delete source
                if(Log)
                    sbLog.AppendLine("Delete source directory begin:...." + source);
                Directory.Delete(source);
                if (Log)
                    sbLog.AppendLine("Delete source directory success....." + source);
             
            }
            catch(Exception ex)
            {
                if(Log)
                {
                    sbLog.Append(ex);
                    sbLog.AppendLine(Environment.NewLine);
                }
            }
        }

        private static string GetSanitizedName(string name, out string friendlyName)
        {
            //Replace hypens and periods with spaces
            string newName = name.Replace("."," ").Replace("-"," ");
            
            //Remove known keywords, 720p, 1980p, etc. so we can determine the year correctly
            foreach(string keyword in keywordremove)
            {
                newName = newName.Replace(keyword, "");
            }

            //Find year, four numeric contiguous digits
            Regex regYear = new Regex(@"[(\d)]{4}", RegexOptions.Compiled & RegexOptions.IgnoreCase);
            MatchCollection year = regYear.Matches(newName);

            //Find last year in the file name, this is so we can parse something like "2001: A space odyssey"
            Match lastValue = year.Cast<Match>().Last<Match>();

            friendlyName = newName; //Assign value so the compiler is happy
            if (lastValue.Success)
            {
                newName = newName.Remove(lastValue.Index + lastValue.Length); //Remove anything after the year in the title
                friendlyName = newName.Replace(lastValue.Value, ""); //Remove the year from the title for a friendly display name
                newName = newName.Replace(lastValue.Value, $"({lastValue.Value})"); //Add parens around the year, plex format
            }
            friendlyName = friendlyName.TrimEnd().TrimStart();
            return newName;
        }

        #endregion
    }
}
