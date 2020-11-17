using Cassia;
using IWshRuntimeLibrary;
using LeanWork.IO.FileSystem;
using menelabs.core;
using Microsoft.Win32;
using SuncatCommon;
using SuncatService.Monitors.SmartFileSystemWatcher;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace SuncatService.Monitors
{
    public static class FileSystemActivityMonitor
    {
        private static readonly string serviceName = new ProjectInstaller().ServiceInstaller.ServiceName;
        private static readonly string rootDrive = Path.GetPathRoot(Environment.SystemDirectory);
        private static readonly string serviceAppData = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), serviceName);
        private static readonly IEnumerable<string> fileTypeKeyNames;
        private static VolumeWatcher volumeWatcher;
        private static Dictionary<string, RecoveringFileSystemWatcher> fileSystemWatchers;
        private static Thread recentFilesChecker;

        public static event TrackEventHandler Track;

        static FileSystemActivityMonitor()
        {
            try
            {
                fileTypeKeyNames = Registry.ClassesRoot.GetSubKeyNames().Where(k => k.IndexOf(".") == 0);
            }
            catch (Exception ex)
            {
                #if DEBUG
                    Debug.WriteLine(ex);
                    
                    if (ex.InnerException != null)
                        Debug.WriteLine(ex.InnerException);
                #else
                    Trace.WriteLine(ex);

                    if (ex.InnerException != null)
                        Trace.WriteLine(ex.InnerException);
                #endif
            }
        }

        private static bool IsAssociatedFileType(string file)
        {
            try
            {
                if (fileTypeKeyNames != null)
                {
                    return fileTypeKeyNames.Any(k => file.EndsWith(k));
                }
            }
            catch (Exception ex)
            {
                #if DEBUG
                    Debug.WriteLine(ex);
                    
                    if (ex.InnerException != null)
                        Debug.WriteLine(ex.InnerException);
                #else
                    Trace.WriteLine(ex);

                    if (ex.InnerException != null)
                        Trace.WriteLine(ex.InnerException);
                #endif
            }

            return false;
        }

        public static bool IgnoreEventCallback(SuncatLog log)
        {
            var ignored = false;

            Func<SuncatLogEvent, string, string, bool> ignoreFilePatterns = delegate (SuncatLogEvent logEvent, string path, string oldPath)
            {
                if (path != null && path.IndexOfAny(System.IO.Path.GetInvalidPathChars()) >= 0)
                {
                    return false;
                }

                if (oldPath != null && oldPath.IndexOfAny(System.IO.Path.GetInvalidPathChars()) >= 0)
                {
                    return false;
                }

                var manager = new TerminalServicesManager();
                
                ignored |= Regex.IsMatch(path, $@"^{rootDrive}Users\[^\]+\AppData\".Replace(@"\", @"\\"), RegexOptions.IgnoreCase);
                ignored |= path.StartsWith(Environment.GetFolderPath(Environment.SpecialFolder.Windows), StringComparison.OrdinalIgnoreCase);
                ignored |= path.StartsWith(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), StringComparison.OrdinalIgnoreCase);
                ignored |= path.IndexOf("$RECYCLE.BIN", StringComparison.OrdinalIgnoreCase) > -1;
                ignored |= path.IndexOf("$WINDOWS", StringComparison.OrdinalIgnoreCase) > -1;
                ignored |= path.IndexOf("Config.Msi", StringComparison.OrdinalIgnoreCase) > -1;
                ignored |= path.IndexOf("System Volume Information", StringComparison.OrdinalIgnoreCase) > -1;
                ignored |= path.IndexOf("WindowsApps", StringComparison.OrdinalIgnoreCase) > -1;
                ignored |= path.IndexOf("SystemApps", StringComparison.OrdinalIgnoreCase) > -1;
                ignored |= path.EndsWith("desktop.ini", StringComparison.OrdinalIgnoreCase);
                ignored |= (Path.GetFileName(path).StartsWith("~") && logEvent != SuncatLogEvent.RenameFile);
                ignored |= (Path.GetFileName(path).EndsWith("~") && logEvent != SuncatLogEvent.RenameFile);
                ignored |= (Path.GetExtension(path).Equals(".tmp", StringComparison.OrdinalIgnoreCase) && logEvent != SuncatLogEvent.RenameFile);
                ignored |= Path.GetExtension(path).Equals(".lnk", StringComparison.OrdinalIgnoreCase);
                ignored |= (manager != null && manager.ActiveConsoleSession != null && !string.IsNullOrEmpty(manager.ActiveConsoleSession.UserName)
                            && path.StartsWith($@"{rootDrive}Users\") && !path.StartsWith($@"{rootDrive}Users\{manager.ActiveConsoleSession.UserName}\"));
                ignored |= (path.StartsWith(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), StringComparison.OrdinalIgnoreCase)
                            && !Path.GetExtension(path).Equals(".exe", StringComparison.OrdinalIgnoreCase));
                ignored |= (path.StartsWith(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), StringComparison.OrdinalIgnoreCase)
                            && !Path.GetExtension(path).Equals(".exe", StringComparison.OrdinalIgnoreCase));
                ignored |= path.Contains(@"\.");
                
                return ignored;
            };

            if (log.Data1 != null)
            {
                ignored |= ignoreFilePatterns(log.Event, log.Data1, null);
            }

            if (log.Data2 != null)
            {
                ignored |= ignoreFilePatterns(log.Event, log.Data2, null);
            }

            if (log.Event == SuncatLogEvent.RenameFile)
            {
                ignored |= ignoreFilePatterns(log.Event, log.Data1, log.Data2);
            }

            return ignored;
        }

        private static async void CheckUserRecentFiles()
        {
            DateTime? firstRecentFileDate = null, lastRecentFileDate = null;
            var newRecentFiles = new List<string>();
            var robocopyProcessStartInfo = new ProcessStartInfo();

            while (true)
            {
                try
                {
                    var manager = new TerminalServicesManager();

                    if (manager != null && manager.ActiveConsoleSession != null && !string.IsNullOrEmpty(manager.ActiveConsoleSession.UserName))
                    {
                        var recentFolder = Path.Combine(serviceAppData, "Recent");

                        using (var process = Process.Start(new ProcessStartInfo()
                        {
                            FileName = "robocopy",
                            Arguments = $"\"{rootDrive}Users\\{manager.ActiveConsoleSession.UserName}\\AppData\\Roaming\\Microsoft\\Windows\\Recent\" \"{recentFolder}\" /XF \"desktop.ini\" /MAX:1000000 /A-:SH /PURGE /NP",
                            WindowStyle = ProcessWindowStyle.Hidden,
                            CreateNoWindow = true,
                            UseShellExecute = false,
                        }))
                        {
                            process.WaitForExit();
                        }

                        newRecentFiles.Clear();

                        // don't use EnumerateFiles here as it is unreliable for us in the foreach, GetFiles is reliable (i think EnumerateFiles doesn't fetch all FileInfo data that we need at the time of execution)
                        var recentFiles = new DirectoryInfo(recentFolder).GetFiles("*.lnk").OrderByDescending(f => f.LastAccessTime);

                        foreach (var recentFile in recentFiles)
                        {
                            try
                            {
                                var recentFileDate = recentFile.LastAccessTime;

                                if (lastRecentFileDate.HasValue)
                                {
                                    if (recentFileDate >= lastRecentFileDate)
                                    {
                                        var shell = new WshShell();
                                        var link = (IWshShortcut)shell.CreateShortcut(recentFile.FullName);

                                        if (newRecentFiles.Count == 0)
                                        {
                                            firstRecentFileDate = recentFileDate;
                                        }

                                        link.TargetPath = link.TargetPath.Replace(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), $@"{rootDrive}Users\{manager.ActiveConsoleSession.UserName}");

                                        if (recentFileDate > lastRecentFileDate && System.IO.File.Exists(link.TargetPath))
                                        {
                                            newRecentFiles.Insert(0, link.TargetPath);
                                        }
                                    }
                                    else
                                    {
                                        lastRecentFileDate = firstRecentFileDate;
                                        break;
                                    }
                                }
                                else
                                {
                                    lastRecentFileDate = recentFileDate;
                                    break;
                                }
                            }
                            catch (Exception ex)
                            {
                                #if DEBUG
                                    Debug.WriteLine(ex);
                                    
                                    if (ex.InnerException != null)
                                        Debug.WriteLine(ex.InnerException);
                                #else
                                    Trace.WriteLine(ex);

                                    if (ex.InnerException != null)
                                        Trace.WriteLine(ex.InnerException);
                                #endif
                            }
                        }

                        foreach (var recentFile in newRecentFiles)
                        {
                            var te = new TrackEventArgs();
                            var log = new SuncatLog();

                            log.DateTime = DateTime.Now;
                            log.Event = SuncatLogEvent.OpenFile;
                            log.Data1 = recentFile;

                            try
                            {
                                log.Data3 = volumeWatcher.DriveList[recentFile.Substring(0, 1)].DriveType.ToString();
                            }
                            catch (Exception ex)
                            {
                                #if DEBUG
                                    Debug.WriteLine(ex);
                                    
                                    if (ex.InnerException != null)
                                        Debug.WriteLine(ex.InnerException);
                                #else
                                    Trace.WriteLine(ex);

                                    if (ex.InnerException != null)
                                        Trace.WriteLine(ex.InnerException);
                                #endif

                                log.Data3 = recentFile.StartsWith(@"\\") ? DriveType.Network.ToString() : DriveType.Unknown.ToString();
                            }

                            te.LogData = log;

                            // Make the event wait a bit to let the Open event appear after a Create event if necessary.
                            await Task.Delay(3000).ContinueWith(_ =>
                            {
                                if (!IgnoreEventCallback(log))
                                {
                                    Track?.Invoke(null, te);
                                }
                            });
                        }

                        if (newRecentFiles.Count != 0)
                        {
                            lastRecentFileDate = firstRecentFileDate;
                        }
                    }
                }
                catch (Exception ex)
                {
                    #if DEBUG
                        Debug.WriteLine(ex);
                        
                        if (ex.InnerException != null)
                            Debug.WriteLine(ex.InnerException);
                    #else
                        Trace.WriteLine(ex);

                        if (ex.InnerException != null)
                            Trace.WriteLine(ex.InnerException);
                    #endif
                }

                Thread.Sleep(1000);
            }
        }

        private static void StartFileSystemWatcher()
        {
            fileSystemWatchers = new Dictionary<string, RecoveringFileSystemWatcher>();

            for (var i = 0; i < 26; i++)
            {
                var driveLetter = (char)(i + 'A');

                // Source: https://petermeinl.wordpress.com/2015/05/18/tamed-filesystemwatcher/
                var watcher = new RecoveringFileSystemWatcher();
                watcher.IncludeSubdirectories = true;
                watcher.OrderByOldestFirst = true;
                // To enable LastAccess filter: fsutil behavior set DisableLastAccess 0 (and reboot?)
                watcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.DirectoryName | NotifyFilters.LastWrite | NotifyFilters.CreationTime | NotifyFilters.Size;// | NotifyFilters.LastAccess | NotifyFilters.Security | NotifyFilters.Attributes;
                watcher.SetIgnoreEventCallback<SuncatLog>(IgnoreEventCallback);

                watcher.Created += FileSystemEventHandler;
                watcher.Changed += FileSystemEventHandler;
                watcher.Deleted += FileSystemEventHandler;
                watcher.Renamed += FileSystemEventHandler;

                fileSystemWatchers.Add(driveLetter.ToString(), watcher);
            }
        }

        private static void StopFileSystemWatcher()
        {
            if (fileSystemWatchers != null)
            {
                foreach (var watcher in fileSystemWatchers)
                {
                    watcher.Value.EnableRaisingEvents = false;
                    watcher.Value.Dispose();
                }

                fileSystemWatchers.Clear();
                fileSystemWatchers = null;
            }
        }

        private static void VolumeWatcher_DriveInserted(object sender, VolumeChangedEventArgs e)
        {
            try
            {
                var watcher = fileSystemWatchers[e.DriveLetter];

                if (watcher.Path == string.Empty)
                {
                    watcher.Path = string.Format(@"{0}:\", e.DriveLetter);
                }

                watcher.EnableRaisingEvents = true;
            }
            catch (Exception ex)
            {
                #if DEBUG
                    Debug.WriteLine(ex);
                    
                    if (ex.InnerException != null)
                        Debug.WriteLine(ex.InnerException);
                #else
                    Trace.WriteLine(ex);

                    if (ex.InnerException != null)
                        Trace.WriteLine(ex.InnerException);
                #endif
            }
        }

        private static void VolumeWatcher_DriveRemoved(object sender, VolumeChangedEventArgs e)
        {
            try
            {
                var fileSystemWatcher = fileSystemWatchers[e.DriveLetter];

                fileSystemWatcher.EnableRaisingEvents = false;
            }
            catch (Exception ex)
            {
                #if DEBUG
                    Debug.WriteLine(ex);
                    
                    if (ex.InnerException != null)
                        Debug.WriteLine(ex.InnerException);
                #else
                    Trace.WriteLine(ex);

                    if (ex.InnerException != null)
                        Trace.WriteLine(ex.InnerException);
                #endif
            }
        }

        private static void FileSystemEventHandler(object sender, FileSystemEventArgs e)
        {
            var originalEvent = e.ChangeType.ToSuncatLogEvent();

            // Only log files, not directories (don't check for files exist on delete)
            if (originalEvent == SuncatLogEvent.DeleteFile || System.IO.File.Exists(e.FullPath))
            {
                var eArgs = e as SmartFileSystemEventArgs;
                var reArgs = e as SmartFileSystemRenamedEventArgs;

                var te = new TrackEventArgs();
                var log = new SuncatLog();
                
                log.Event = originalEvent;
                log.DateTime = DateTime.Now;
                
                switch (originalEvent)
                {
                    case SuncatLogEvent.CreateFile:
                    case SuncatLogEvent.ChangeFile:
                        {
                            string copiedFile = null;
                            var isCopiedFile = false;

                            if (HookActivityMonitor.LastCopiedFiles != null && HookActivityMonitor.LastCopiedFiles.Count > 0)
                            {
                                var currentFileInfo = new FileInfo(e.FullPath);

                                foreach (var copiedFileInfo in HookActivityMonitor.LastCopiedFiles.OrderByDescending(f => f.FileInfo.Name.Length))
                                {
                                    if (copiedFileInfo.IsDirectory)
                                    {
                                        copiedFile = FindCopiedFile(new DirectoryInfo(copiedFileInfo.FileInfo.FullName), currentFileInfo, currentFileInfo.Directory);

                                        if (copiedFile != null)
                                        {
                                            isCopiedFile = true;
                                            break;
                                        }
                                    }
                                    else
                                    {
                                        if (currentFileInfo.Name.StartsWith(Path.GetFileNameWithoutExtension(copiedFileInfo.FileInfo.Name)))
                                        {
                                            isCopiedFile = true;
                                            copiedFile = copiedFileInfo.FileInfo.FullName;
                                            break;
                                        }
                                    }
                                }
                            }

                            if (isCopiedFile)
                            {
                                log.Event = SuncatLogEvent.CopyFile;
                                log.Data1 = copiedFile;
                                log.Data2 = eArgs.FullPath;

                                try
                                {
                                    log.Data3 = volumeWatcher.DriveList[copiedFile.Substring(0, 1)].DriveType.ToString();
                                }
                                catch (Exception ex)
                                {
                                    #if DEBUG
                                        Debug.WriteLine(ex);
                                        
                                        if (ex.InnerException != null)
                                            Debug.WriteLine(ex.InnerException);
                                    #else
                                        Trace.WriteLine(ex);

                                        if (ex.InnerException != null)
                                            Trace.WriteLine(ex.InnerException);
                                    #endif

                                    log.Data3 = copiedFile.StartsWith(@"\\") ? DriveType.Network.ToString() : DriveType.Unknown.ToString();
                                }

                                log.Data3 += ",";
                            }
                            else
                            {
                                log.Data1 = eArgs.FullPath;
                            }
                        }
                        break;

                    case SuncatLogEvent.RenameFile:
                        {
                            var isOldFileTypeAssociated = IsAssociatedFileType(reArgs.OldFullPath);
                            var isNewFileTypeAssociated = IsAssociatedFileType(reArgs.FullPath);

                            // Only 1 operation on the file (rename in this case), treat it as RenameFile, else treat is as ChangeFile
                            if ((isOldFileTypeAssociated && isNewFileTypeAssociated && Path.GetExtension(reArgs.OldFullPath).Equals(Path.GetExtension(reArgs.FullPath), StringComparison.OrdinalIgnoreCase) && reArgs.SameFileEventCount() == 1)
                                 || fileTypeKeyNames == null)
                            {
                                log.Data1 = reArgs.OldFullPath;
                                log.Data2 = reArgs.FullPath;
                            }
                            else if (!isOldFileTypeAssociated && isNewFileTypeAssociated)
                            {
                                log.Event = SuncatLogEvent.ChangeFile;
                                log.Data1 = reArgs.FullPath;
                            }
                            else // Discard rename event if the file extension is unknown
                            {
                                log.Event = SuncatLogEvent.None;
                            }
                        }
                        break;

                    default:
                        {
                            log.Data1 = eArgs.FullPath;
                        }
                        break;
                }

                if (log.Event != SuncatLogEvent.None)
                {
                    try
                    {
                        if (originalEvent == SuncatLogEvent.RenameFile)
                        {
                            log.Data3 += volumeWatcher.DriveList[reArgs.FullPath.Substring(0, 1)].DriveType.ToString();
                        }
                        else
                        {
                            log.Data3 += volumeWatcher.DriveList[eArgs.FullPath.Substring(0, 1)].DriveType.ToString();
                        }
                    }
                    catch (Exception ex)
                    {
                        #if DEBUG
                            Debug.WriteLine(ex);
                            
                            if (ex.InnerException != null)
                                Debug.WriteLine(ex.InnerException);
                        #else
                            Trace.WriteLine(ex);
                            
                            if (ex.InnerException != null)
                                Trace.WriteLine(ex.InnerException);
                        #endif
                    }

                    te.LogData = log;

                    // The ignore part of these events are processed in the custom FileSystemWatcher's core
                    switch (log.Event)
                    {
                        case SuncatLogEvent.CreateFile:
                        case SuncatLogEvent.ChangeFile:
                        case SuncatLogEvent.RenameFile:
                            {
                                Track?.Invoke(null, te);
                            }
                            break;

                        case SuncatLogEvent.DeleteFile:
                            {
                                // Discard deleted event if the same file is renamed right after
                                if (!eArgs.HasEvent(WatcherChangeTypes.Renamed))
                                {
                                    Track?.Invoke(null, te);
                                }
                            }
                            break;

                        default:
                            {
                                if (!IgnoreEventCallback(log))
                                {
                                    Track?.Invoke(null, te);
                                }
                            }
                            break;
                    }
                }
            }
        }

        private static string FindCopiedFile(DirectoryInfo baseDirectory, FileInfo currentFile, DirectoryInfo currentFileDirectory)
        {
            if (currentFileDirectory.Name.StartsWith(baseDirectory.Name))
            {
                return Path.Combine(baseDirectory.FullName, currentFile.FullName.Replace(currentFileDirectory.FullName, string.Empty).TrimStart(Path.DirectorySeparatorChar));
            }
            else if (currentFileDirectory.Parent != null)
            {
                return FindCopiedFile(baseDirectory, currentFile, currentFileDirectory.Parent);
            }
            else
            {
                return null;
            }
        }

        public static void Start()
        {
            StartFileSystemWatcher();

            volumeWatcher = new VolumeWatcher();
            volumeWatcher.DriveInserted += VolumeWatcher_DriveInserted;
            volumeWatcher.DriveRemoved += VolumeWatcher_DriveRemoved;
            volumeWatcher.Start();

            recentFilesChecker = new Thread(() => CheckUserRecentFiles());
            recentFilesChecker.IsBackground = true;
            recentFilesChecker.Start();
        }

        public static void Stop()
        {
            if (volumeWatcher != null)
            {
                volumeWatcher.DriveInserted -= VolumeWatcher_DriveInserted;
                volumeWatcher.DriveRemoved -= VolumeWatcher_DriveRemoved;
                volumeWatcher = null;
            }

            StopFileSystemWatcher();
        }
    }
}
