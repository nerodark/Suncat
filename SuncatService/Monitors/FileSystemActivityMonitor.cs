using Cassia;
using IWshRuntimeLibrary;
using LeanWork.IO.FileSystem;
using SuncatObjects;
using SuncatService.Monitors.CustomFileSystemWatcher;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reactive.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace SuncatService.Monitors
{
    public static class Extensions
    {
        public static IObservable<T> BufferUntilInactive<T>(this IObservable<T> stream, TimeSpan delay)
        {
            var closes = stream.Throttle(delay);
            return stream.Window(() => closes).SelectMany(x => x.FirstAsync());
        }
    }

    public static class FileSystemActivityMonitor
    {
        private static readonly string serviceName = new ProjectInstaller().ServiceInstaller.ServiceName;
        private static readonly string rootDrive = Path.GetPathRoot(Environment.SystemDirectory);
        private static readonly string serviceAppData = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), serviceName);
        //private static EventWatcher filter;
        private static VolumeWatcher volumeWatcher;
        private static Dictionary<string, RecoveringFileSystemWatcher> fileSystemWatchers;
        private static Thread recentFilesChecker;

        public static event TrackEventHandler Track;

        //private static IObservable<FileSystemEvent>[] CreateSourcesDriverMode(EventWatcher watcher)
        //{
        //    return new[]
        //    {
        //        Observable.FromEventPattern<FileEventHandler, FileSystemEvent>(
        //            h => watcher.OnCreate += h,
        //            h => watcher.OnCreate -= h)
        //            .Select(ev => ev.EventArgs),

        //        Observable.FromEventPattern<FileEventHandler, FileSystemEvent>(
        //            h => watcher.OnDelete += h,
        //            h => watcher.OnDelete -= h)
        //            .Select(ev => ev.EventArgs),

        //        Observable.FromEventPattern<FileEventHandler, FileSystemEvent>(
        //            h => watcher.OnChange += h,
        //            h => watcher.OnChange -= h)
        //            .Select(ev => ev.EventArgs),

        //        Observable.FromEventPattern<MoveEventHandler, RenameOrMoveEvent>(
        //            h => watcher.OnRenameOrMove += h,
        //            h => watcher.OnRenameOrMove -= h)
        //            .Select(ev => ev.EventArgs),
        //    };
        //}

        private static IObservable<FileSystemEventArgs>[] CreateSourcesNormalMode(BufferingFileSystemWatcher watcher)
        {
            return new[]
            {
                Observable.FromEventPattern<FileSystemEventHandler, FileSystemEventArgs>(
                        h => watcher.Created += h,
                        h => watcher.Created -= h)
                        .Select(ev => ev.EventArgs),

                Observable.FromEventPattern<FileSystemEventHandler, FileSystemEventArgs>(
                        h => watcher.Deleted += h,
                        h => watcher.Deleted -= h)
                        .Select(ev => ev.EventArgs),

                Observable.FromEventPattern<FileSystemEventHandler, FileSystemEventArgs>(
                        h => watcher.Changed += h,
                        h => watcher.Changed -= h)
                        .Select(ev => ev.EventArgs),

                Observable.FromEventPattern<RenamedEventHandler, FileSystemEventArgs>(
                        h => watcher.Renamed += h,
                        h => watcher.Renamed -= h)
                        .Select(ev => ev.EventArgs),
            };
        }

        public static bool IsIgnoredPath(string path)
        {
            var ignored = false;
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
            ignored |= path.IndexOf("rempl", StringComparison.OrdinalIgnoreCase) > -1;
            ignored |= path.EndsWith("desktop.ini", StringComparison.OrdinalIgnoreCase);
            ignored |= Path.GetFileName(path).StartsWith("~");
            ignored |= Path.GetFileName(path).EndsWith("~");
            ignored |= Path.GetExtension(path).Equals(".tmp", StringComparison.OrdinalIgnoreCase);
            ignored |= Path.GetExtension(path).Equals(".lnk", StringComparison.OrdinalIgnoreCase);
            ignored |= (manager != null && manager.ActiveConsoleSession != null && !string.IsNullOrEmpty(manager.ActiveConsoleSession.UserName)
                        && path.StartsWith($@"{rootDrive}Users\") && !path.StartsWith($@"{rootDrive}Users\{manager.ActiveConsoleSession.UserName}\"));

            return ignored;
        }

        //private static void StartFileSystemWatcherDriverMode()
        //{
        //    filter = new EventWatcher();

        //    while (true)
        //    {
        //        try
        //        {
        //            filter.Connect();
        //            break;
        //        }
        //        catch (Exception ex)
        //        {
        //            Debug.WriteLine(ex);

        //            if (ex.InnerException != null)
        //                Debug.WriteLine(ex.InnerException);

        //            Thread.Sleep(1000);
        //        }
        //    }

        //    filter.NotWatchProcess(EventWatcher.GetCurrentProcessId());
        //    filter.WatchPath("*");

        //    var watcher = Observable.Using(
        //        () => filter,
        //        w => CreateSourcesDriverMode(w)
        //        .Merge()
        //        .GroupBy(fse => fse.Filename)
        //        .SelectMany(fse => fse.BufferUntilInactive(new TimeSpan(0, 0, 1))));

        //    watcher.Subscribe(fse =>
        //    {
        //        if (fse != null)
        //        {
        //            try
        //            {
        //                var manager = new TerminalServicesManager();

        //                if (manager != null && manager.ActiveConsoleSession != null && !string.IsNullOrEmpty(manager.ActiveConsoleSession.UserName))
        //                {
        //                    var process = Process.GetProcesses().FirstOrDefault(p => p.Id == Convert.ToInt32(fse.ProcessId));

        //                    if (process != null
        //                        && !IsIgnoredPath(fse.Filename)
        //                        && (!(fse is RenameOrMoveEvent) || !IsIgnoredPath(((RenameOrMoveEvent)fse).OldFilename))
        //                        && process.SessionId == manager.ActiveConsoleSession.SessionId
        //                        // Only log files, not directories (don't check for files exist on delete)
        //                        && (fse.Type == CenterDevice.MiniFSWatcher.Types.EventType.Delete || System.IO.File.Exists(fse.Filename)))
        //                    {
        //                        var te = new TrackEventArgs();
        //                        var log = new SuncatLog();

        //                        log.DateTime = DateTime.Now;

        //                        switch (fse.Type)
        //                        {
        //                            case CenterDevice.MiniFSWatcher.Types.EventType.Create: log.Event = SuncatLogEvent.CreateFile; break;
        //                            case CenterDevice.MiniFSWatcher.Types.EventType.Delete: log.Event = SuncatLogEvent.DeleteFile; break;
        //                            case CenterDevice.MiniFSWatcher.Types.EventType.Change: log.Event = SuncatLogEvent.ChangeFile; break;
        //                            case CenterDevice.MiniFSWatcher.Types.EventType.Move: log.Event = SuncatLogEvent.MoveFile; break;
        //                        }

        //                        if (fse.Type == CenterDevice.MiniFSWatcher.Types.EventType.Move)
        //                        {
        //                            var rme = (RenameOrMoveEvent)fse;
        //                            log.Data1 = rme.OldFilename;
        //                            log.Data2 = rme.Filename;
        //                        }
        //                        else
        //                        {
        //                            log.Data1 = fse.Filename;
        //                        }

        //                        try
        //                        {
        //                            log.Data3 = volumeWatcher.DriveList[fse.Filename.Substring(0, 1)].DriveType.ToString();
        //                        }
        //                        catch (Exception ex)
        //                        {
        //                            Debug.WriteLine(ex);

        //                            if (ex.InnerException != null)
        //                                Debug.WriteLine(ex.InnerException);
        //                        }

        //                        te.LogData = log;

        //                        Track?.Invoke(null, te);
        //                    }
        //                }
        //            }
        //            catch (Exception ex)
        //            {
        //                Debug.WriteLine(ex);

        //                if (ex.InnerException != null)
        //                    Debug.WriteLine(ex.InnerException);
        //            }
        //        }
        //    },
        //    ex =>
        //    {
        //        Debug.WriteLine(ex);

        //        if (ex.InnerException != null)
        //            Debug.WriteLine(ex.InnerException);
        //    });
        //}

        //private static void StopFileSystemWatcherDriverMode()
        //{
        //    try
        //    {
        //        if (filter != null)
        //        {
        //            using (filter)
        //            {
        //                filter.Disconnect();
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Debug.WriteLine(ex);

        //        if (ex.InnerException != null)
        //            Debug.WriteLine(ex.InnerException);
        //    }
        //}

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

                                        if (recentFileDate > lastRecentFileDate
                                            && !IsIgnoredPath(link.TargetPath)
                                            && System.IO.File.Exists(link.TargetPath))
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
                                Track?.Invoke(null, te);
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

        private static void StartFileSystemWatcherNormalMode()
        {
            fileSystemWatchers = new Dictionary<string, RecoveringFileSystemWatcher>();

            for (var i = 0; i < 26; i++)
            {
                var driveLetter = (char)(i + 'A');

                var watcher = new RecoveringFileSystemWatcher();
                watcher.IncludeSubdirectories = true;
                // To enable LastAccess filter: fsutil behavior set DisableLastAccess 0 (and reboot?)
                watcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.DirectoryName | NotifyFilters.LastAccess | NotifyFilters.LastWrite | NotifyFilters.CreationTime | NotifyFilters.Size | NotifyFilters.Security | NotifyFilters.Attributes;

                watcher.Created += FileSystemEventHandler;
                watcher.Changed += FileSystemEventHandler;
                watcher.Deleted += FileSystemEventHandler;
                watcher.Renamed += FileSystemEventHandler;

                fileSystemWatchers.Add(driveLetter.ToString(), watcher);
            }
        }

        private static void StopFileSystemWatcherNormalMode()
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
            // Only log files, not directories (don't check for files exist on delete)
            if (e.ChangeType == WatcherChangeTypes.Deleted || System.IO.File.Exists(e.FullPath))
            {
                var te = new TrackEventArgs();
                var log = new SuncatLog();

                log.DateTime = DateTime.Now;

                switch (e.ChangeType)
                {
                    case WatcherChangeTypes.Created: log.Event = SuncatLogEvent.CreateFile; break;
                    case WatcherChangeTypes.Deleted: log.Event = SuncatLogEvent.DeleteFile; break;
                    case WatcherChangeTypes.Changed: log.Event = SuncatLogEvent.ChangeFile; break;
                    case WatcherChangeTypes.Renamed: log.Event = SuncatLogEvent.MoveFile; break;
                }
            
                if (e.ChangeType == WatcherChangeTypes.Renamed)
                {
                    var rea = (RenamedEventArgs)e;
                    log.Data1 = rea.OldFullPath;
                    log.Data2 = rea.FullPath;
                }
                else
                {
                    switch (e.ChangeType)
                    {
                        case WatcherChangeTypes.Created:
                        case WatcherChangeTypes.Changed:
                            {
                                string copiedFile = null;
                                var isCopiedFile = false;

                                if (HookActivityMonitor.LastCopiedFiles != null && HookActivityMonitor.LastCopiedFiles.Count > 0)
                                {
                                    var currentFileInfo = new FileInfo(e.FullPath);

                                    foreach (var copiedFileInfo in HookActivityMonitor.LastCopiedFiles)
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
                                            if (currentFileInfo.Name == copiedFileInfo.FileInfo.Name)
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
                                    log.Data2 = e.FullPath;

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
                                    log.Data1 = e.FullPath;
                                }
                            }
                            break;

                        default:
                            {
                                log.Data1 = e.FullPath;
                            }
                            break;
                    }
                }

                try
                {
                    log.Data3 += volumeWatcher.DriveList[e.FullPath.Substring(0, 1)].DriveType.ToString();
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

                Track?.Invoke(null, te);
            }
        }

        private static string FindCopiedFile(DirectoryInfo baseDirectory, FileInfo currentFile, DirectoryInfo currentFileDirectory)
        {
            if (currentFileDirectory.Name == baseDirectory.Name)
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
            //StartFileSystemWatcherDriverMode();
            StartFileSystemWatcherNormalMode();

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

            //StopFileSystemWatcherDriverMode();
            StopFileSystemWatcherNormalMode();
        }
    }
}
