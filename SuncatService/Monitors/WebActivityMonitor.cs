using PListNet;
using PListNet.Nodes;
using SuncatCommon;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Threading;

namespace SuncatService.Monitors
{
    public static class WebActivityMonitor
    {
        private static readonly string serviceName = new ProjectInstaller().ServiceInstaller.ServiceName;
        private static readonly string rootDrive = Path.GetPathRoot(Environment.SystemDirectory);
        private static readonly string serviceAppData = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), serviceName);
        private static readonly string localDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        private static Dictionary<string, string> lastActiveUsernames = new Dictionary<string, string>();
        private static Thread firefoxHistoryChecker;
        private static Thread chromeHistoryChecker;
        private static Thread vivaldiHistoryChecker;
        private static Thread operaHistoryChecker;
        private static Thread safariHistoryChecker;
        private static Thread paleMoonHistoryChecker;
        private static Thread seaMonkeyHistoryChecker;
        private static Thread canaryHistoryChecker;
        private static Thread chromiumHistoryChecker;
        private static Thread yandexHistoryChecker;
        private static Thread edgeHistoryChecker;

        public static event TrackEventHandler Track;

        private static string GetFirefoxDefaultProfile(string dataDir)
        {
            string profile = null;

            try
            {
                var profilesFile = $@"{dataDir}\profiles.ini";

                if (File.Exists(profilesFile))
                {
                    var reader = new StreamReader(profilesFile);
                    var content = reader.ReadToEnd();
                    var lines = content.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                    profile = lines.First(l => l.Contains("Path=")).Split(new string[] { "=", "/" }, StringSplitOptions.None)[2];
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

            return profile;
        }

        private static void CheckFirefoxHistory(string dataDir, string tempFileName)
        {
            var lastDatabaseMD5Hash = default(string);
            DateTime? firstUrlTime = null, lastUrlTime = null;
            var newUrls = new List<Tuple<string, string>>();

            while (true)
            {
                try
                {
                    var session = SuncatUtilities.GetActiveSession();

                    if (session != null && !string.IsNullOrEmpty(session.UserName))
                    {
                        if (!lastActiveUsernames.ContainsKey(tempFileName) || lastActiveUsernames[tempFileName] == session.UserName)
                        {
                            var realDataDir = dataDir.Replace("[USERNAME]", session.UserName);
                            var profile = GetFirefoxDefaultProfile(realDataDir);

                            if (profile != null)
                            {
                                var database = $@"{realDataDir}\Profiles\{profile}\places.sqlite";

                                if (File.Exists(database) && File.Exists($"{database}-wal"))
                                {
                                    var databaseMD5Hash = SuncatUtilities.GetMD5HashFromFile($"{database}-wal");

                                    if (databaseMD5Hash != lastDatabaseMD5Hash)
                                    {
                                        Directory.CreateDirectory($@"{serviceAppData}\History");
                                        File.Copy(database, $@"{serviceAppData}\History\{tempFileName}", true);
                                        File.Copy($"{database}-wal", $@"{serviceAppData}\History\{tempFileName}-wal", true);

                                        using (var connection = new SQLiteConnection($@"Data Source={serviceAppData}\History\{tempFileName};PRAGMA journal_mode=WAL"))
                                        {
                                            connection.Open();

                                            using (var command = new SQLiteCommand("select mp.url, mp.title, strftime('%Y-%m-%d %H:%M:%f', substr(mhv.visit_date, 1, 17) / 1000000.0, 'unixepoch', 'localtime') as visit_date from moz_places mp join moz_historyvisits mhv on mhv.place_id = mp.id where mp.hidden = 0 order by visit_date desc, mhv.id desc", connection))
                                            {
                                                using (var reader = command.ExecuteReader())
                                                {
                                                    newUrls.Clear();

                                                    if (reader.HasRows)
                                                    {
                                                        string lastUrl = null;

                                                        while (reader.Read())
                                                        {
                                                            var url = Convert.ToString(reader["url"]);
                                                            var title = Convert.ToString(reader["title"]);
                                                            var urlTime = Convert.ToDateTime(reader["visit_date"]);

                                                            if (url != lastUrl)
                                                            {
                                                                if (lastUrlTime.HasValue)
                                                                {
                                                                    if (urlTime >= lastUrlTime)
                                                                    {
                                                                        if (newUrls.Count == 0)
                                                                        {
                                                                            firstUrlTime = urlTime;
                                                                        }

                                                                        if (urlTime > lastUrlTime && url.StartsWith("http"))
                                                                        {
                                                                            var tuple = new Tuple<string, string>(url, string.IsNullOrWhiteSpace(title) ? null : title.Trim());

                                                                            newUrls.Insert(0, tuple);
                                                                        }
                                                                    }
                                                                    else
                                                                    {
                                                                        lastUrlTime = firstUrlTime;
                                                                        break;
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    lastUrlTime = urlTime;
                                                                    break;
                                                                }
                                                            }

                                                            lastUrl = url;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        lastUrlTime = default(DateTime);
                                                    }

                                                    foreach (var url in newUrls)
                                                    {
                                                        var te = new TrackEventArgs();
                                                        var log = new SuncatLog();

                                                        log.DateTime = DateTime.Now;
                                                        log.Event = SuncatLogEvent.OpenURL;
                                                        log.Data1 = url.Item1;
                                                        log.Data2 = url.Item2;
                                                        log.Data3 = tempFileName;
                                                        te.LogData = log;

                                                        Track?.Invoke(null, te);
                                                    }

                                                    if (newUrls.Count != 0)
                                                    {
                                                        lastUrlTime = firstUrlTime;
                                                    }
                                                }
                                            }
                                        }

                                        lastDatabaseMD5Hash = databaseMD5Hash;
                                    }
                                }
                            }
                            else
                            {
                                lastUrlTime = default(DateTime);
                            }
                        }
                        else
                        {
                            lastUrlTime = null;
                        }

                        lastActiveUsernames[tempFileName] = session.UserName;
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

        private static void CheckChromeHistory(string dataDir, string tempFileName)
        {
            var lastDatabaseMD5Hash = default(string);
            DateTime? firstUrlTime = null, lastUrlTime = null;
            var newUrls = new List<Tuple<string, string>>();

            while (true)
            {
                try
                {
                    var session = SuncatUtilities.GetActiveSession();

                    if (session != null && !string.IsNullOrEmpty(session.UserName))
                    {
                        if (!lastActiveUsernames.ContainsKey(tempFileName) || lastActiveUsernames[tempFileName] == session.UserName)
                        {
                            var realDataDir = dataDir.Replace("[USERNAME]", session.UserName);
                            var database = $@"{realDataDir}\History";

                            if (File.Exists(database))
                            {
                                var databaseMD5Hash = SuncatUtilities.GetMD5HashFromFile(database);

                                if (databaseMD5Hash != lastDatabaseMD5Hash)
                                {
                                    Directory.CreateDirectory($@"{serviceAppData}\History");
                                    File.Copy(database, $@"{serviceAppData}\History\{tempFileName}", true);

                                    using (var connection = new SQLiteConnection($@"Data Source={serviceAppData}\History\{tempFileName}"))
                                    {
                                        connection.Open();

                                        using (var command = new SQLiteCommand("select u.url, u.title, strftime('%Y-%m-%d %H:%M:%f', substr(v.visit_time, 1, 17) / 1000000.0 - 11644473600, 'unixepoch', 'localtime') as visit_time from urls u join visits v on v.url = u.id where u.hidden = 0 order by visit_time desc, v.id desc", connection))
                                        {
                                            using (var reader = command.ExecuteReader())
                                            {
                                                newUrls.Clear();

                                                if (reader.HasRows)
                                                {
                                                    string lastUrl = null;

                                                    while (reader.Read())
                                                    {
                                                        var url = Convert.ToString(reader["url"]);
                                                        var title = Convert.ToString(reader["title"]);
                                                        var urlTime = Convert.ToDateTime(reader["visit_time"]);

                                                        if (url != lastUrl)
                                                        {
                                                            if (lastUrlTime.HasValue)
                                                            {
                                                                if (urlTime >= lastUrlTime)
                                                                {
                                                                    if (newUrls.Count == 0)
                                                                    {
                                                                        firstUrlTime = urlTime;
                                                                    }

                                                                    if (urlTime > lastUrlTime && url.StartsWith("http"))
                                                                    {
                                                                        var tuple = new Tuple<string, string>(url, string.IsNullOrWhiteSpace(title) ? null : title.Trim());

                                                                        newUrls.Insert(0, tuple);
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    lastUrlTime = firstUrlTime;
                                                                    break;
                                                                }
                                                            }
                                                            else
                                                            {
                                                                lastUrlTime = urlTime;
                                                                break;
                                                            }
                                                        }

                                                        lastUrl = url;
                                                    }
                                                }
                                                else
                                                {
                                                    lastUrlTime = default(DateTime);
                                                }

                                                foreach (var url in newUrls)
                                                {
                                                    var te = new TrackEventArgs();
                                                    var log = new SuncatLog();

                                                    log.DateTime = DateTime.Now;
                                                    log.Event = SuncatLogEvent.OpenURL;
                                                    log.Data1 = url.Item1;
                                                    log.Data2 = url.Item2;
                                                    log.Data3 = tempFileName;
                                                    te.LogData = log;

                                                    Track?.Invoke(null, te);
                                                }

                                                if (newUrls.Count != 0)
                                                {
                                                    lastUrlTime = firstUrlTime;
                                                }
                                            }
                                        }
                                    }

                                    lastDatabaseMD5Hash = databaseMD5Hash;
                                }
                            }
                            else
                            {
                                lastUrlTime = default(DateTime);
                            }
                        }
                        else
                        {
                            lastUrlTime = null;
                        }

                        lastActiveUsernames[tempFileName] = session.UserName;
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

        private static void CheckSafariWinHistory(string dataDir, string tempFileName)
        {
            var lastDatabaseMD5Hash = default(string);
            Double? firstUrlTime = null, lastUrlTime = null;
            var newUrls = new List<Tuple<string, string>>();

            while (true)
            {
                try
                {
                    var session = SuncatUtilities.GetActiveSession();

                    if (session != null && !string.IsNullOrEmpty(session.UserName))
                    {
                        if (!lastActiveUsernames.ContainsKey(tempFileName) || lastActiveUsernames[tempFileName] == session.UserName)
                        {
                            var realDataDir = dataDir.Replace("[USERNAME]", session.UserName);
                            var database = $@"{realDataDir}\History.plist";

                            if (File.Exists(database))
                            {
                                var databaseMD5Hash = SuncatUtilities.GetMD5HashFromFile(database);

                                if (databaseMD5Hash != lastDatabaseMD5Hash)
                                {
                                    Directory.CreateDirectory($@"{serviceAppData}\History");
                                    File.Copy(database, $@"{serviceAppData}\History\{tempFileName}", true);

                                    var rootNode = (DictionaryNode)PList.Load(File.Open($@"{serviceAppData}\History\{tempFileName}", FileMode.Open, FileAccess.Read, FileShare.ReadWrite));
                                    var urlsNode = (ArrayNode)rootNode["WebHistoryDates"];

                                    newUrls.Clear();

                                    if (urlsNode.Count > 0)
                                    {
                                        string lastUrl = null;

                                        foreach (var urlNode in urlsNode.OfType<DictionaryNode>().Select(
                                            u => new
                                            {
                                                Url = ((StringNode)u[""]).Value,
                                                Title = (u.ContainsKey("title") ? ((StringNode)u["title"]).Value : string.Empty),
                                                LastVisitedDate = Convert.ToDouble(((StringNode)u["lastVisitedDate"]).Value, CultureInfo.InvariantCulture),
                                            }).OrderByDescending(u => u.LastVisitedDate))
                                        {
                                            var url = urlNode.Url;
                                            var title = urlNode.Title;
                                            var urlTime = urlNode.LastVisitedDate;

                                            if (url != lastUrl)
                                            {
                                                if (lastUrlTime.HasValue)
                                                {
                                                    if (urlTime >= lastUrlTime)
                                                    {
                                                        if (newUrls.Count == 0)
                                                        {
                                                            firstUrlTime = urlTime;
                                                        }

                                                        if (urlTime > lastUrlTime && url.StartsWith("http"))
                                                        {
                                                            var tuple = new Tuple<string, string>(url, string.IsNullOrWhiteSpace(title) ? null : title.Trim());

                                                            newUrls.Insert(0, tuple);
                                                        }
                                                    }
                                                    else
                                                    {
                                                        lastUrlTime = firstUrlTime;
                                                        break;
                                                    }
                                                }
                                                else
                                                {
                                                    lastUrlTime = urlTime;
                                                    break;
                                                }
                                            }

                                            lastUrl = url;
                                        }
                                    }
                                    else
                                    {
                                        lastUrlTime = default(Double);
                                    }

                                    foreach (var url in newUrls)
                                    {
                                        var te = new TrackEventArgs();
                                        var log = new SuncatLog();

                                        log.DateTime = DateTime.Now;
                                        log.Event = SuncatLogEvent.OpenURL;
                                        log.Data1 = url.Item1;
                                        log.Data2 = url.Item2;
                                        log.Data3 = tempFileName;
                                        te.LogData = log;

                                        Track?.Invoke(null, te);
                                    }

                                    if (newUrls.Count != 0)
                                    {
                                        lastUrlTime = firstUrlTime;
                                    }

                                    lastDatabaseMD5Hash = databaseMD5Hash;
                                }
                            }
                            else
                            {
                                lastUrlTime = default(Double);
                            }
                        }
                        else
                        {
                            lastUrlTime = null;
                        }

                        lastActiveUsernames[tempFileName] = session.UserName;
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

        private static void CheckSafariMacHistory(string dataDir, string tempFileName)
        {
            // ~/Library/Safari/History.db
            // select hi.url, hv.title, strftime('%Y-%m-%d %H:%M:%f', substr(hv.visit_time, 1, 17) + 978307200, 'unixepoch', 'localtime') as visit_time from history_items hi join history_visits hv on hv.history_item = hi.id order by visit_time desc, hv.id desc

            throw new NotImplementedException();
        }

        private static void EnableFirefoxProxy(string dataDir, bool enabled)
        {
            try
            {
                var profile = GetFirefoxDefaultProfile(dataDir);

                if (profile != null)
                {
                    var preferencesFile = $@"{dataDir}\Profiles\{profile}\prefs.js";

                    if (File.Exists(preferencesFile))
                    {
                        var linesToKeep = File.ReadAllLines(preferencesFile).Where(l => !l.Contains("\"network.proxy.type\"")).ToList();

                        if (!enabled)
                        {
                            linesToKeep.Add("user_pref(\"network.proxy.type\", 0);");
                        }

                        File.WriteAllLines(preferencesFile, linesToKeep);
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

        public static void Start()
        {
            // Proxy is still a bit unstable after a long uptime, seems to not respond to in browsers and is eating a lot of memory after a few hours up.
            //if (false)
            //{
            //    EnableFirefoxProxy($@"{rootDrive}Users\{session.UserName}\AppData\Roaming\Mozilla\Firefox", false);
            //    EnableFirefoxProxy($@"{rootDrive}Users\{session.UserName}\AppData\Roaming\Moonchild Productions\Pale Moon", false);
            //    EnableFirefoxProxy($@"{rootDrive}Users\{session.UserName}\AppData\Roaming\Mozilla\SeaMonkey", false);
            //    StartProxyServer();
            //}

            firefoxHistoryChecker = new Thread(() => CheckFirefoxHistory($@"{rootDrive}Users\[USERNAME]\AppData\Roaming\Mozilla\Firefox", "Firefox"));
            firefoxHistoryChecker.IsBackground = true;
            firefoxHistoryChecker.Start();

            chromeHistoryChecker = new Thread(() => CheckChromeHistory($@"{rootDrive}Users\[USERNAME]\AppData\Local\Google\Chrome\User Data\Default", "Chrome"));
            chromeHistoryChecker.IsBackground = true;
            chromeHistoryChecker.Start();

            vivaldiHistoryChecker = new Thread(() => CheckChromeHistory($@"{rootDrive}Users\[USERNAME]\AppData\Local\Vivaldi\User Data\Default", "Vivaldi"));
            vivaldiHistoryChecker.IsBackground = true;
            vivaldiHistoryChecker.Start();

            operaHistoryChecker = new Thread(() => CheckChromeHistory($@"{rootDrive}Users\[USERNAME]\AppData\Roaming\Opera Software\Opera Stable", "Opera"));
            operaHistoryChecker.IsBackground = true;
            operaHistoryChecker.Start();

            safariHistoryChecker = new Thread(() => CheckSafariWinHistory($@"{rootDrive}Users\[USERNAME]\AppData\Roaming\Apple Computer\Safari", "Safari"));
            safariHistoryChecker.IsBackground = true;
            safariHistoryChecker.Start();

            paleMoonHistoryChecker = new Thread(() => CheckFirefoxHistory($@"{rootDrive}Users\[USERNAME]\AppData\Roaming\Moonchild Productions\Pale Moon", "PaleMoon"));
            paleMoonHistoryChecker.IsBackground = true;
            paleMoonHistoryChecker.Start();

            seaMonkeyHistoryChecker = new Thread(() => CheckFirefoxHistory($@"{rootDrive}Users\[USERNAME]\AppData\Roaming\Mozilla\SeaMonkey", "SeaMonkey"));
            seaMonkeyHistoryChecker.IsBackground = true;
            seaMonkeyHistoryChecker.Start();

            canaryHistoryChecker = new Thread(() => CheckChromeHistory($@"{rootDrive}Users\[USERNAME]\AppData\Local\Google\Chrome SxS\User Data\Default", "Canary"));
            canaryHistoryChecker.IsBackground = true;
            canaryHistoryChecker.Start();

            chromiumHistoryChecker = new Thread(() => CheckChromeHistory($@"{rootDrive}Users\[USERNAME]\AppData\Local\Chromium\User Data\Default", "Chromium"));
            chromiumHistoryChecker.IsBackground = true;
            chromiumHistoryChecker.Start();

            yandexHistoryChecker = new Thread(() => CheckChromeHistory($@"{rootDrive}Users\[USERNAME]\AppData\Local\Yandex\YandexBrowser\User Data\Default", "Yandex"));
            yandexHistoryChecker.IsBackground = true;
            yandexHistoryChecker.Start();

            edgeHistoryChecker = new Thread(() => CheckChromeHistory($@"{rootDrive}Users\[USERNAME]\AppData\Local\Microsoft\Edge\User Data\Default", "Edge"));
            edgeHistoryChecker.IsBackground = true;
            edgeHistoryChecker.Start();
        }

        public static void Stop()
        {
        //        if (proxyServer != null)
        //        {
        //            StopProxyServer();
        //            EnableFirefoxProxy($@"{rootDrive}Users\{session.UserName}\AppData\Roaming\Mozilla\Firefox", true);
        //            EnableFirefoxProxy($@"{rootDrive}Users\{session.UserName}\AppData\Roaming\Moonchild Productions\Pale Moon", true);
        //            EnableFirefoxProxy($@"{rootDrive}Users\{session.UserName}\AppData\Roaming\Mozilla\SeaMonkey", true);
        //        }
        }
    }
}
